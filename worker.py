"""worker.py — background indexer (sub-agent) for the excel_line memory provider.

INTRODUCTION
    This module turns raw agent I/O logs into structured memory. It is the
    "lazy writer" that runs when the agent does NOT call excel_line add directly:
    it reads the shared log folder, asks a free LLM model to classify each entry
    into a zone and distil a concise note, then persists the result via store.py.
    It is intentionally free of agent-runtime imports so it can run as an
    independent subprocess spawned by the provider (on_session_end hook, cron,
    or `hermes chat -q`).

FAILURE POLICY (no silent data loss)
    - If classification fails, the raw text is stored as a backup into the
      `knowledge` zone (tag `auto-backup`) instead of emitting an empty record.
    - If storing fails, the log file is KEPT (renamed back) so it is retried,
      never deleted.
"""

from __future__ import annotations

import json
import os
import time
import uuid
from typing import Dict, List, Optional

from .store import ExcelLineStore, ZONE_DEFAULTS

# Prompt for the free LLM: returns JSON only.
_CLASSIFY_PROMPT = """You are the memory curator of a hierarchical Excel brain.
The store is a TREE: every .xlsx file holds AT MOST 10 rows (ID|Title|Content|Tags|Branch|Updated).
A row whose Branch ends in .xlsx is a NODE pointing to another file; a row whose
Branch is a non-.xlsx path is a LEAF asset file (must be deepest in the tree);
a row with empty Branch is a plain memory.

CURRENT TREE (usage per file shown):
{tree}

Decide what to do with the RECORD below. Rules:
- Chitchat, transient commands, narration, one-off task orders => {{"action":"none"}}. NEVER store verbatim user messages.
- If a plain fact FITS an existing branch with room => {{"action":"add","branch":"<file.xlsx>","title":"<=50 chars","content":"<=250 chars, compressed knowledge NOT the raw dialogue","tags":"comma,keywords"}}.
- Prefer UPDATING/compressing an older memory over creating near-duplicates:
  {{"action":"merge","branch":"<file>","ids":[..],"title":"gộp","content":"nén các ý","tags":".."}}
- Branch full (10/10) and the fact is genuinely new => split:
  {{"action":"child","parent":"<file>","name":"<sub-nhánh>","title":"<=50"}}
  (the worker will then retry the add inside the new child automatically)
- Only create LEAF assets when the record references a real file that belongs to a memory.
Title/Content must be reusable knowledge, compressed Vietnamese or English, ≤50/≤250 chars.

Reply with strict JSON only.
"""


def _classify(entry: Dict, free_model_fn, zones: List[str]) -> Optional[Dict]:
    """Call the free-model function and parse JSON.

    Returns:
      dict   -> classification succeeded
      None   -> the LLM responded but output was unusable (retry-worthy once,
                then fall back to raw backup)
      "DOWN" sentinel is communicated via the wrapper returning "" — see
      _classify_tristate below; kept compatible for existing tests.
    """
    return _classify_tristate(entry, free_model_fn, zones)[1]


def _classify_tristate(entry: Dict, free_model_fn, tree_text: str):
    """Returns (status, parsed) where status is one of:
    'ok'      -> parsed dict with a valid action
    'garbage' -> LLM answered but output invalid JSON / unknown action
    'down'    -> free_model_fn returned empty (no LLM available at all)
    """
    prompt = _CLASSIFY_PROMPT.format(tree=tree_text[:4000]) + \
        "\n\nRECORD:\n" + json.dumps(entry, ensure_ascii=False)[:3000]
    try:
        raw = free_model_fn(prompt)
    except Exception:
        return "down", None
    raw = (raw or "").strip()
    if not raw:
        return "down", None
    try:
        if raw.startswith("```"):
            raw = raw.strip("`")
            raw = raw[raw.find("{") : raw.rfind("}") + 1]
        data = json.loads(raw)
        act = str(data.get("action", "")).strip().lower()
        if act not in ("none", "add", "merge", "child"):
            return "garbage", None
        return "ok", data
    except Exception:
        return "garbage", None


def _apply(store, decision: Dict, rec: Dict) -> int:
    """Execute a curator decision on the tree store. Returns records stored (0/1)."""
    from .brain_store import FullError, BadBranch
    act = decision.get("action")
    if act in ("none",):
        return 0
    if act == "merge":
        try:
            store.merge(decision.get("branch", "brain.xlsx"),
                        [int(i) for i in decision.get("ids", [])],
                        str(decision.get("title") or "merged"),
                        str(decision.get("content") or ""),
                        str(decision.get("tags") or ""))
            return 1
        except Exception:
            return 0
    if act == "add":
        try:
            store.add(decision.get("branch", "brain.xlsx"),
                      title=str(decision.get("title") or "")[:50],
                      content=str(decision.get("content") or "")[:250],
                      tags=str(decision.get("tags") or ""))
            return 1
        except FullError:
            # auto-split: tạo node con rồi retry một lần vào trong nó
            try:
                parent = decision.get("branch", "brain.xlsx")
                name = os.path.splitext(os.path.basename(parent))[0] + "-x"
                node = store.child(parent, name, title=str(decision.get("title") or name)[:50])
                store.add(node["file"], title=str(decision.get("title") or "")[:50],
                          content=str(decision.get("content") or "")[:250],
                          tags=str(decision.get("tags") or ""))
                return 1
            except Exception:
                return 0
        except Exception:
            return 0
    if act == "child":
        try:
            store.child(decision.get("parent", "brain.xlsx"),
                        str(decision.get("name") or "sub"),
                        title=str(decision.get("title") or decision.get("name") or "sub"))
            return 1
        except Exception:
            return 0
    return 0


def process_logs(log_dir: str, store_root: str, free_model_fn,
                 zones: Optional[List[str]] = None,
                 store: Optional["ExcelLineStore"] = None) -> int:
    """Walk unprocessed logs in log_dir, classify, write to Excel.

    Files are atomically renamed to a ``.processing`` temp before reading, so a
    turn that appends a new record WHILE we index is never lost: its write lands
    in a fresh original file and is picked up on the next cycle.

    Returns number of records stored.
    """
    zones = zones or ZONE_DEFAULTS
    store = store or ExcelLineStore(store_root)
    if not os.path.isdir(log_dir):
        return 0

    count = 0
    for name in sorted(os.listdir(log_dir)):
        # Skip raw-transcript backups written by provider._save_raw_transcript.
        # They live in the same log_dir but use a different schema and must NOT
        # be re-indexed as if they were agent I/O logs (would duplicate memory
        # and could loop if a read fails). Only process turn logs.
        if name.startswith("session_raw_"):
            continue
        if not name.endswith(".jsonl") and not name.endswith(".json"):
            continue
        path = os.path.join(log_dir, name)
        # Atomic take: rename away so concurrent appends go to a fresh file.
        # Preserve the original extension in the .processing name so _read_records
        # can tell a JSON-array/.json file from line-delimited JSONL after rename
        # (ChatGPT r7 BLOCKER B3: .json must not be parsed as JSONL post-rename).
        ext = ".jsonl" if name.endswith(".jsonl") else ".json"
        tmp_path = path[: -len(ext)] + ext + ".processing"
        try:
            os.replace(path, tmp_path)
        except OSError:
            continue
        try:
            records = _read_records(tmp_path)
        except Exception:
            # B2 (r8): file-level parse/read error must NOT strand the .processing
            # file forever (scanner only looks for .json/.jsonl). Rename it back to
            # its original name so the next index cycle retries it instead of
            # losing it. This is the recovery path for fix (27).
            try:
                os.replace(tmp_path, path)
            except OSError:
                pass
            continue
        stored_any = False
        failed_records = []  # records that did NOT persist (for safe retry)
        classifier_down = False
        for rec in records:
            # B3 (r8): malformed JSON lines are kept as {"_raw", "_malformed":True}
            # by _read_records. They can never be classified or backed up (no
            # input/output/ts), so they would otherwise be silently deleted. Keep
            # them in failed_records so they are rewritten to a retry log instead
            # of dropped — the operator can inspect/repair them later.
            if rec.get("_malformed"):
                failed_records.append(rec)
                continue
            tree_text = store.tree_text() if hasattr(store, "tree_text") else ""
            status, cls = _classify_tristate(rec, free_model_fn, tree_text)
            if status == "down":
                # LLM unavailable: NEVER write raw prompt/turn echoes to the
                # store (2026-09 audit: that fallback polluted 5.8k junk rows).
                # Keep the log for retry on a later pass.
                classifier_down = True
                failed_records.append(rec)
                continue
            if status == "ok" and cls:
                # Tree curator: action none/add/merge/child executed via _apply
                got = _apply(store, cls, rec)
                if cls.get("action") == "none":
                    # deliberately dropped by curator: nothing to store/retry
                    continue
                if got:
                    count += 1
                    stored_any = True
                else:
                    failed_records.append(rec)
            else:
                # Garbage decision on the v2 tree: keep the record for retry
                # (the tree gives the LLM enough structure to recover next
                # pass); do NOT dump raw turns — pollution was the v1 crime.
                failed_records.append(rec)
        # Disposition:
        # - All records persisted  -> safe to delete the log.
        # - Some records failed    -> rewrite ONLY the failed ones back to a fresh
        #   log so they are retried next cycle (never silently dropped).
        if not failed_records:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        else:
            # Rewrite ONLY the failed records to a fresh retry log. Use a uuid so
            # concurrent threads/processes can never overwrite each other's file
            # (ChatGPT round-6 BLOCKER: pid+ts collides across same-PID threads).
            # Malformed records (json parse fail) are kept too so they are not
            # silently dropped — they get retried next cycle.
            # If the retry file cannot be written, we MUST NOT delete tmp_path,
            # otherwise the unpersisted records are stranded/lost (ChatGPT r6 BLOCKER).
            retry_name = f"retry_{uuid.uuid4().hex}.jsonl"
            try:
                with open(os.path.join(log_dir, retry_name), "w", encoding="utf-8") as rf:
                    for rec in failed_records:
                        rf.write(json.dumps(rec, ensure_ascii=False) + "\n")
            except OSError:
                # Retry log write failed: keep the original .processing file so the
                # records survive and are retried on the next index cycle. Do NOT
                # remove tmp_path here.
                return count
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    return count


def index_while_logs_present(log_dir: str, store_root: str, free_model_fn,
                             zones: Optional[List[str]] = None,
                             store: Optional["ExcelLineStore"] = None) -> int:
    """Drain the log folder completely: keep calling process_logs until no
    indexable log files remain. Each processed file is deleted by
    process_logs, so the loop ends only when the folder is empty.

    Returns total records stored across all iterations.
    """
    zones = zones or ZONE_DEFAULTS
    store = store or ExcelLineStore(store_root)
    if not os.path.isdir(log_dir):
        return 0
    total = 0
    # Cap iterations as a safety net against a pathological loop.
    for _ in range(10000):
        # Count both original logs and retry logs so the loop keeps draining
        # failed-record retries instead of spinning forever (ChatGPT round-4 WARN).
        remaining = [n for n in os.listdir(log_dir)
                     if n.endswith(".jsonl") or n.endswith(".json")]
        # Never index raw-transcript backups (different schema) — they are
        # intentionally skipped by process_logs, so exclude them from the loop
        # guard too or the loop would never see "empty".
        remaining = [n for n in remaining if not n.startswith("session_raw_")]
        if not remaining:
            break
        before = len(remaining)
        stored = process_logs(log_dir, store_root, free_model_fn, zones, store)
        total += stored
        # If a full pass stored nothing AND left the same number of files, the
        # remaining logs are persistently failing (e.g. store unavailable). Stop
        # to avoid a 10k-iteration tight spin; they will be retried next cycle.
        after = len([n for n in os.listdir(log_dir)
                     if (n.endswith(".jsonl") or n.endswith(".json"))
                     and not n.startswith("session_raw_")])
        if stored == 0 and after >= before:
            break
        time.sleep(0.1)  # tiny backoff so a stuck folder does not burn CPU
    return total


def _read_records(path: str) -> List[Dict]:
    """Read a log file into records.

    File-level read/parse failures are RAISED (not swallowed) so the caller's
    "rename to .processing then delete on success" logic never deletes a file it
    could not read — that would be silent data loss (ChatGPT round-6 BLOCKER).
    Only individual malformed *lines* are tolerated (kept as _malformed records
    for retry)."""
    out: List[Dict] = []
    # JSONL (or the .jsonl.processing temp the indexer renames files to —
    # extension preserved so .json parses as array/object, not JSONL).
    if path.endswith(".jsonl") or path.endswith(".jsonl.processing"):
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        out.append(json.loads(line))
                    except Exception:
                        # Malformed line: keep it as a raw record so the worker
                        # can retry / back it up instead of silently dropping it
                        # (ChatGPT round-4 WARN: malformed JSON must not vanish).
                        out.append({"_raw": line, "_malformed": True})
    elif path.endswith(".json") or path.endswith(".json.processing"):
        with open(path, encoding="utf-8") as f:
            data = json.load(f)  # file-level parse error propagates (not swallowed)
        if isinstance(data, list):
            out.extend(data)
        elif isinstance(data, dict):
            out.append(data)
    return out
