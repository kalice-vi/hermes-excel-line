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
_CLASSIFY_PROMPT = """You are a memory classifier for an AI agent's long-term store.
Given one agent I/O record (input + output), decide:
1. zone: one of {zones}
2. brief: a 1-sentence human-readable summary (max 120 chars, Vietnamese OK)
3. title: short title (max 40 chars)
4. content: the concise KEY knowledge to keep (max 300 chars; drop chatter,
   keep decisions, facts, preferences, how-tos, contacts)
5. tags: comma-separated keywords

Reply with strict JSON only:
{{"zone":"...","brief":"...","title":"...","content":"...","tags":"..."}}
"""


def _classify(entry: Dict, free_model_fn, zones: List[str]) -> Optional[Dict]:
    """Call the free-model function and parse JSON. Returns dict or None."""
    prompt = _CLASSIFY_PROMPT.format(zones=", ".join(zones)) + \
        "\n\nRECORD:\n" + json.dumps(entry, ensure_ascii=False)[:3000]
    try:
        raw = free_model_fn(prompt)
        # tolerate fenced json
        raw = raw.strip()
        if raw.startswith("```"):
            raw = raw.strip("`")
            raw = raw[raw.find("{") : raw.rfind("}") + 1]
        data = json.loads(raw)
        # Validate zone membership (ChatGPT review WARN): the LLM may return a
        # zone that is not in the whitelist; reject it so we fall back to the
        # raw-text backup instead of creating an arbitrary .xlsx.
        z = str(data.get("zone", "")).strip().lower()
        if z and z not in [s.lower() for s in zones]:
            return None
        return data
    except Exception:
        return None


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
        for rec in records:
            # B3 (r8): malformed JSON lines are kept as {"_raw", "_malformed":True}
            # by _read_records. They can never be classified or backed up (no
            # input/output/ts), so they would otherwise be silently deleted. Keep
            # them in failed_records so they are rewritten to a retry log instead
            # of dropped — the operator can inspect/repair them later.
            if rec.get("_malformed"):
                failed_records.append(rec)
                continue
            cls = _classify(rec, free_model_fn, zones)
            if cls:
                rid = store.add(
                    zone=cls.get("zone", "knowledge"),
                    brief=cls.get("brief", "")[:120],
                    content=cls.get("content", "")[:300],
                    title=cls.get("title", "")[:40],
                    tags=cls.get("tags", ""),
                )
                # ChatGPT review B2: add() returns -1 on failure (it swallows
                # exceptions defensively). Track success vs failure so we never
                # delete a log that still has unpersisted records (mixed-success
                # case: A persists, B fails -> B must be retried, not dropped).
                if rid and rid > 0:
                    count += 1
                    stored_any = True
                else:
                    failed_records.append(rec)
            else:
                # Classifier unavailable: fall back to a raw-text backup so the
                # memory is NEVER silently dropped (mirrors provider.on_session_end).
                raw = " ".join(str(rec.get(k, "")) for k in ("input", "output", "ts"))[:300]
                if len(raw.strip()) >= 10:
                    rid = store.add(zone="knowledge", brief=raw[:120],
                                    content=raw[:300], title="auto-backup",
                                    tags="auto-backup,llm-unavailable")
                    if rid and rid > 0:
                        count += 1
                        stored_any = True
                    else:
                        failed_records.append(rec)
                else:
                    # too short to be meaningful -> skip but do not retry
                    pass
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
