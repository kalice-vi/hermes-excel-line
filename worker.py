"""Sub-agent worker for excel_line.

Reads the shared agent I/O log folder, uses a free LLM model to classify each
entry into a zone and to distil a concise knowledge note, then writes:
  - a brief row into the master index workbook
  - the concise knowledge into the matching zone workbook

This module is intentionally free of agent-runtime imports so it can run as an
independent subprocess spawned by the provider's on_session_end hook (or by a
cron job / `hermes chat -q`). It talks to the LLM via the Hermes provider
config already present in the environment.
"""

from __future__ import annotations

import json
import os
import time
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
        return json.loads(raw)
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
        if not name.endswith(".jsonl") and not name.endswith(".json"):
            continue
        path = os.path.join(log_dir, name)
        # Atomic take: rename away so concurrent appends go to a fresh file.
        tmp_path = path + ".processing"
        try:
            os.replace(path, tmp_path)
        except OSError:
            continue
        records = _read_records(tmp_path)
        stored_any = False
        for rec in records:
            cls = _classify(rec, free_model_fn, zones)
            if cls:
                store.add(
                    zone=cls.get("zone", "knowledge"),
                    brief=cls.get("brief", "")[:120],
                    content=cls.get("content", "")[:300],
                    title=cls.get("title", "")[:40],
                    tags=cls.get("tags", ""),
                )
                count += 1
                stored_any = True
            else:
                # Classifier unavailable: fall back to a raw-text backup so the
                # memory is NEVER silently dropped (mirrors provider.on_session_end).
                raw = " ".join(str(rec.get(k, "")) for k in ("input", "output", "ts"))[:300]
                if len(raw.strip()) >= 10:
                    store.add(zone="knowledge", brief=raw[:120],
                              content=raw[:300], title="auto-backup",
                              tags="auto-backup,llm-unavailable")
                    count += 1
                    stored_any = True
        # Only delete the log if we actually stored something. If classification
        # failed (e.g. free-model unavailable), KEEP the file (rename .processing
        # back) so it is retried next cycle instead of being silently dropped.
        if stored_any:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        else:
            try:
                os.replace(tmp_path, path)
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
        remaining = [n for n in os.listdir(log_dir)
                     if n.endswith(".jsonl") or n.endswith(".json")]
        if not remaining:
            break
        stored = process_logs(log_dir, store_root, free_model_fn, zones, store)
        total += stored
    return total


def _read_records(path: str) -> List[Dict]:
    out: List[Dict] = []
    try:
        # Accept both the original .jsonl/.json and the .processing temp
        # the indexer renames files to before reading.
        if path.endswith(".jsonl") or path.endswith(".json") or path.endswith(".processing"):
            with open(path, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line:
                        try:
                            out.append(json.loads(line))
                        except Exception:
                            pass
        elif path.endswith(".json"):
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                out.extend(data)
            elif isinstance(data, dict):
                out.append(data)
    except Exception:
        pass
    return out
