# excel_line — Fix: make memory actually retrievable (write is not enough)

## Problem
The original plugin logged every `add` to a file and relied on a free-model
sub-agent to *classify + store* it. Two failure modes made it effectively
**write-only** (data saved but never retrievable):

1. The free model was unreachable from the plugin context
   (`No module named 'agent'`), so `_classify` returned `None`.
2. `worker.process_logs` **deleted the log even when classification failed**,
   silently dropping the memory. The master index never received the row, so
   `search` / `prefetch` returned nothing — exactly the "saved but unusable"
   symptom the user reported.

## Fix (round 1 — core durability)
- **`__init__.py` — direct-store mode** in `_handle_add`: when the agent passes
  explicit `zone` + `brief`/`content`, the record is written **straight to the
  Excel store** via `store.add`, with no LLM / indexer dependency. This is the
  durable path that behaves like built-in memory.
- **`worker.py` — no silent drop**: a log is now kept (renamed back) when
  classification fails, so it is retried instead of lost.

## Fix (round 2 — hardening, found via QA loop)
- **`store.py` — master index gains a `title` column**; `search_index` now
  scans `zone + brief + title + path + tags` (was brief-only), so retrieval by
  skill/entity name works even when the name is only in the title.
- **`worker.py` — raw-text fallback**: when the classifier is unavailable the
  worker stores a raw-text backup into `knowledge` (tag `auto-backup`) instead
  of emitting an empty record — memory is never silently lost.
- **`store.py` — exception safety**: `add()` is wrapped so a transient file/lock
  error returns `-1` and logs, never crashes the provider.
- **`store.py` — cross-process file lock**: an atomic lock file serializes
  writes across multiple Hermes processes sharing one `root` (in-process RLock
  alone was insufficient), preventing master-index corruption.

## Verification
`tests/test_excel_line.py` — 30 tests, all green:
- store add / search (unicode, Vietnamese, case-insensitive, **by title**, limit) / read / count / zones
- concurrent writes (5×20 threads, no corruption)
- concurrent **direct-store + worker** on the same store (no corruption)
- **cross-process** two `ExcelLineStore` instances, 60 rows, no loss
- worker: keeps log on classify-fail, raw-backup on fail, deletes on success
- provider: direct-store returns `stored`, retrievable, `prefetch` matches,
  `system_prompt_block` active; direct-store works **with LLM broken**
- legacy sequence mode still logs
- `add` returns `-1` (no crash) on file failure

## Fix (round 3 — Gemini external QA)
Gemini reviewed the code and flagged 7 points; 3 were real gaps, now fixed:
- **Lazy index**: `prefetch()` (auto-retrieve hook) now drains any pending
  seq-logs before searching, so memory is never stranded on disk if the
  background indexer crashed or never ran.
- **`on_session_end` no data loss**: when the LLM classifier fails,
  `_auto_extract` already falls back to `_fallback_store` (raw turns into the
  `knowledge` zone) — data is never silently dropped. Added an explicit
  raw-transcript backup path as defence in depth.
- **Cross-process lock hardened**: the lock file now records `<pid>:<ts>`; a
  stale lock is only cleared after confirming the recorded pid is dead
  (`os.kill(pid, 0)`), eliminating the stale-lock race Gemini identified.

## Verification
`tests/test_excel_line.py` — **33 tests, all green** (added: lazy-index +
prefetch finds drained log, on_session_end keeps data on LLM failure,
stale-lock pid-check).

## Remaining trade-offs (acceptable, noted)
- `search_index` scans `brief`/`title` (not `content`) — agents put the
  entity name in `brief`/`title`; scanning every zone file would be slower.
- `prefetch` returns the first 3 matched details (bounded context).
- Master index uses openpyxl; fine for typical sizes (tens of thousands of
  rows). For very large stores, SQLite would be faster (future enhancement).

## How to test
```
cd plugins/excel_line
uv run --with openpyxl python tests/test_excel_line.py
```
