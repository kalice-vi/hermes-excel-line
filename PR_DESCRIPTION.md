# PR / Submission: excel_line — Excel-backed long-term memory provider

> Status: ready to publish. Passes `hermes plugins doctor`. 33/33 QA tests green.
> No user data is shipped (`.gitignore` excludes `*.xlsx`, `logs/`, `__pycache__`).

## What it is

A Hermes **memory provider** plugin that keeps durable, human-auditable
knowledge in Excel workbooks alongside the built-in `MEMORY.md` / `USER.md`.
It mirrors the built-in memory write path, adds a proactive `prefetch()` hook
that injects indexed memories into every session, and exposes `search` / `read`
tools for on-demand retrieval — so the agent retrieves as well as stores.

## Why (the bug this fixes)

The original plugin logged every `add` to a file and relied on a free-model
sub-agent to *classify + store* it. When that sub-agent was unreachable, the
classifier returned nothing and **the worker silently deleted the log** —
memory was written but never retrievable ("saved but unusable"). This is a
**data-loss bug**, which Hermes lists as the #1 contribution priority.

## How it was fixed (verified, not theoretical)

- **`__init__.py` — direct-store mode** in `_handle_add`: when the agent passes
  explicit `zone` + `brief`/`content`, the record is written **straight to the
  Excel store** via `store.add`, with no LLM / indexer dependency. Durable path
  that behaves like built-in memory.
- **`worker.py` — no silent drop**: a log is kept (renamed back) when
  classification fails, so it is retried instead of lost.
- **`store.py` — master index gains a `title` column**; `search_index` now scans
  `zone + brief + title + tags`.
- **`worker.py` — raw-text fallback**: when the classifier is unavailable the
  worker stores a raw-text backup into `knowledge` (tag `auto-backup`) instead
  of emitting an empty record.
- **`store.py` — exception safety**: `add()` is wrapped so a transient file/lock
  error returns `-1` and logs, never crashes the provider.
- **`store.py` — cross-process file lock**: an atomic lock file (records
  `<pid>:<ts>`, clears stale locks only after confirming the owner pid is dead)
  serializes writes across multiple Hermes processes sharing one `root`.
- **`__init__.py` — lazy index**: `prefetch()` drains any pending seq-logs
  before searching, so memory is never stranded on disk if the background
  indexer crashed.
- **`__init__.py` — `on_session_end` no data loss**: when the LLM classifier
  fails, `_auto_extract` already falls back to `_fallback_store` (raw turns into
  the `knowledge` zone); an explicit raw-transcript backup is added as defence
  in depth.

## Verification

`tests/test_excel_line.py` — **33 tests, all green**:
- store add / search (unicode, Vietnamese, case-insensitive, **by title**, limit) / read / count / zones
- concurrent writes (5×20 threads, no corruption)
- concurrent **direct-store + worker** on the same store (no corruption)
- **cross-process** two `ExcelLineStore` instances, 60 rows, no loss, stale-lock pid recovery
- worker: keeps log on classify-fail, raw-backup on fail, deletes on success
- provider: direct-store returns `stored`, retrievable, `prefetch` matches,
  `system_prompt_block` active; direct-store works **with LLM broken**
- legacy sequence mode still logs
- `add` returns `-1` (no crash) on file failure
- `prefetch()` lazily drains pending seq-logs and finds them
- `on_session_end` keeps data when the LLM fails (fallback store)

Run: `uv run --with openpyxl python tests/test_excel_line.py`

Also verified with the real runtime: `hermes plugins doctor .` →
`OK: runtime discovery, manifest parsing, import, and registration passed`.

## Submission path

Per `AGENTS.md`, third-party plugins ship as **standalone plugin repos**
installed into `~/.hermes/plugins/` (or via `hermes plugins install <git-url>`),
not into the core tree. This plugin is already installed standalone at
`~/.hermes/plugins/excel_line/`. Recommended publish steps:

1. Push this repo to GitHub (only the tracked files — no `*.xlsx` user data).
2. Users install with: `hermes plugins install <owner>/excel_line`
3. Promote in Nous Research Discord `#plugins-skills-and-skins`.

If a community plugin **index** repo exists separately, a PR adding
`excel_line` to that index (pointing at the Git URL + a short description) is
the path to appear in `hermes plugins search`.

## Setup for installers (also in README.md)

Prerequisites: Python 3.10+, `openpyxl` (`pip install openpyxl`, or it is
auto-installed by the plugin's `pip_dependencies`).

Config (`config.yaml`):

```yaml
memory:
  provider: excel_line
  enabled: true
flush_min_turns: 6
```

Restart Hermes. Done — memory auto-retrieves every session and mirrors
built-in writes. No API key or manual setup required; the optional
auto-extract sub-agent uses the configured free model automatically and falls
back to a raw-text backup if it is unavailable.
