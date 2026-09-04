# excel_line — Sub-agent Memory Pipeline (Tracking)

How a raw conversation turn becomes a row in the Excel memory tree.
Every stage below is implemented in `worker.py` (pipeline), `__init__.py`
(entry points + model rotation) and `brain_store.py` (tree storage).

```
USER TURN                     BACKGROUND (never blocks the agent)
──────────                    ────────────────────────────────────
[agent hook on_turn] ───────► logs/sess_<id>_i<n>_o<n>.jsonl   (append record)
                                       │
                                       ▼
                       ┌──────────────────────────────┐
                       │ TRIGGER (one of)             │
                       │ • on_session_end hook        │
                       │ • provider._schedule_index   │
                       │ • lazy drain on prefetch*    │
                       └──────────────────────────────┘
                                       │  index_while_logs_present()
                                       ▼
                          process_logs(log_dir, root, fn)
                                       │
        ┌──────────────────────────────┼─────────────────────────────┐
        │ STAGE 1: atomic take         │ STAGE 2: read + parse       │
        │ rename → .processing         │ JSONL line-by-line;         │
        │ (concurrent append goes to   │ bad LINE kept as            │
        │  a fresh file, next cycle)   │ {_malformed:true} — never   │
        │                              │ dropped silently            │
        └──────────────────────────────┴─────────────────────────────┘
                                       ▼
                        STAGE 3: LLM CURATOR (tri-state)
              prompt = policy text + CURRENT TREE ascii + the RECORD
              model chain (all keyless-fallback, local-first):
                user pin (/excel-line model) → config free_model
                → OpenCode-Zen free rotation → host ctx.llm (last resort)
              ┌───────────────┬───────────────┬───────────────┐
              │ status "ok"   │ "garbage"     │ "down"        │
              │ valid JSON    │ unparseable / │ no model at   │
              │ action field  │ unknown action│ all answered  │
              └───────┬───────┴───────┬───────┴───────┬───────┘
                      ▼               │               │
        STAGE 4: _apply(store, decision)              │
        • none  → log deleted, NOTHING stored         │
        • add   → store.add(branch,...)               │
                   ├ branch full → AUTO-SPLIT:        │
                   │   child("<name>-x") then retry   │
                   │   add inside new file            │
        • merge → store.merge(ids→1 compressed row)   │
        • child → store.child(parent,name)            │
                      │               │               │
                      ▼               ▼               ▼
              stored ok?      failed_records ────► retry_<uuid>.jsonl
              yes → count++   (kept, rewritten      (classifier down or
              no  → keep for   with uuid name,       garbage: retried
                    retry       never lost)          next cycle)
```

\* `prefetch()` itself never calls an LLM (stays cheap); draining is done by
the background indexer thread.

## Decision policy given to the curator LLM (worker.py `_CLASSIFY_PROMPT`)
1. `none` — chitchat, transient commands, narration, one-off task orders.
   **Verbatim user messages are never stored.**
2. Store only durable, reusable knowledge / preferences / conventions.
3. Write title/content/tags in **the user's primary language**
   (agent & code stay English; memory content follows the user).
4. Prefer `merge` (compress into older rows) over near-duplicate `add`.
5. Branch full (10/10) + genuinely new fact → `child` split.
6. Leaf assets only when the record references a real file.

## Invariants (anti-pollution / anti-loss)
| Rule | Enforcement |
|---|---|
| LLM down ⇒ store untouched | v1's raw-dump fallback was removed after a 5.8k-junk-row audit; logs stay on disk and retry |
| Malformed JSONL line survives | rewritten into a `retry_*.jsonl`, never deleted |
| Retry filename collision | `uuid4` suffix (same-PID threads can't clash) |
| Retry write fails | original `.processing` file is NOT deleted (no data loss) |
| Drain loop can't spin forever | stops when a pass stores 0 and file count doesn't drop |
| IDs never reused / collide | `.id_seq` + tree-max floor guard in `BrainStore._next_id` |
| User map edits respected | rows added via map carry tag `map-added`; the agent curator must not re-process them |
| Formula injection | every cell written through `_safe_cell` (`=`, `+`, `-`, `@` escaped) |

## Where things live
| Path | Role |
|---|---|
| `<root>/logs/session_*.jsonl` | queued turn records awaiting classification |
| `<root>/logs/retry_*.jsonl` | records the curator could not process yet |
| `<root>/logs/session_raw_*.jsonl` | raw transcript backups (skipped by indexer on purpose) |
| `<root>/brain.xlsx` | tree root, ≤10 rows |
| `<root>/**.xlsx` | branch workbooks (≤10 rows each) |
| `<root>/.id_seq` | monotonic id counter |

## Manual / batch entry points
* `scripts/sync_builtin_memory.py` — mirror Hermes builtin `MEMORY.md`/`USER.md`
  into the tree **locally, with no LLM and no network** (direct store writes).
  Also deletes stale v1 `[MEMORY (profile)] #N` mirror rows first.
* Provider tool `excel_line(add, branch=…)` — direct write, bypasses the pipeline.
* `python scripts/brain_server.py` — live mind map (reads the same tree).

## Privacy
All memory data stays under `<HERMES_HOME>/excel_line` on the local disk and
is `.gitignore`-excluded. Only code, tests and docs are pushed. Never feed
memory content into the pipeline's only network call path (the model chain)
from ad-hoc scripts unless the user explicitly asks for LLM classification.
