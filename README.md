# excel_line — Excel-backed long-term memory for Hermes

A Hermes **memory provider** plugin that keeps durable, human-auditable
knowledge in Excel workbooks alongside the built-in `MEMORY.md` / `USER.md`.

- A **master index** workbook (`excel-line_index.xlsx`) — one brief row per
  memory (zone, brief, title, path, tags) for fast keyword scan.
- **Per-zone** workbooks (`user.xlsx`, `skill.xlsx`, `knowledge.xlsx`, …) hold
  the detailed content.

It runs **alongside** the built-in memory: built-in is the fast cache,
excel_line is the durable, spreadsheet-readable store.

## What it does

| Capability | How |
|---|---|
| Save a fact | `excel_line(add, zone=…, brief=…, content=…)` — written **straight to Excel**, no LLM needed |
| Recall a fact | `excel_line(search, query=…)` or `excel_line(read, zone=…)` |
| Auto-recall each turn | the provider's `prefetch()` hook searches the index on every query and the `system_prompt_block` advertises the store |
| Mirror built-in memory | `on_memory_write` hook copies `MEMORY.md`/`USER.md` writes into Excel |
| Backup pass | `on_session_end` asks a free model to distil the transcript into facts |

## Setup / Installation

### 1. Prerequisites
- Python ≥ 3.10 with [`uv`](https://docs.astral.sh/uv/) available on `PATH`
  (Hermes already uses `uv`).
- The `openpyxl` package (auto-installed via `plugin.yaml` `pip_dependencies`).

### 2. Install the plugin
Copy this folder into your Hermes plugins directory:

```bash
# global install (all profiles)
cp -r excel_line "$HERMES_HOME/plugins/"

# or, per-profile
cp -r excel_line "$HERMES_HOME/profiles/<profile>/plugins/"
```

Then enable it in your Hermes config (`config.yaml`):

```yaml
memory:
  memory_enabled: true
  provider: excel_line      # use excel_line instead of the default file memory
plugins:
  excel_line:
    root: "$HERMES_HOME/excel_line"   # where the workbooks live (default)
    log_dir: "$HERMES_HOME/excel_line/logs"
    free_model: "gemini-3.5-flash-lite"   # model for the auto-classifier
```

> The `root` directory is created automatically on first run. **Do NOT commit
> the `*.xlsx` data files or `logs/`** — they contain your personal memory.
> See `.gitignore`.

### 3. Restart Hermes
The plugin is loaded at startup. After restart you should see a line like
`# Excel-Line Memory — Active. N indexed memories` in the agent's system
prompt, and the `excel_line` tool becomes available.

### 4. (Optional) Sub-agent log processing — works WITHOUT any setup
When you call `excel_line(add, …)` with explicit `zone`+`brief`+`content`, the
record is stored **immediately and directly** (no model call). This is the
recommended, always-works path.

The legacy "log + classify" path (passing `input_seq`/`output_seq` or
`input_text`/`output_text`) uses a free-model sub-agent to classify the entry.
- **If a free model is configured** (`free_model` in config), it classifies
  automatically in the background.
- **If no model is reachable**, the worker **falls back to a raw-text backup**
  so nothing is lost, and keeps retrying the log until a model is available.
  No manual setup required — memory is never silently dropped.

## How the agent should use it

```text
# To SAVE a durable fact (preferred — no LLM dependency):
excel_line(add, zone="skill", brief="Skill: foo — does X",
           content="…concise knowledge…", title="foo", tags="skill,foo")

# To RECALL:
excel_line(search, query="foo")
excel_line(read, zone="skill")
```

`zone` is one of: `user`, `project`, `pref`, `task`, `knowledge`, `contact`,
or any custom name (a new `<zone>.xlsx` is created automatically).

## Running the tests

```bash
cd excel_line
uv run --with openpyxl python tests/test_excel_line.py
```

25 tests cover: store CRUD, unicode/Vietnamese search, search-by-title,
concurrent writes, worker fallback (no data loss), provider direct-store,
prefetch, and backward-compatible sequence mode.

## Privacy note
The workbooks under `root/` are **your** data. They are excluded from this
repo via `.gitignore`. Never publish them.
