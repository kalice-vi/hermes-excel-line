# excel_line

Excel-backed long-term memory with a hierarchical tree store. Each `.xlsx`
holds at most 10 rows (`ID | Title | Content | Tags | Branch | Updated`); a row
whose `Branch` ends in `.xlsx` is a node pointing to a child workbook, a
non-`.xlsx` branch is a leaf asset, and an empty branch is a plain memory.

## Harness-agnostic core

`excel_line_core/` is pure Python + `openpyxl` with **zero** agent-runtime
imports. It can be pip-installed and used from any harness:

```bash
pip install excel-line   # → import excel_line_core
```

```python
from excel_line_core import BrainStore
store = BrainStore("./brain")
rid = store.add("brain.xlsx", title="note", content="...", tags="x")
print(store.tree())
```

## Hermes plugin

The Hermes entry point in `__init__.py` wraps the core; `store.py`,
`brain_store.py` and `worker.py` are thin shims kept for backward
compatibility and still re-export the core API.

## Privacy & model egress

The optional memory classifier can call an LLM to summarize agent I/O logs
into knowledge rows. **By default it uses the host's own configured model**
(`preferred = "host"`) — no transcript excerpts leave your machine or account.

Rotation through third-party keyless models is strictly **opt-in**: run
`/excel-line model auto` (or pick a specific free model by number) to enable
it. Revert any time with `/excel-line model host`. Until you opt in, nothing
is sent to any external endpoint.

## Brain viewer server

The bundled brain-map HTTP server (`scripts/brain_server.py`, port 8766) is
**not started automatically**. Start it manually when needed:

```bash
python scripts/brain_server.py   # serves http://127.0.0.1:8766
```
