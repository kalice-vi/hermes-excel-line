# excel_line — Hierarchical Excel-backed Memory for Hermes

An advanced Hermes **memory provider** plugin that keeps durable, human-auditable knowledge organized in a **hierarchical tree of Excel workbooks** alongside the built-in `MEMORY.md` / `USER.md`.

- **Root Workbook (`brain.xlsx`)** — Layer 1 root node (header + **MAX 10 rows** per file).
- **Sub-Branch Nodes (`<branch>.xlsx`)** — Branch points pointing to child workbooks (`user.xlsx`, `skill.xlsx`, `knowledge.xlsx`, `knowledge/misc.xlsx`, etc.).
- **Leaf Pointer Assets** — References to scripts, documents, images, or external files (`.py`, `.md`, `.png`) sitting at the bottom of the tree.
- **Dynamic Branching & Compression** — Enforces a strict 10-row cap per file. Automatically compresses older memories via `merge` or splits branches via `child`.
- **Keyless Model Rotation** — Automatically classifies background memory logs using keyless models (OpenCode Zen free rotation), with `/excel-line model` CLI configuration.

It runs **alongside** built-in memory: built-in is the fast cache, `excel_line` is the durable, spreadsheet-readable long-term memory tree.

---

## 🌳 How the Memory Tree Works

```text
brain.xlsx (Root Node, Max 10 rows)
  ├── knowledge.xlsx (2/10)
  │     ├── knowledge/misc.xlsx (9/10)
  │     └── knowledge/accounting.xlsx (4/10)
  ├── user.xlsx (2/10)
  │     └── user/profile.xlsx (10/10)
  ├── skill.xlsx (2/10)
  │     └── 🔗 Skill Script → C:/Users/Admin/.hermes/skills/my_tool.py (Leaf)
  ├── pref.xlsx (2/10)
  ├── contact.xlsx (1/10)
  └── project.xlsx (1/10)
```

### Key Invariants & Rules:
1. **10-Row Cap Per Workbook**: Every `.xlsx` file holds at most 10 rows (`ID | Title | Content | Tags | Branch | Updated`).
2. **Nodes vs. Leaves**:
   - A row whose `Branch` ends in `.xlsx` is a **Node** pointing to a child workbook.
   - A row whose `Branch` points to a non-`.xlsx` path is a **Leaf Asset** (unmoved resource file).
   - A row with an empty `Branch` is a **Pure Memory Entry**.
3. **Free Leaves**: Asset files (like Hermes skills or local user documents) sit freely in their original locations; the tree holds a relative or absolute leaf pointer (`🔗`).

---

## 🚀 Tool Actions

| Action | Description | Example Usage |
|---|---|---|
| `add` | Save a memory directly to a branch | `excel_line(action="add", branch="brain.xlsx", title="…", content="…", tags="…")` |
| `read` | Read rows and usage of a branch | `excel_line(action="read", branch="skill.xlsx")` |
| `tree` | Render an ASCII visual tree | `excel_line(action="tree")` |
| `child` | Create a new child sub-branch workbook | `excel_line(action="child", parent="skill.xlsx", name="devops", title="DevOps Branch")` |
| `merge` | Compress multiple row IDs into 1 | `excel_line(action="merge", branch="skill.xlsx", ids=[1, 2, 3], title="Summary", content="…")` |
| `promote` | Promote a leaf asset to a node directory | `excel_line(action="promote", row_id=5, name="my_tool")` |
| `move` | Rebalance row IDs between workbooks | `excel_line(action="move", src="skill.xlsx", ids=[4], dst="skill/devops.xlsx")` |
| `search` | Search keywords across all tree workbooks | `excel_line(action="search", query="accounting")` |
| `delete` | Delete a row by ID | `excel_line(action="delete", row_id=12)` |

---

## 🛠️ Setup & Installation

### 1. Prerequisites
- Python ≥ 3.10
- The `openpyxl` package (auto-installed via `plugin.yaml`).

### 2. Enable in Hermes Config (`config.yaml`)
```yaml
memory:
  memory_enabled: true
  provider: excel_line
plugins:
  excel_line:
    root: "$HERMES_HOME/excel_line"
    log_dir: "$HERMES_HOME/excel_line/logs"
```

### 3. Model Configuration (`/excel-line model`)
Configure the background classifier model directly from the chat:

```text
/excel-line model          # List available keyless models and current selection
/excel-line model auto     # OpenCode Zen keyless free rotation -> host default
/excel-line model host     # Use main Hermes agent model
/excel-line model <number> # Select specific free model by number
```

---

## 🧪 Testing

Run the full test suites:

```bash
cd plugins/excel_line
python tests/test_brain_store.py   # Unit tests for BrainStore tree (21/21)
python tests/test_v2_e2e.py        # E2E provider tests (19/19)
python tests/test_excel_line.py     # Integration tests (42/42)
python tests/qa_excel_line.py      # Standalone QA runner
```

---

## 🔒 Privacy & Security

Memory workbooks saved in `$HERMES_HOME/excel_line` are personal data and excluded from Git via `.gitignore`. Never publish workbook data files.
