"""mermaid_map.py — common text language between the excel_line v2 tree and Mermaid.

Each .mmd line is one node; indentation = tree depth. Node shapes mirror
brain_store's invariants:

    mindmap
      root((🧠 BRAIN))
        📁 skill.xlsx · 2/10          <- branch workbook (usage counter)
          📄34 Skill title            <- memory row (id glued after 📄)
          📁 skill/cloud.xlsx · 1/10  <- nested branch
            📄40 Sub memory
              🔗40 filename.py        <- leaf-asset pointer of the memory

Excel -> .mmd : build_mmd(root)    pure projection, normalized formatting.
.mmd -> Excel : diff_ops(old,new)  infers ops; the server applies them via
               BrainStore (never touching .xlsx through side paths).
User-added map memories carry the 'map-added' tag (agent leaves them alone).
"""
from __future__ import annotations
import os, re, sys
from typing import Dict, List

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from brain_store import BrainStore, MASTER_V2, BadBranch

IND = "  "   # two spaces per level

# ---------------------------------------------------------------- Excel -> mmd

def _clean(s: str) -> str:
    """Escape characters that break mermaid mindmap syntax; collapse to one line."""
    s = re.sub(r"\s+", " ", str(s))
    return re.sub(r"[\(\)\[\]{}<>#&`;]", " ", s).strip()


def build_mmd(root: str) -> str:
    store = BrainStore(root)
    lines = []

    def walk(branch: str, depth: int):
        try:
            rows = store.load_rows(branch)
        except BadBranch:
            rows = []
        for r in rows:
            rid = int(r["id"])
            rb = str(r.get("branch") or "")
            title = _clean(str(r.get("title") or ""))[:80]
            if rb.lower().endswith(".xlsx"):
                try:
                    usage = len(store.load_rows(rb))
                except BadBranch:
                    usage = 0
                lines.append(f"{IND*(depth)}📄{rid} {title}")
                lines.append(f"{IND*(depth+1)}📁 {_clean(os.path.basename(rb))} · {usage}")
                walk(rb, depth + 2)
            else:
                if rb:
                    title += f" 🔗{_clean(os.path.basename(rb))}"
                lines.append(f"{IND*(depth)}📄{rid} {title}")

    total = store.count()
    lines.append("mindmap")
    lines.append(f"{IND}root((🧠 BRAIN · {total} memories))")
    walk(MASTER_V2, 2)
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------- mmd -> ops

MEM_RE = re.compile(r"📄(\d+)\s+(.*)")           # existing memory
NEW_RE = re.compile(r"📄\s+(\S.*)")              # user typed 📄 with no id


def _parse(text: str) -> Dict[str, Dict]:
    """rid -> {branch, brief}. The enclosing branch = nearest ancestor 📁 line."""
    out: Dict[str, Dict] = {}
    # stack of (indent, branch_name); branch_name None = root level (brain.xlsx)
    stack: List = [(0, None)]

    def enclosing_branch(indent):
        while len(stack) > 1 and stack[-1][0] >= indent:
            stack.pop()
        return stack[-1][1] or MASTER_V2

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped in ("mindmap",):
            continue
        indent = len(line) - len(stripped.lstrip(" "))
        if stripped.startswith("root("):
            stack = [(indent, None)]
            continue
        if stripped.startswith("📁"):
            m = re.match(r"📁\s*([^\s·]+\.xlsx)", stripped)
            if m:
                stack.append((indent, m.group(1)))
            continue
        if stripped.startswith("🔗") or stripped.startswith("l:"):
            continue                                   # leaf pointer — read-only
        if stripped.startswith("📄"):
            branch = enclosing_branch(indent)
            m = MEM_RE.match(stripped)
            if m:
                rid = int(m.group(1))
                brief = re.sub(r"\s*🔗.*$", "", m.group(2)).strip()
                key = rid if rid not in out else f"dup:{rid}:{len(out)}"
                out[key] = {"branch": branch, "brief": brief}
            else:
                m2 = NEW_RE.match(stripped)
                if m2:
                    brief = re.sub(r"\s*🔗.*$", "", m2.group(1)).strip()
                    out["new:" + brief] = {"branch": branch, "brief": brief}
    return out


def diff_ops(old: str, new: str, store: BrainStore = None) -> List[Dict]:
    """Compare the edited .mmd against the live Excel projection -> ops."""
    a, b = _parse(old), _parse(new)
    ops: List[Dict] = []
    for rid, entry in a.items():
        if isinstance(rid, str):
            continue                                   # 'dup:'/'new:' node — ignore
        if rid not in b:
            ops.append({"op": "delete", "id": rid})
        else:
            nb = b[rid]
            if entry["brief"] != nb["brief"]:
                ops.append({"op": "rename", "id": rid, "topic": nb["brief"],
                            "branch": entry["branch"]})
            if entry["branch"] != nb["branch"]:
                ops.append({"op": "move", "id": rid, "src": entry["branch"],
                            "dst": nb["branch"]})
    for key, entry in b.items():
        if isinstance(key, str) and key.startswith("new:"):
            ops.append({"op": "add", "branch": entry["branch"],
                        "topic": entry["brief"], "map_new_id": key})
    return ops


if __name__ == "__main__":
    r = os.environ.get("EXCEL_LINE_ROOT", r"C:\Users\Admin\AppData\Local\hermes\excel_line")
    mmd = build_mmd(r)
    print("lines:", mmd.count("\n"))
    print("memories parsed:", len(_parse(mmd)))
    # self-test: diffing the projection against itself must yield zero ops
    print("self-diff ops:", diff_ops(mmd, mmd))
