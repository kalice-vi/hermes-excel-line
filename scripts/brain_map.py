"""brain_map.py — SHARED LANGUAGE between excel_line (v2 tree) and the mind map.

The tree is rendered straight from the BrainStore workbooks — Excel is the
single source of truth; the map is a live projection of it.

Node conventions (understood by server + editor):
    root              : the brain node
    f:<file>.xlsx     : a branch FILE node (points to a workbook)
    m:<row_id>        : a memory row inside its branch workbook
    l:<row_id>        : a leaf-asset pointer of that memory (🔗)
    new:...           : ephemeral node created in the editor (becomes m:<id>
                        once the server has written it to Excel)

Data direction:
    Excel (store, with locks) --render--> Map          : always
    Map --ops (add/rename/retitle/move/delete)--> store : every edit
    Memories added by the user on the map carry tag 'map-added'
    => the agent only curates memories IT created, never map-added ones.

v2 layout (see brain_store.py): every .xlsx holds <= 10 rows
(ID | Title | Content | Tags | Branch | Updated). A row's Branch cell is
either a child .xlsx (node), a file path (leaf), or empty (plain memory).
"""
from __future__ import annotations
import os, sys, time
from typing import Dict, List

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from brain_store import BrainStore, MASTER_V2, BadBranch


def build_tree(root: str) -> Dict:
    """Project the Excel tree into a jsMind-style node tree (read-only)."""
    store = BrainStore(root)

    def walk(branch: str) -> Dict:
        try:
            rows = store.load_rows(branch)
        except BadBranch:
            rows = []
        children = []
        for r in rows:
            rid = int(r["id"])
            rb = str(r.get("branch") or "")
            sub = []
            if rb.lower().endswith(".xlsx"):
                sub.append(walk(rb))                       # nested file node
            elif rb:
                sub.append({"id": f"l:{rid}",
                            "topic": "🔗 " + os.path.basename(rb.replace("\\\\", "\\"))})
            node = {"id": f"m:{rid}",
                    "topic": f"📄{rid} {str(r.get('title') or '').strip()}",
                    "branch": branch}
            if sub:
                node["children"] = sub
            children.append(node)
        name = os.path.basename(branch)
        return {"id": f"f:{branch}", "topic": f"📁 {name} · {len(rows)}",
                "branch": branch, "usage": f"{len(rows)}/10", "children": children}

    tree = walk(MASTER_V2)
    tree["id"] = "root"
    tree["topic"] = f"🧠 BRAIN · {store.count()} memories · {time.strftime('%H:%M:%S')}"
    return {"meta": {"name": "brain", "author": "excel_line", "version": "2.0.0"},
            "format": "node_tree", "data": tree}


def topic_of(tags) -> str:
    """First meaningful tag of a row (legacy helper kept for compatibility)."""
    t = (str(tags or "")).strip()
    for part in t.split(","):
        part = part.strip()
        if part and part != "map-added":
            return part
    return "misc"


if __name__ == "__main__":
    import json
    root = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\Admin\AppData\Local\hermes\excel_line"
    t = build_tree(root)
    print(json.dumps(t["data"]["topic"], ensure_ascii=False))
    n = len(t["data"]["children"])
    print("top-level branch files:", n)
