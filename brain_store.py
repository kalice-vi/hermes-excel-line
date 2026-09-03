"""brain_store.py — excel_line v2: hierarchical tree store, max 10 rows per .xlsx file.

CONVENTIONS (strict invariants):
  brain.xlsx          : Root node (layer 1) — header + MAX 10 rows
  6 columns           : ID (int) | Title (≤50) | Content (≤250) | Tags
                        | Branch (relative path from root or absolute URL/path) | Updated
  Branch .xlsx        : BRANCH NODE — points to a child workbook (internal)
  Branch non-.xlsx    : LEAF — resource asset file (.py/.png/.md/…), must sit at
                        the bottom of the tree; directory for leaves = parent name
                        ('skill.xlsx' → resources in folder 'skill/')
  Branch empty        : Pure memory entry without an associated file asset
  Rule                : If an object has ≥1 sub-objects behind it, it MUST be a .xlsx node
                        (promote). No exceptions.

  Max 10 rows per file: Overflow raises FullError -> LLM/Agent must merge
                        (compress into existing row) or child+move (branch out).
"""
from __future__ import annotations
import os, re, time, json, shutil
from typing import Dict, List, Optional
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

try:
    from .store import ExcelLineStore, _now
except ImportError:
    from store import ExcelLineStore, _now  # standalone / single-file test run

_safe_cell = ExcelLineStore._safe_cell

MAX_ROWS = 10
MASTER_V2 = "brain.xlsx"
TITLE_CAP = 50
CONTENT_CAP = 250
V2_HDR = ["ID", "Title", "Content", "Tags", "Branch", "Updated"]


class BadBranch(ValueError):
    """Raised when a branch path is invalid or unresolvable."""


class FullError(Exception):
    """Raised when an attempt is made to write to a branch node with 10/10 rows."""
    def __init__(self, branch):
        super().__init__(f"{branch} is full ({MAX_ROWS}/{MAX_ROWS}) — merge existing rows or use child() to branch")
        self.branch = branch


class BrainStore(ExcelLineStore):
    """v2 Tree Store inheriting multi-process locking from v1 base store.

    Manages hierarchical tree workbooks rooted at brain.xlsx.
    """

    def __init__(self, root_dir: str):
        super().__init__(root_dir)
        self._root = os.path.abspath(root_dir)
        os.makedirs(self._root, exist_ok=True)
        self._ensure_v2_root()

    def _ensure_v2_root(self):
        root_file = self.resolve(MASTER_V2)
        if not os.path.exists(root_file):
            wb = Workbook()
            ws = wb.active
            ws.title = "Index"
            ws.append(V2_HDR)
            ws.freeze_panes = "A2"
            wb.save(root_file)

    def resolve(self, branch: str) -> str:
        """Resolve branch path to an absolute on-disk path.

        Two types:
          - Internal (relative to store root, e.g. 'user.xlsx', 'skill/x.py'):
            located inside the excel_line root directory.
          - External Leaf (absolute path, e.g. 'C:/Users/Admin/.hermes/skills/x/SKILL.md'):
            unmoved external asset — the tree holds a pointer (leaf stem) to it.
        """
        b = (branch or "").strip().replace("\\", "/").lstrip("/")
        if not b or ".." in b.split("/") and ":" not in b:
            raise BadBranch("Invalid branch path: " + repr(branch))
        # absolute windows/posix path => external leaf (read/write pointer only)
        if re.match(r"^[A-Za-z]:/", b) or b.startswith("/"):
            return os.path.abspath(b)
        # relative path => internal asset inside store root
        abs_p = os.path.abspath(os.path.join(self._root, b))
        if not abs_p.startswith(self._root):
            raise BadBranch("Branch path escapes store root: " + repr(branch))
        return abs_p

    def is_external(self, branch: str) -> bool:
        """Return True if branch is an absolute external path outside store root."""
        b = (branch or "").strip().replace("\\", "/")
        if re.match(r"^[A-Za-z]:/", b) or b.startswith("/"):
            abs_p = os.path.abspath(b)
            return not abs_p.startswith(self._root)
        return False

    def rel_branch(self, path: str) -> str:
        """Convert path to relative branch string if inside root, else return absolute path."""
        p = os.path.abspath(path).replace("\\", "/")
        r = self._root.replace("\\", "/")
        if p.startswith(r + "/"):
            return p[len(r) + 1:]
        return p

    def next_id() -> int:
        pass  # inherited from base class via self._next_seq()

    def _next_id(self) -> int:
        seq_file = os.path.join(self._root, ".id_seq")
        val = 0
        if os.path.exists(seq_file):
            try:
                with open(seq_file, "r", encoding="utf-8") as f:
                    val = int(f.read().strip() or "0")
            except Exception:
                val = 0
        val += 1
        with open(seq_file, "w", encoding="utf-8") as f:
            f.write(str(val))
        return val

    # ---------------- I/O ----------------
    def load_rows(self, branch: str) -> List[Dict]:
        fp = self.resolve(branch)
        if not os.path.exists(fp):
            raise BadBranch("File not found: " + branch)
        try:
            wb = load_workbook(fp, data_only=True)
        except Exception as e:
            raise BadBranch(f"Cannot read workbook {branch}: {e}")
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        hdr = [str(c or "").strip().lower() for c in rows[0]]
        out = []
        for r in rows[1:]:
            if not any(r):
                continue
            rec = {}
            for col_name, val in zip(hdr, r):
                rec[col_name] = "" if val is None else val
            try:
                rec["id"] = int(rec.get("id") or 0)
            except (ValueError, TypeError):
                continue
            if rec["id"] > 0:
                rec["title"] = str(rec.get("title") or "")
                rec["content"] = str(rec.get("content") or "")
                rec["tags"] = str(rec.get("tags") or "")
                b_val = str(rec.get("branch") or "").strip()
                rec["branch"] = b_val if b_val else None
                rec["updated"] = str(rec.get("updated") or "")
                out.append(rec)
        return out

    def _write_rows(self, path: str, rows: List[Dict]):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        wb = Workbook()
        ws = wb.active
        ws.title = "Index"
        ws.append(V2_HDR)
        ws.freeze_panes = "A2"
        for r in rows:
            ws.append([
                r["id"],
                _safe_cell(str(r.get("title") or "")[:TITLE_CAP]),
                _safe_cell(str(r.get("content") or "")[:CONTENT_CAP]),
                _safe_cell(str(r.get("tags") or "")),
                _safe_cell(str(r.get("branch") or "")),
                _safe_cell(str(r.get("updated") or _now()))
            ])
        for col_num in range(1, len(V2_HDR) + 1):
            col_letter = get_column_letter(col_num)
            ws.column_dimensions[col_letter].width = 22
        wb.save(path)

    # ---------------- validation ----------------
    def _validate_branch_value(self, link: str):
        if not link:
            return
        norm = link.replace("\\", "/")
        if norm.endswith("/"):
            raise BadBranch("Leaf pointer cannot be a directory ending in '/': " + repr(link))
        if norm.lower().endswith(".xlsx"):
            return
        fp = self.resolve(link)
        if not os.path.exists(fp):
            raise BadBranch(f"Leaf file does not exist on disk: {link} (resolved: {fp})")

    # ---------------- CORE OPERATIONS ----------------
    def add(self, branch: str = MASTER_V2, title: str = "", content: str = "",
            tags: str = "", link: str = "", **kwargs) -> int:
        """Add a row to a branch workbook. Max 10 rows allowed per file."""
        if not title and "brief" in kwargs:
            title = str(kwargs["brief"])[:TITLE_CAP]
        if not content and "brief" in kwargs:
            content = str(kwargs["brief"])[:CONTENT_CAP]
        if not branch or branch in ("knowledge", "user", "pref", "project", "skill", "contact", "task"):
            target_branch = f"{branch}.xlsx" if branch else MASTER_V2
        else:
            target_branch = branch
        if not target_branch.lower().endswith(".xlsx"):
            target_branch = target_branch + ".xlsx"

        target_fp = self.resolve(target_branch)
        if not os.path.exists(target_fp):
            self._write_rows(target_fp, [])

        self._validate_branch_value(link)

        with self._cross_process_lock(timeout=15.0):
            rows = self.load_rows(target_branch)
            if len(rows) >= MAX_ROWS:
                raise FullError(target_branch)
            rid = self._next_id()
            rec = {
                "id": rid,
                "title": (title or "Untitled")[:TITLE_CAP],
                "content": (content or "")[:CONTENT_CAP],
                "tags": tags or "",
                "branch": link if link else None,
                "updated": _now()
            }
            rows.append(rec)
            self._write_rows(target_fp, rows)
            return rid

    def child(self, parent: str, name: str, title: str = "", content: str = "",
              tags: str = "") -> Dict:
        """Create a new sub-branch .xlsx file under parent + linking row in parent."""
        parent_rel = parent if parent.endswith(".xlsx") else parent + ".xlsx"
        clean_name = re.sub(r"[^a-zA-Z0-9_\-]", "_", name).strip("_")
        if not clean_name:
            raise BadBranch("Invalid sub-branch name: " + repr(name))

        if parent_rel == MASTER_V2:
            child_rel = f"{clean_name}.xlsx"
        else:
            parent_dir = parent_rel[:-5]
            child_rel = f"{parent_dir}/{clean_name}.xlsx"

        child_fp = self.resolve(child_rel)
        os.makedirs(os.path.dirname(child_fp), exist_ok=True)
        if not os.path.exists(child_fp):
            self._write_rows(child_fp, [])

        node_id = self.add(parent_rel, title=(title or clean_name)[:TITLE_CAP],
                           content=(content or f"Branch node -> {child_rel}")[:CONTENT_CAP],
                           tags=tags, link=child_rel)
        return {"node_id": node_id, "parent": parent_rel, "file": child_rel}

    def set(self, branch: str, row_id: int, title: str = "", content: str = "",
            tags: str = "", link: Optional[str] = None) -> bool:
        """Update an existing row in a branch workbook."""
        with self._cross_process_lock(timeout=15.0):
            rows = self.load_rows(branch)
            hit = False
            for r in rows:
                if r["id"] == row_id:
                    if title:
                        r["title"] = title[:TITLE_CAP]
                    if content:
                        r["content"] = content[:CONTENT_CAP]
                    if tags is not None and tags != "":
                        r["tags"] = tags
                    if link is not None:
                        self._validate_branch_value(link)
                        r["branch"] = link if link else None
                    r["updated"] = _now()
                    hit = True
                    break
            if hit:
                self._write_rows(self.resolve(branch), rows)
            return hit

    def rm(self, branch: str, row_id: int, purge_file: bool = False) -> bool:
        """Remove a row by ID from a branch workbook."""
        with self._cross_process_lock(timeout=15.0):
            rows = self.load_rows(branch)
            for r in rows:
                if r["id"] == row_id:
                    rows.remove(r)
                    self._write_rows(self.resolve(branch), rows)
                    if purge_file and r.get("branch") and not r["branch"].lower().endswith(".xlsx"):
                        if self.is_external(r["branch"]):
                            pass  # External leaf: keep asset intact
                        else:
                            try:
                                os.remove(self.resolve(r["branch"]))
                            except OSError:
                                pass
                    return True
            return False

    def promote(self, branch: str, row_id: int, name: str) -> Dict:
        """Promote a single leaf asset row into a .xlsx node directory."""
        rows = self.load_rows(branch)
        orig = next((r for r in rows if r["id"] == row_id), None)
        if orig is None:
            raise BadBranch(f"Row #{row_id} not found in {branch}")
        asset = str(orig.get("branch") or "")
        self.rm(branch, row_id)
        node = self.child(branch, name, title=str(orig.get("title") or name)[:TITLE_CAP],
                          content=str(orig.get("content") or ""),
                          tags=str(orig.get("tags") or ""))
        node_fp = self.resolve(node["file"])
        new_asset_rel = ""
        if asset:
            if self.is_external(asset):
                new_asset_rel = asset
            else:
                src_abs = self.resolve(asset)
                dst_rel = os.path.join(os.path.basename(node["file"]).replace(".xlsx", ""),
                                       os.path.basename(asset)).replace("\\", "/")
                dst_abs = self.resolve(dst_rel)
                os.makedirs(os.path.dirname(dst_abs), exist_ok=True)
                if os.path.exists(src_abs) and src_abs != dst_abs:
                    shutil.move(src_abs, dst_abs)
                new_asset_rel = dst_rel

        asset_id = self.add(node["file"], title=str(orig.get("title") or name)[:TITLE_CAP],
                            content=str(orig.get("content") or "")[:CONTENT_CAP],
                            tags=str(orig.get("tags") or ""), link=new_asset_rel)
        return {"node_id": node["node_id"], "file": node["file"], "asset_id": asset_id, "asset": new_asset_rel}

    def move(self, src_branch: str, row_ids: List[int], dst_branch: str) -> int:
        """Move row IDs from src_branch to dst_branch."""
        if src_branch == dst_branch:
            return 0
        moved = 0
        for rid in list(row_ids):
            f = self.find(rid)
            if not f or f["branch"] != src_branch:
                continue
            r = f["row"]
            self.add(dst_branch, title=r["title"], content=r["content"],
                     tags=r["tags"], link=r.get("branch") or "")
            self.rm(src_branch, rid)
            moved += 1
        return moved

    def merge(self, branch: str, row_ids: List[int], title: str, content: str,
              tags: str = "", link: str = "") -> int:
        """Compress several row IDs in branch into ONE new summarized memory row."""
        if not row_ids:
            raise BadBranch("merge requires at least one row ID")
        for rid in row_ids:
            self.rm(branch, rid)
        new_id = self.add(branch, title=title[:TITLE_CAP], content=content[:CONTENT_CAP],
                          tags=tags, link=link)
        return new_id

    # ---------------- LOOKUP & TREE ----------------
    def find(self, row_id: int) -> Optional[Dict]:
        """Find a row ID across the entire tree."""
        def walk(b):
            for r in self.load_rows(b):
                if r["id"] == row_id:
                    return {"branch": b, "row": r}
                rb = str(r.get("branch") or "")
                if rb.lower().endswith(".xlsx"):
                    got = walk(rb)
                    if got:
                        return got
            return None
        return walk(MASTER_V2)

    def search(self, q: str, limit: int = 20) -> List[Dict]:
        """Search keyword across all tree workbooks."""
        q = (q or "").lower().strip()
        hits = []
        def walk(b):
            for r in self.load_rows(b):
                blob = " ".join(str(r.get(k) or "") for k in ("title", "content", "tags")).lower()
                if q and q in blob:
                    hits.append({"branch": b, "path": self.path_of(r["id"]),
                                 "id": r["id"], "title": r["title"],
                                 "tags": r["tags"]})
                rb = str(r.get("branch") or "")
                if rb.lower().endswith(".xlsx"):
                    walk(rb)
        if q:
            try:
                walk(MASTER_V2)
            except BadBranch:
                pass
        return hits[:limit]

    def path_of(self, row_id: int) -> str:
        """Return human path chain: brain.xlsx › user.xlsx › #17."""
        chain = []
        def walk(b, acc):
            for r in self.load_rows(b):
                curr = acc + [b]
                if r["id"] == row_id:
                    chain.extend(curr + [f"#{row_id}"])
                    return True
                rb = str(r.get("branch") or "")
                if rb.lower().endswith(".xlsx"):
                    if walk(rb, curr):
                        return True
            return False
        walk(MASTER_V2, [])
        return " › ".join(chain) if chain else f"#{row_id}"

    def count() -> int:
        pass

    def count(self) -> int:
        """Total memory rows across all .xlsx workbooks in the tree."""
        total = 0
        def walk(b):
            nonlocal total
            rows = self.load_rows(b)
            total += len(rows)
            for r in rows:
                rb = str(r.get("branch") or "")
                if rb.lower().endswith(".xlsx"):
                    walk(rb)
        try:
            walk(MASTER_V2)
        except BadBranch:
            pass
        return total

    def tree(self) -> str:
        return self.tree_text()

    def tree_text(self) -> str:
        """Render ASCII tree representation of the memory tree with usage counters."""
        lines = []
        def walk(b, indent=""):
            try:
                rows = self.load_rows(b)
            except BadBranch:
                return
            if indent == "":
                lines.append(b)
            for r in rows:
                rb = str(r.get("branch") or "")
                if rb.lower().endswith(".xlsx"):
                    try:
                        child_count = len(self.load_rows(rb))
                    except Exception:
                        child_count = 0
                    lines.append(f"{indent}  {rb} ({child_count}/{MAX_ROWS})")
                    walk(rb, indent + "    ")
                else:
                    link_mark = " 🔗" if rb else " "
                    link_str = f" → {rb}" if rb else ""
                    lines.append(f"{indent}  #{r['id']}{link_mark}{r['title']}{link_str}")
        walk(MASTER_V2)
        return "\n".join(lines)

    # ---------------- COMPATIBILITY ADAPTERS FOR V1 ----------------
    def read_zone(self, path: str, limit: int = 10) -> List[Dict]:
        b = self.rel_branch(path)
        try:
            return self.load_rows(b)[:limit]
        except BadBranch:
            return []

    def zone_path(self, zone: str) -> str:
        clean = (zone or "knowledge").strip()
        if not clean.endswith(".xlsx"):
            clean += ".xlsx"
        return self.resolve(clean)

    def list_zones(self) -> List[str]:
        out = []
        for root_dir, _, files in os.walk(self._root):
            for f in files:
                if f.endswith(".xlsx"):
                    rel = os.path.relpath(os.path.join(root_dir, f), self._root).replace("\\", "/")
                    out.append(os.path.splitext(rel)[0])
        return sorted(out)

    def delete(self, zone: str = "", row_id: int = 0) -> bool:
        f = self.find(int(row_id)) if row_id else None
        if not f:
            return False
        return self.rm(f["branch"], int(row_id))

    def update(self, zone: str = "", row_id: int = 0, brief: str = "", content: str = "",
               title: str = "", tags: str = "") -> bool:
        f = self.find(int(row_id)) if row_id else None
        if not f:
            return False
        return self.set(f["branch"], int(row_id),
                        title=(title or brief)[:TITLE_CAP], content=content, tags=tags)

    def forget(self, query: str) -> int:
        q = (query or "").lower().strip()
        if not q:
            return 0
        rows = [r for r in _flatten_rows(self)
                if q in " ".join(str(r.get(k) or "") for k in ("title", "content", "tags")).lower()]
        for r in rows:
            self.rm(r["_branch"], r["id"])
        return len(rows)


def _flatten_rows(store: BrainStore) -> List[Dict]:
    out = []
    def walk(b):
        for r in store.load_rows(b):
            out.append({**r, "_branch": b})
            rb = str(r.get("branch") or "")
            if rb.lower().endswith(".xlsx"):
                walk(rb)
    try:
        walk(MASTER_V2)
    except BadBranch:
        pass
    return out


def search_index(self, query: str, limit: int = 10) -> List[Dict]:
    q = (query or "").lower().strip()
    if not q:
        return []
    hits = []
    for r in _flatten_rows(self):
        blob = " ".join(str(r.get(k) or "") for k in ("title", "content", "tags", "branch")).lower()
        if q in blob:
            hits.append({
                "id": r["id"],
                "zone": os.path.splitext(r["_branch"])[0],
                "brief": r.get("title") or str(r.get("content") or "")[:120],
                "title": r.get("title"),
                "path": self.resolve(r["_branch"]),
                "tags": r.get("tags") or ""
            })
    return hits[:limit]


BrainStore.search_index = search_index
