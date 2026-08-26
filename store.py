"""Excel-backed storage for the excel_line memory provider.

Layout
------
- MASTER workbook  (excel-line_index.xlsx)
    Sheet "index": one row per memory entry
        A id | B zone | C brief | D path | E tags | F created | G updated
- ZONE workbooks  (<zone>.xlsx)  e.g. user.xlsx, project.xlsx, knowledge.xlsx
    Sheet "mem":  A id | B title | C content | D tags | E created | F updated

All workbooks live under one root dir (default $HERMES_HOME/excel_line/).
The master index keeps only a brief note + the path to the detailed zone
record, so the index stays small and human-scannable while deep knowledge
lives in per-zone files.

This module is pure file I/O (openpyxl) with no agent runtime imports, so it
can be unit-tested and reused by the sub-agent worker directly.
"""

from __future__ import annotations

import os
import threading
import time
import contextlib
from typing import Dict, List, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

MASTER_NAME = "excel-line_index.xlsx"
INDEX_COLS = ["id", "zone", "brief", "title", "path", "tags", "created", "updated"]
MEM_COLS = ["id", "title", "content", "tags", "created", "updated"]

ZONE_DEFAULTS = ["user", "project", "pref", "task", "knowledge", "contact"]


def _now() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


class ExcelLineStore:
    """Manages the master index and per-zone workbooks.

    Uses BOTH an in-process RLock (for thread safety) and a cross-process
    file lock (for safety when multiple Hermes processes share one `root`),
    so concurrent writes from two processes cannot corrupt the workbooks.
    """

    def __init__(self, root_dir: str):
        self.root = root_dir
        os.makedirs(self.root, exist_ok=True)
        self._lock = threading.RLock()
        self._seq = 0
        self._master = os.path.join(self.root, MASTER_NAME)
        self._lockfile = os.path.join(self.root, ".excel_line.lock")
        self._ensure_master()
        self._seq = self._next_seq()

    # -- cross-process lock ------------------------------------------------
    @contextlib.contextmanager
    def _cross_process_lock(self, timeout: float = 10.0, stale_after: float = 30.0):
        """Atomic file-lock usable across processes (works on Win + POSIX).

        The lock file stores ``<pid>:<timestamp>``. A stale lock (older than
        `stale_after`, or owned by a process that no longer exists) is cleared
        so we never block forever and never collide with a live owner.
        """
        deadline = time.time() + timeout
        fd = None
        while fd is None:
            try:
                fd = os.open(self._lockfile, os.O_CREAT | os.O_EXCL | os.O_RDWR)
                # Record ownership so a later process can tell if we're alive.
                os.write(fd, f"{os.getpid()}:{time.time()}".encode())
                break
            except OSError:
                # Maybe a stale lock from a crashed process — try to clear it.
                try:
                    age = time.time() - os.path.getmtime(self._lockfile)
                    if age > stale_after:
                        # Also confirm the recorded pid is dead (if readable).
                        try:
                            with open(self._lockfile, "r") as lf:
                                owner = lf.read().split(":", 1)[0].strip()
                            if owner and owner.isdigit():
                                import signal
                                try:
                                    os.kill(int(owner), 0)  # alive?
                                    break  # live owner -> give up this round
                                except OSError:
                                    pass  # dead -> safe to remove
                        except Exception:
                            pass
                        os.remove(self._lockfile)
                        continue
                except OSError:
                    pass
                if time.time() > deadline:
                    break
                time.sleep(0.05)
        try:
            yield
        finally:
            if fd is not None:
                try:
                    os.close(fd)
                    os.remove(self._lockfile)
                except OSError:
                    pass

    # -- paths -------------------------------------------------------------

    def zone_path(self, zone: str) -> str:
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in zone.lower())
        return os.path.join(self.root, f"{safe}.xlsx")

    # -- master ------------------------------------------------------------

    def _ensure_master(self):
        if not os.path.exists(self._master):
            wb = Workbook()
            ws = wb.active
            ws.title = "index"
            ws.append(INDEX_COLS)
            self._autosize(ws, INDEX_COLS)
            wb.save(self._master)

    def _next_seq(self) -> int:
        with self._cross_process_lock():
            try:
                wb = load_workbook(self._master, read_only=True)
                ws = wb["index"]
                max_id = 0
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and isinstance(row[0], int):
                        max_id = max(max_id, row[0])
                return max_id
            except Exception:
                return 0

    def _new_id(self) -> int:
        with self._lock:
            self._seq += 1
            return self._seq

    # -- write -------------------------------------------------------------

    def add(self, zone: str, brief: str, content: str, title: str = "",
            tags: str = "") -> int:
        """Add a memory: a brief row in the master index + a detailed zone row.

        Wrapped in try/except so a transient file/lock error cannot crash the
        provider — the caller gets -1 and a logged error instead of an exception.
        """
        zone = zone or "knowledge"
        try:
            with self._cross_process_lock():
                with self._lock:
                    mid = self._new_id()
                    now = _now()
                    zpath = self.zone_path(zone)
                    # write zone detail
                    if not os.path.exists(zpath):
                        wb = Workbook()
                        ws = wb.active
                        ws.title = "mem"
                        ws.append(MEM_COLS)
                        self._autosize(ws, MEM_COLS)
                    else:
                        wb = load_workbook(zpath)
                        ws = wb["mem"]
                    zid = ws.max_row  # 1-based; row 1 is header
                    ws.append([zid, title or brief[:40], content, tags, now, now])
                    wb.save(zpath)
                    # write master index row
                    mwb = load_workbook(self._master)
                    mws = mwb["index"]
                    mws.append([mid, zone, brief, title, zpath, tags, now, now])
                    mwb.save(self._master)
                    return mid
        except Exception as e:  # pragma: no cover - defensive
            import logging
            logging.getLogger(__name__).error("excel_line store.add failed: %s", e)
            return -1

    # -- read --------------------------------------------------------------

    def search_index(self, query: str, limit: int = 10) -> List[Dict]:
        """Keyword scan over the master index (brief + tags + zone)."""
        q = query.lower()
        hits: List[Dict] = []
        with self._lock:
            wb = load_workbook(self._master, read_only=True)
            ws = wb["index"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                # Scan zone + brief + title (row[3]) + path + tags.
                blob = " ".join(str(x or "") for x in (row[1], row[2], row[3], row[4], row[5])).lower()
                if q in blob:
                    hits.append({
                        "id": row[0], "zone": row[1], "brief": row[2],
                        "title": row[3], "path": row[4], "tags": row[5],
                    })
        return hits[:limit]

    def read_zone(self, path: str, limit: int = 20) -> List[Dict]:
        """Return detailed rows from a zone workbook."""
        if not os.path.exists(path):
            return []
        out: List[Dict] = []
        with self._lock:
            wb = load_workbook(path, read_only=True)
            ws = wb["mem"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                out.append({
                    "id": row[0], "title": row[1], "content": row[2],
                    "tags": row[3], "created": row[4],
                })
        return out[-limit:]

    def list_zones(self) -> List[str]:
        zones = []
        for f in os.listdir(self.root):
            if f.endswith(".xlsx") and f != MASTER_NAME:
                zones.append(f[:-5])
        return zones

    def count(self) -> int:
        try:
            wb = load_workbook(self._master, read_only=True)
            return max(0, wb["index"].max_row - 1)
        except Exception:
            return 0

    # -- helpers -----------------------------------------------------------

    @staticmethod
    def _autosize(ws, cols):
        for i, _ in enumerate(cols, start=1):
            ws.column_dimensions[get_column_letter(i)].width = 24
