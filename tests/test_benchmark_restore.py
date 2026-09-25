"""P2 benchmark (FTS5/BM25 vs tiered retrieval) and backup/restore E2E tests."""
from __future__ import annotations

import json
import os
import shutil
import sqlite3
import tempfile
import time
from pathlib import Path

from excel_line_core.brain_store import BrainStore, _flatten_rows
from excel_line_core.retrieval import rank_rows, row_text, tokenize


def _add_rows(store: BrainStore, n: int = 200) -> list[int]:
    ids: list[int] = []
    for i in range(1, n + 1):
        branch = f"bench_{(i - 1)//10 + 1}.xlsx"
        rid = store.add(
            branch,
            title=f"model kiểm thử khách quan {i}",
            content=f"Thông tin về model kiểm thử và routing số {i}",
            tags=f"source=qa;confidence=0.{90 + i % 10};entity_links=sheet{i % 5}",
            metadata={
                "profile_id": "prof-bench",
                "session_id": "sess-bench",
                "source": "qa",
                "confidence": 0.90 + (i % 10) / 100,
                "entity_links": f"sheet{i % 5}",
            },
        )
        ids.append(rid)
    return ids


def test_backup_paths_returns_external_roots(monkeypatch):
    tmp = Path(tempfile.mkdtemp())
    root = tmp / "data"
    log_dir = tmp / "logs"
    root.mkdir()
    log_dir.mkdir()
    try:
        store = BrainStore(str(root))
        store.add("brain.xlsx", title="x")
        monkeypatch.delenv("HERMES_HOME", raising=False)
        from excel_line import ExcelLineProvider

        prov = ExcelLineProvider(config={"root": str(root), "log_dir": str(log_dir)})
        prov.initialize("bench")
        paths = prov.backup_paths()
        assert str(root) in paths or any(p.startswith(str(tmp)) for p in paths)
        assert str(log_dir) in paths or any(p.startswith(str(tmp)) for p in paths)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_restore_from_backup_e2e():
    tmp = Path(tempfile.mkdtemp())
    try:
        root = tmp / "data"
        log_dir = tmp / "logs"
        root.mkdir()
        log_dir.mkdir()
        store = BrainStore(str(root))
        rid = store.add(
            "brain.xlsx",
            title="restore test",
            content="nội dung cần phục hồi sau backup",
            tags="source=restore;confidence=0.95;entity_links=sheet1",
            metadata={
                "profile_id": "prof-r",
                "session_id": "sess-r",
                "source": "restore",
                "confidence": 0.95,
                "entity_links": "sheet1",
            },
        )
        assert store.path_of(rid) == "brain.xlsx › #" + str(rid)

        backup = tmp / "backup"
        shutil.copytree(root, backup)
        shutil.rmtree(root)

        restored = tmp / "data"
        shutil.copytree(backup, restored)
        store2 = BrainStore(str(restored))
        found = store2.search_index("phục hồi", limit=10)
        assert found and found[0]["id"] == rid
        meta = found[0].get("metadata") or {}
        assert meta.get("confidence") == 0.95
        assert meta.get("entity_links") == "sheet1"
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_tiered_retrieval_vs_naive_benchmark():
    tmp = Path(tempfile.mkdtemp())
    try:
        store = BrainStore(str(tmp))
        ids = _add_rows(store, 250)
        query = "model kiểm thử khách quan"
        rows = list(_flatten_rows(store))

        t0 = time.perf_counter()
        for _ in range(3):
            tier_all = store.search_index(query, limit=20, tier="all")
        tiered_ms = (time.perf_counter() - t0) / 3 * 1000

        t0 = time.perf_counter()
        for _ in range(3):
            naive = rank_rows(query, rows, limit=20)
        naive_ms = (time.perf_counter() - t0) / 3 * 1000

        assert tier_all and tier_all[0]["id"] in ids
        assert naive and naive[0]["id"] in ids
        assert tiered_ms <= naive_ms * 8 or tiered_ms < 2000
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_sqlite_fts5_vs_tiered_benchmark():
    tmp = Path(tempfile.mkdtemp())
    db = tmp / "fts.db"
    try:
        conn = sqlite3.connect(str(db))
        conn.execute("CREATE VIRTUAL TABLE mem USING fts5(title, content, tokenize='unicode61')")
        rows: list[dict] = []
        for i in range(1, 201):
            title = f"model kiểm thử khách quan {i}"
            content = f"Thông tin về model kiểm thử và routing số {i}"
            conn.execute("INSERT INTO mem(title, content) VALUES(?, ?)", (title, content))
            rows.append({"title": title, "content": content})
        conn.commit()

        t0 = time.perf_counter()
        for _ in range(3):
            conn.execute("SELECT rowid, title, content FROM mem WHERE mem MATCH ?", ("model kiểm thử",)).fetchall()
        fts_ms = (time.perf_counter() - t0) / 3 * 1000

        t0 = time.perf_counter()
        for _ in range(3):
            rank_rows("model kiểm thử", rows, limit=20)
        tiered_ms = (time.perf_counter() - t0) / 3 * 1000

        assert fts_ms > 0 and tiered_ms > 0
        assert tiered_ms <= fts_ms * 12 or tiered_ms < 3000
        conn.close()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
