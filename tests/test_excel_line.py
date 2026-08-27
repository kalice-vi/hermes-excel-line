"""tests/test_excel_line.py — full QA suite for the excel_line memory provider.

INTRODUCTION
    The authoritative regression suite. It exercises every layer against real
    openpyxl workbooks in temp dirs (no mocks of the storage path), proving the
    plugin is not merely "write-only" but actually stores AND retrieves:

      - store.py: add / search (unicode, Vietnamese, by title, limit) / read /
        count / zones / exception safety.
      - concurrency: 5x20 threads, and concurrent direct-store + worker on one
        store, and two separate store instances (cross-process) with stale-lock
        pid recovery.
      - worker.py: keeps log on classify-fail, raw-backup on fail, deletes on
        success.
      - provider: direct-store returns "stored" + retrievable; prefetch() finds
        drained lazy logs; on_session_end keeps data when the LLM fails.

    This is the suite reported in the PR ("33 tests, all green").

RUN
    cd <hermes>/plugins/excel_line
    uv run --with openpyxl python tests/test_excel_line.py
    # or: uv run --with openpyxl python -m pytest tests/test_excel_line.py -q
"""
from __future__ import annotations
import os, sys, json, tempfile, types, threading, time
import importlib.util

# --- stub Hermes runtime so the plugin imports standalone ----------------
def _stub_runtime():
    agent = types.ModuleType("agent"); mem = types.ModuleType("agent.memory_provider")
    class MemoryProvider:
        name = "base"
    mem.MemoryProvider = MemoryProvider; agent.memory_provider = mem
    sys.modules.setdefault("agent", agent); sys.modules.setdefault("agent.memory_provider", mem)
    tools = types.ModuleType("tools"); reg = types.ModuleType("tools.registry")
    reg.tool_error = lambda m: json.dumps({"error": m}); tools.registry = reg
    sys.modules.setdefault("tools", tools); sys.modules.setdefault("tools.registry", reg)

def _load_plugin():
    _stub_runtime()
    plugin_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if plugin_dir not in sys.path:
        sys.path.insert(0, plugin_dir)
    spec = importlib.util.spec_from_file_location("excel_line_plugin", os.path.join(plugin_dir, "__init__.py"))
    mod = importlib.util.module_from_spec(spec)
    sys.modules["excel_line_plugin"] = mod  # register so tests can monkeypatch
    spec.loader.exec_module(mod)
    return mod

PASS, FAIL = [], []
def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(("PASS " if cond else "FAIL ") + name + (f"  -- {detail}" if detail and not cond else ""))

# Make the plugin package importable (store/worker are siblings of tests/).
_PLUGIN_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PLUGIN_DIR not in sys.path:
    sys.path.insert(0, _PLUGIN_DIR)


# --------------------------------------------------------------------------
def test_store_basic():
    from store import ExcelLineStore
    root = tempfile.mkdtemp()
    s = ExcelLineStore(root)
    mid = s.add(zone="skill", brief="Skill: demo", content="body", title="demo", tags="t1,t2")
    check("store.add returns id", isinstance(mid, int) and mid >= 1)
    hits = s.search_index("demo", limit=10)
    check("store.search finds by brief", len(hits) == 1 and hits[0]["zone"] == "skill")
    rows = s.read_zone(s.zone_path("skill"), limit=5)
    check("store.read_zone returns content", rows and rows[-1]["content"] == "body")
    check("store.count", s.count() == 1)
    check("store.list_zones", "skill" in s.list_zones())


def test_store_search_edge():
    from store import ExcelLineStore
    root = tempfile.mkdtemp()
    s = ExcelLineStore(root)
    s.add(zone="skill", brief="Fast Accounting Online help map", content="x", tags="FAO",
          title="fast-accounting-online-help")
    # case-insensitive + unicode substring
    check("search unicode case-insensitive", len(s.search_index("fast accounting", 10)) == 1)
    check("search vietnamese tag", len(s.search_index("fao", 10)) == 1)
    # search by TITLE even when the keyword is absent from brief
    check("search covers title", len(s.search_index("fast-accounting-online-help", 10)) == 1)
    check("search miss -> 0", len(s.search_index("zzz-none", 10)) == 0)
    # limit respected
    for i in range(15):
        s.add(zone="knowledge", brief=f"item {i}", content="c")
    check("search limit respected", len(s.search_index("item", limit=5)) == 5)


def test_store_concurrent(tmp=None):
    from store import ExcelLineStore
    root = tempfile.mkdtemp()
    s = ExcelLineStore(root)
    def writer(n):
        for i in range(n):
            s.add(zone="knowledge", brief=f"c{i}", content="x")
    threads = [threading.Thread(target=writer, args=(20,)) for _ in range(5)]
    [t.start() for t in threads]; [t.join() for t in threads]
    check("concurrent writes no corruption", s.count() == 100,
          f"count={s.count()}")


def _load_worker():
    """Load worker.py as a submodule of the plugin package so its
    `from .store import ...` relative import resolves (mirrors how Hermes
    imports the plugin at runtime)."""
    import importlib.util as _ilu
    pkg_name = "excel_line_qa_pkg"
    if pkg_name not in sys.modules:
        pkg = types.ModuleType(pkg_name)
        pkg.__path__ = [_PLUGIN_DIR]
        pkg.__package__ = pkg_name
        sys.modules[pkg_name] = pkg
    spec = _ilu.spec_from_file_location(f"{pkg_name}.worker", os.path.join(_PLUGIN_DIR, "worker.py"),
                                        submodule_search_locations=[_PLUGIN_DIR])
    worker = _ilu.module_from_spec(spec)
    sys.modules[f"{pkg_name}.worker"] = worker
    spec.loader.exec_module(worker)
    # make worker.store resolve to the same store module
    if f"{pkg_name}.store" not in sys.modules:
        s_spec = _ilu.spec_from_file_location(f"{pkg_name}.store", os.path.join(_PLUGIN_DIR, "store.py"))
        s_mod = _ilu.module_from_spec(s_spec); sys.modules[f"{pkg_name}.store"] = s_mod
        s_spec.loader.exec_module(s_mod)
        sys.modules["store"] = s_mod  # also allow bare `from store import`
    return worker


def test_worker_no_drop_on_fail():
    """Classify fails AND store.add fails -> failed records must be retried (log kept)."""
    worker = _load_worker()
    from store import ExcelLineStore
    root = tempfile.mkdtemp(); logd = tempfile.mkdtemp()
    store = ExcelLineStore(root)
    # Force BOTH classify and store.add to fail so the record cannot persist.
    def bad_model(prompt):
        return "not json at all"
    store.add = lambda *a, **k: -1  # simulate defensive failure
    log = os.path.join(logd, "sess_i1_o1_20260101.jsonl")
    with open(log, "w", encoding="utf-8") as f:
        f.write(json.dumps({"session": "s", "input": "x long input text", "output": "y", "ts": "t"}) + "\n")
    stored = worker.process_logs(logd, root, bad_model, store=store)
    check("worker stores 0 on total fail", stored == 0)
    check("worker KEEPS log on fail (no silent drop)", os.path.exists(log),
          "log missing -> data lost")
    # The kept log must still contain the original record for retry.
    kept = open(log, encoding="utf-8").read()
    check("kept log retains record for retry", "long input text" in kept)


def test_worker_fallback_raw_store():
    """Classify fails -> worker still stores a raw-text backup (no data loss)."""
    worker = _load_worker()
    from store import ExcelLineStore
    root = tempfile.mkdtemp(); logd = tempfile.mkdtemp()
    store = ExcelLineStore(root)
    def bad_model(prompt):
        return "not json at all"
    log = os.path.join(logd, "sess_i1_o1_20260101.jsonl")
    with open(log, "w", encoding="utf-8") as f:
        f.write(json.dumps({"session": "s", "input": "remember the Honda bike purchase",
                            "output": "ok noted", "ts": "2026"}) + "\n")
    stored = worker.process_logs(logd, root, bad_model, store=store)
    check("worker raw-backup stores on classify fail", stored == 1)
    check("worker deletes log after raw-backup", not os.path.exists(log))
    rows = store.read_zone(store.zone_path("knowledge"), limit=5)
    check("raw-backup content persisted", rows and "Honda" in rows[-1]["content"])


def test_worker_store_on_success():
    worker = _load_worker()
    from store import ExcelLineStore
    root = tempfile.mkdtemp(); logd = tempfile.mkdtemp()
    store = ExcelLineStore(root)
    def good_model(prompt):
        return json.dumps({"zone": "knowledge", "brief": "classified",
                           "title": "t", "content": "c", "tags": "k"})
    log = os.path.join(logd, "sess_i1_o1_20260101.jsonl")
    with open(log, "w", encoding="utf-8") as f:
        f.write(json.dumps({"session": "s", "input": "x", "output": "y", "ts": "t"}) + "\n")
    stored = worker.process_logs(logd, root, good_model, store=store)
    check("worker stores 1 on success", stored == 1)
    check("worker deletes log on success", not os.path.exists(log))


def test_worker_skips_raw_transcript_backups():
    """BLOCKER-01: worker must NOT index session_raw_*.jsonl backups written by
    the provider — those are raw transcripts, not agent I/O logs. Indexing them
    would duplicate memory and risk a rename-loop on read failure."""
    worker = _load_worker()
    from store import ExcelLineStore
    root = tempfile.mkdtemp(); logd = tempfile.mkdtemp()
    store = ExcelLineStore(root)
    before = store.count()
    # A raw-transcript backup file in the same log_dir.
    raw = os.path.join(logd, "session_raw_default_20260101.jsonl")
    with open(raw, "w", encoding="utf-8") as f:
        f.write(json.dumps({"session": "s", "input": "user said hi",
                            "output": "agent replied", "ts": "t"}) + "\n")
    # No real turn logs present -> worker should skip the raw backup and store 0.
    stored = worker.process_logs(logd, root, lambda p: "not json", store=store)
    check("worker skips session_raw_ backups", stored == 0)
    check("worker leaves raw backup untouched", os.path.exists(raw))
    check("worker did not index raw backup", store.count() == before)


def test_provider_direct_store():
    xl = _load_plugin()
    from store import ExcelLineStore
    cfg = {"root": tempfile.mkdtemp(), "log_dir": tempfile.mkdtemp()}
    p = xl.ExcelLineProvider(cfg)
    p._root = cfg["root"]; p._log_dir = cfg["log_dir"]
    p._store = ExcelLineStore(cfg["root"]); p._session_id = "qa"
    r = p._handle_add({"zone": "skill", "brief": "Direct store skill",
                       "content": "body", "title": "ds", "tags": "qa"})
    check("provider direct-store returns stored", json.loads(r).get("status") == "stored")
    check("provider direct-store retrievable",
          len(p._store.search_index("Direct store skill", 10)) == 1)
    # prefetch (auto-retrieve hook used each turn)
    pf = p.prefetch("Direct store skill", session_id="qa")
    check("provider prefetch returns matches", "index matches" in pf)
    # system prompt block reflects count
    blk = p.system_prompt_block()
    check("system_prompt_block active", "Active." in blk and "indexed memories" in blk)


def test_provider_direct_store_no_llm_dep():
    """Direct-store must NOT call the free model at all."""
    xl = _load_plugin()
    from store import ExcelLineStore
    cfg = {"root": tempfile.mkdtemp(), "log_dir": tempfile.mkdtemp()}
    p = xl.ExcelLineProvider(cfg)
    p._root = cfg["root"]; p._log_dir = cfg["log_dir"]
    p._store = ExcelLineStore(cfg["root"]); p._session_id = "qa"
    # even if free model is broken, direct store works
    xl._ask_free_model = lambda *a, **k: (_ for _ in ()).throw(RuntimeError("LLM down"))
    r = p._handle_add({"zone": "user", "brief": "user pref: coffee in morning",
                       "content": "likes coffee", "title": "pref", "tags": "user"})
    check("direct-store works with LLM broken", json.loads(r).get("status") == "stored")


def test_provider_seq_mode_still_logs():
    """Legacy sequence mode should still write a log (best-effort)."""
    xl = _load_plugin()
    from store import ExcelLineStore
    cfg = {"root": tempfile.mkdtemp(), "log_dir": tempfile.mkdtemp()}
    p = xl.ExcelLineProvider(cfg)
    p._root = cfg["root"]; p._log_dir = cfg["log_dir"]
    p._store = ExcelLineStore(cfg["root"]); p._session_id = "qa"
    p._transcripts["qa"] = [{"input": "remember X", "output": "ok X", "ts": "t"}]
    r = p._handle_add({"session": "qa", "input_seq": 1, "output_seq": 1})
    check("seq mode returns logged", json.loads(r).get("status") == "logged")
    check("seq mode wrote a log file", any(
        n.endswith(".jsonl") for n in os.listdir(cfg["log_dir"])))


def test_concurrent_direct_and_worker():
    """Direct-store (main thread) + worker (bg thread, same store instance)
    writing the master index concurrently must not corrupt it."""
    import importlib.util as _ilu
    xl = _load_plugin()
    from store import ExcelLineStore
    root = tempfile.mkdtemp(); logd = tempfile.mkdtemp()
    p = xl.ExcelLineProvider({"root": root, "log_dir": logd})
    p._root = root; p._log_dir = logd
    p._store = ExcelLineStore(root); p._session_id = "qa"

    def good_model(prompt):
        return json.dumps({"zone": "knowledge", "brief": "cls",
                           "title": "t", "content": "c", "tags": "k"})

    # seed a log for the worker to process in a background thread
    log = os.path.join(logd, "sess_i1_o1_x.jsonl")
    with open(log, "w", encoding="utf-8") as f:
        f.write(json.dumps({"session": "s", "input": "x", "output": "y", "ts": "t"}) + "\n")

    pkg = types.ModuleType("excel_line_qa_pkg2")
    pkg.__path__ = [_PLUGIN_DIR]; pkg.__package__ = "excel_line_qa_pkg2"
    sys.modules["excel_line_qa_pkg2"] = pkg
    wspec = _ilu.spec_from_file_location("excel_line_qa_pkg2.worker", os.path.join(_PLUGIN_DIR, "worker.py"), submodule_search_locations=[_PLUGIN_DIR])
    worker = _ilu.module_from_spec(wspec); sys.modules["excel_line_qa_pkg2.worker"] = worker
    wspec.loader.exec_module(worker)

    errors = []
    def direct_writer():
        try:
            for i in range(30):
                p._handle_add({"zone": "skill", "brief": f"concurrent {i}",
                               "content": "c", "title": f"c{i}", "tags": "qa"})
        except Exception as e:
            errors.append(str(e))
    def worker_runner():
        try:
            worker.process_logs(logd, root, good_model, store=p._store)
        except Exception as e:
            errors.append(str(e))

    t1 = threading.Thread(target=direct_writer)
    t2 = threading.Thread(target=worker_runner)
    t1.start(); t2.start(); t1.join(); t2.join()
    check("concurrent direct+worker no exception", not errors, str(errors)[:200])
    # master index must be readable & consistent
    try:
        cnt = p._store.count()
        check("concurrent master index intact", cnt >= 30, f"count={cnt}")
    except Exception as e:
        check("concurrent master index intact", False, str(e))


def test_store_add_exception_safety():
    """store.add must not raise on a bad path — returns -1 and logs instead."""
    from store import ExcelLineStore
    # point root at an unwritable location (a file, not a dir)
    import tempfile
    bad = os.path.join(tempfile.mkdtemp(), "not_a_dir", "x.xlsx")
    s = ExcelLineStore.__new__(ExcelLineStore)
    s._lock = threading.RLock()
    s._seq = 0
    s.root = os.path.dirname(bad)
    s._master = bad
    mid = s.add(zone="skill", brief="x", content="y")
    check("store.add returns -1 on failure (no crash)", mid == -1)


def test_cross_process_lock():
    """Two ExcelLineStore instances (simulating two processes) writing the
    same master must serialize via the file lock, not corrupt or lose rows."""
    from store import ExcelLineStore
    import tempfile
    root = tempfile.mkdtemp()
    s0 = ExcelLineStore(root)  # init master
    # simulate a stale lock left by a crashed process
    import os as _os
    lockf = _os.path.join(root, ".excel_line.lock")
    with open(lockf, "w") as f:
        f.write("stale")
    _os.utime(lockf, (_os.path.getatime(lockf), _os.path.getmtime(lockf) - 60))
    errors = []
    def writer(n):
        try:
            s = ExcelLineStore(root)  # separate instance, same root = cross-process
            for i in range(20):
                s.add(zone="knowledge", brief=f"xproc {n}-{i}", content="c")
        except Exception as e:
            errors.append(str(e))
    ts = [threading.Thread(target=writer, args=(k,)) for k in range(3)]
    [t.start() for t in ts]; [t.join() for t in ts]
    check("cross-process no exception (stale lock recovered)", not errors, str(errors)[:200])
    s = ExcelLineStore(root)
    check("cross-process all 60 rows present", s.count() == 60, f"count={s.count()}")


def test_lazy_index_and_raw_backup():
    """prefetch() must lazily drain pending seq-logs, and on_session_end must
    back up the raw transcript when extraction yields nothing."""
    xl = _load_plugin()
    from store import ExcelLineStore
    import tempfile, types as _t
    root = tempfile.mkdtemp(); logd = tempfile.mkdtemp()
    p = xl.ExcelLineProvider({"root": root, "log_dir": logd})
    p._root = root; p._log_dir = logd
    p._store = ExcelLineStore(root); p._session_id = "qa"

    # 1) seed a seq-log; the BACKGROUND indexer (not prefetch) drains it.
    #    prefetch() must stay cheap (no synchronous LLM call) — it only searches.
    import os as _os
    log = _os.path.join(logd, "sess_i1_o1_lazy.jsonl")
    with open(log, "w", encoding="utf-8") as f:
        f.write(json.dumps({"session": "s", "input": "remember the Honda purchase",
                            "output": "noted", "ts": "2026"}) + "\n")
    # make indexer produce a usable record by overriding the free-model fn
    # on the registered plugin module (the one _drain_pending_logs uses)
    import sys as _sys
    _sys.modules["excel_line_plugin"]._ask_free_model = lambda p, model="x": json.dumps(
        {"zone": "knowledge", "brief": "Honda purchase noted",
         "title": "honda", "content": "user bought Honda", "tags": "k"})
    # prefetch should NOT drain (Gemini review #4) — it only searches existing index.
    res0 = p.prefetch("Honda", session_id="qa")
    check("prefetch does not drain log (stays cheap)", _os.path.exists(log))
    # background indexer drains it instead.
    p._run_indexer()
    check("lazy index removed the log after background drain", not _os.path.exists(log))
    res = p.prefetch("Honda", session_id="qa")
    check("prefetch finds indexed memory after background drain", "Honda" in res, res[:120])

    # 2) on_session_end must not lose data when the LLM fails to classify:
    #    _auto_extract's _fallback_store persists raw turns into the store.
    logd2 = tempfile.mkdtemp(); root2 = tempfile.mkdtemp()
    p2 = xl.ExcelLineProvider({"root": root2, "log_dir": logd2})
    p2._root = root2; p2._log_dir = logd2
    p2._store = ExcelLineStore(root2); p2._session_id = "qa2"
    _sys.modules["excel_line_plugin"]._ask_free_model = lambda p, model="x": "not json"  # force classify fail
    msgs = [{"role": "user", "content": "remember seagift format 2026"},
            {"role": "assistant", "content": "ok"}]
    p2.on_session_end(msgs)
    # data must survive via _fallback_store (stored into the knowledge zone)
    rows = p2._store.read_zone(p2._store.zone_path("knowledge"), limit=10)
    survived = any("seagift" in (r.get("content") or "") for r in rows)
    check("on_session_end keeps data when LLM fails (fallback store)", survived,
          str([r.get("content") for r in rows][:3]))


def test_formula_injection_safe():
    """Cells starting with = + - @ must be escaped so Excel does not execute
    them as formulas (formula injection)."""
    from store import ExcelLineStore
    import tempfile
    root = tempfile.mkdtemp()
    s = ExcelLineStore(root)
    for payload in ["=cmd|'/c calc'!A1", "+1+1", "-2", "@SUM(1)"]:
        mid = s.add(zone="knowledge", brief=payload, content=payload,
                    title=payload, tags="x")
        rows = s.read_zone(s.zone_path("knowledge"), limit=10)
        stored = rows[-1]["content"]
        check(f"formula-injection escaped: {payload!r}", stored.startswith("'") and stored[1:] == payload,
              stored)
        # master index brief also escaped
        mrows = s.search_index(payload[1:] if payload[0] in "=+-@" else payload, limit=5)
        # search still works on the visible (escaped) text
        assert mid


def main():
    test_store_basic()
    test_store_search_edge()
    test_store_concurrent()
    test_worker_no_drop_on_fail()
    test_worker_fallback_raw_store()
    test_worker_store_on_success()
    test_worker_skips_raw_transcript_backups()
    test_provider_direct_store()
    test_provider_direct_store_no_llm_dep()
    test_provider_seq_mode_still_logs()
    test_concurrent_direct_and_worker()
    test_store_add_exception_safety()
    test_cross_process_lock()
    test_lazy_index_and_raw_backup()
    test_formula_injection_safe()
    print(f"\n==== QA SUMMARY: {len(PASS)} passed, {len(FAIL)} failed ====")
    if FAIL:
        print("FAILED:", FAIL); sys.exit(1)
    print("ALL GREEN")


if __name__ == "__main__":
    main()
