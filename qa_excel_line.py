"""qa_excel_line.py — quick smoke test for the excel_line memory provider.

INTRODUCTION
    A standalone, dependency-light check you run AFTER Hermes has reloaded the
    plugin (restart Hermes, or it picks up the edited __init__.py / worker.py on
    next launch). It verifies the three behaviours that make excel_line behave
    like built-in memory:

      1. Direct-store mode writes immediately (no LLM / indexer dependency).
      2. search_index / the `search` tool retrieves what was stored.
      3. prefetch() (the auto-retrieve hook) returns the stored brief.

    This is the "fast confidence" script; the full suite lives in
    tests/test_excel_line.py.

USAGE
    cd <hermes>/plugins/excel_line
    uv run --with openpyxl python qa_excel_line.py
"""
from __future__ import annotations
import os, sys, tempfile, types, json

# --- stub Hermes runtime modules so the plugin imports standalone ---
def _stub():
    agent = types.ModuleType("agent"); mem = types.ModuleType("agent.memory_provider")
    class MemoryProvider:
        name = "base"
    mem.MemoryProvider = MemoryProvider; agent.memory_provider = mem
    sys.modules["agent"] = agent; sys.modules["agent.memory_provider"] = mem
    tools = types.ModuleType("tools"); reg = types.ModuleType("tools.registry")
    reg.tool_error = lambda m: json.dumps({"error": m}); tools.registry = reg
    sys.modules["tools"] = tools; sys.modules["tools.registry"] = reg

def main():
    _stub()
    sys.path.insert(0, os.path.dirname(__file__))
    import importlib.util
    spec = importlib.util.spec_from_file_location("xl", "__init__.py")
    xl = importlib.util.module_from_spec(spec); spec.loader.exec_module(xl)
    from store import ExcelLineStore

    cfg = {"root": tempfile.mkdtemp(), "log_dir": tempfile.mkdtemp()}
    prov = xl.ExcelLineProvider(cfg)
    prov._root = cfg["root"]; prov._log_dir = cfg["log_dir"]
    prov._store = ExcelLineStore(cfg["root"]); prov._session_id = "qa"

    ok = True
    # 1. direct-store
    r = prov._handle_add({"zone": "skill", "brief": "QA probe skill",
                          "content": "prove direct store works", "title": "qa", "tags": "qa"})
    if json.loads(r).get("status") != "stored":
        print("FAIL: direct-store did not return 'stored'"); ok = False
    # 2. search retrieves
    if not prov._store.search_index("QA probe skill"):
        print("FAIL: search_index found nothing after direct store"); ok = False
    # 3. prefetch
    pf = prov.prefetch("QA probe skill", session_id="qa")
    if "index matches" not in pf:
        print("FAIL: prefetch returned no matches"); ok = False

    print("QA RESULT:", "PASS" if ok else "FAIL")
    print("count:", prov._store.count(), "| prompt_block head:",
          prov.system_prompt_block().splitlines()[0])
    sys.exit(0 if ok else 1)

if __name__ == "__main__":
    main()
