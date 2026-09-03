"""QA v2 đầu cuối: provider + BrainStore + rotation + worker tri-state.
Chạy trong stub runtime như test suite chính."""
import importlib, json, os, re, sys, tempfile, types
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, r"C:\Users\Admin\AppData\Local\hermes\hermes-agent")
sys.path.insert(0, r"C:\Users\Admin\AppData\Local\hermes\plugins")

stub_src = open(r"C:\Users\Admin\AppData\Local\hermes\plugins\excel_line\tests\test_excel_line.py", encoding="utf-8").read()
m = re.search(r"def _stub_runtime\(\):.*?(?=\ndef )", stub_src, re.S)
ns = {}; exec("import os,sys,json,types\n" + m.group(0), ns); ns["_stub_runtime"]()

spec = importlib.util.spec_from_file_location(
    "excel_line", r"C:\Users\Admin\AppData\Local\hermes\plugins\excel_line\__init__.py",
    submodule_search_locations=[r"C:\Users\Admin\AppData\Local\hermes\plugins\excel_line"])
xl = importlib.util.module_from_spec(spec); sys.modules["excel_line"] = xl
spec.loader.exec_module(xl)

root = tempfile.mkdtemp()
prov = xl.ExcelLineProvider({"root": root, "log_dir": os.path.join(root, "logs"),
                             "free_model": "x"})
prov.initialize("qa-v2")
T = []
def check(name, cond, extra=""):
    T.append((name, bool(cond)))
    print(("PASS " if cond else "FAIL ") + name + (" | " + str(extra)[:140] if extra else ""))

# 1. store là BrainStore?
check("provider dùng BrainStore", type(prov._store).__name__ == "BrainStore", type(prov._store).__name__)

# 2. add tree mode
r = json.loads(prov._handle({"action": "add", "branch": "brain.xlsx",
                             "title": "test v2", "content": "nén gọn", "tags": "t"}))
check("add → stored", r.get("status") == "stored", r)

# 3. read tree
r = json.loads(prov._handle({"action": "read", "branch": "brain.xlsx"}))
check("read có usage", r.get("usage") == "1/10", r)

# 4. tree
r = json.loads(prov._handle({"action": "tree"}))
check("tree render", "brain.xlsx" in r.get("tree", ""), r.get("tree", "")[:80])

# 5. child
r = json.loads(prov._handle({"action": "child", "parent": "brain.xlsx",
                             "name": "sub1", "title": "nhánh con"}))
check("child → created", r.get("status") == "created" and r.get("file") == "sub1.xlsx", r)

# 6. overflow: đôn 9 dòng nữa vào brain.xlsx rồi thêm dòng 11 → branch_full
for n in range(8):
    prov._handle({"action": "add", "branch": "brain.xlsx", "title": f"m{n}", "content": "x"})
r = json.loads(prov._handle({"action": "add", "branch": "brain.xlsx", "title": "tràn", "content": "y"}))
check("đầy → branch_full", r.get("status") == "branch_full", r)

# 7. search toàn cây
r = json.loads(prov._handle({"action": "search", "query": "nén gọn"}))
check("search thấy", r.get("count", 0) >= 1 and r["results"][0]["title"] == "test v2", r["results"][:1])

# 8. merge
r0 = json.loads(prov._handle({"action": "search", "query": "m3"}))
mid = r0["results"][0]["id"]
r1 = json.loads(prov._handle({"action": "search", "query": "m4"}))
rid = r1["results"][0]["id"]
r = json.loads(prov._handle({"action": "merge", "branch": "brain.xlsx",
                             "ids": [mid, rid], "title": "gộp m3m4", "content": "x"}))
check("merge → id mới", r.get("status") == "merged", r)

# 9. update + delete
r = json.loads(prov._handle({"action": "update", "row_id": r["id"], "title": "gộp đã sửa", "zone": ""}))
check("update ok", r.get("status") == "updated", r)
r0 = json.loads(prov._handle({"action": "search", "query": "m0"}))
r = json.loads(prov._handle({"action": "delete", "row_id": r0["results"][0]["id"]}))
check("delete ok", r.get("status") == "deleted", r)

# 10. legacy zone add vẫn chạy
r = json.loads(prov._handle({"action": "add", "zone": "user", "brief": "legacy",
                             "content": "vẫn lưu được", "title": "legacy", "tags": "z"}))
check("legacy zone-add", r.get("status") == "stored", r)

# 11. external leaf pointer
ext = os.path.join(root, "..", "external_asset.txt")
ext = os.path.abspath(ext)
open(ext, "w").write("skill file nơi khác sống")
r = json.loads(prov._handle({"action": "add", "branch": "brain.xlsx", "title": "trỏ ngoài",
                             "content": "lá tự do", "tags": "t", "link": ext.replace("\\", "/")}))
check("leaf ngoài được chấp nhận", r.get("status") == "stored", r)
r = json.loads(prov._handle({"action": "read", "branch": "brain.xlsx"}))
exts = [x for x in r["rows"] if str(x.get("branch") or "").lower().endswith(".txt")]
check("cuống lá trỏ đúng file ngoài", exts and os.path.exists(exts[0]["branch"].replace("\\", "/")), exts[:1])

# 12. lá ngoài KHÔNG được xóa khi delete memory
pid = exts[0]["id"] if exts else None
prov._handle({"action": "delete", "row_id": pid})
check("xóa memory → file ngoài còn nguyên", os.path.exists(ext))

# 13. worker + curator tri-state với tree prompt
from excel_line.worker import _classify_tristate, _apply
calls = {"n": 0}
def fn(prompt):
    calls["n"] += 1
    assert "CURRENT TREE" in prompt and "action" in prompt, "prompt phải chứa cây"
    return json.dumps({"action": "add", "branch": "brain.xlsx",
                       "title": "fact từ curator", "content": "nén bởi LLM", "tags": "cur"})
st, dec = _classify_tristate({"input": "x", "output": "y"}, fn, prov._store.tree_text())
check("worker prompt có cây + action", st == "ok" and dec["action"] == "add", st)
got = _apply(prov._store, dec, {})
check("_apply add vào cây", got == 1)
dec2 = {"action": "none"}
check("none = 0 stored", _apply(prov._store, dec2, {}) == 0)

# 14. count/list_zones
check("count > 0", prov._store.count() > 0, prov._store.count())
check("zones liệt kê file", "brain" in prov._store.list_zones(), prov._store.list_zones())

print(f"\n{sum(1 for _, ok in T if ok)}/{len(T)} passed")
sys.exit(0 if all(ok for _, ok in T) else 1)
