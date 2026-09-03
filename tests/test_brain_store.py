"""Test brain_store v2: 10 dòng/file, node/lá, FullError, promote, move, merge."""
import os, sys, tempfile, traceback
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, r"C:\Users\Admin\AppData\Local\hermes\plugins")
import types
pkg = types.ModuleType("excel_line")
pkg.__path__ = [r"C:\Users\Admin\AppData\Local\hermes\plugins\excel_line"]
sys.modules["excel_line"] = pkg
from excel_line.brain_store import BrainStore, FullError, BadBranch, MAX_ROWS, MASTER_V2

T = []
def check(name, cond, extra=""):
    T.append((name, bool(cond)))
    print(("PASS " if cond else "FAIL ") + name + (" | " + str(extra)[:120] if extra else ""))

root = tempfile.mkdtemp()
s = BrainStore(root)

# 1. gốc tồn tại, trống
check("gốc brain.xlsx tạo sẵn", os.path.exists(os.path.join(root, "brain.xlsx")))
check("gốc 0 dòng", len(s.load_rows(MASTER_V2)) == 0)

# 2. add thường
i1 = s.add(MASTER_V2, title="user", content="hồ sơ người dùng", tags="profile")
i2 = s.add(MASTER_V2, title="knowledge", content="fact", tags="k")
check("add trả id tăng dần", i2 == i1 + 1)

# 3. Title/Content caps
i3 = s.add(MASTER_V2, title="x" * 90, content="y" * 500)
r3 = s.load_rows(MASTER_V2)[2]
check("title cắt ≤50", len(str(r3["title"])) <= 55, len(r3["title"]))
check("content cắt ≤250", len(str(r3["content"])) <= 260, len(r3["content"]))

# 4. link tới file chưa tồn tại → lỗi
try:
    s.add(MASTER_V2, title="bad", link="nothing.png"); ok = False
except BadBranch:
    ok = True
check("lá phải trỏ file có thật", ok)

# 5. child: tạo node .xlsx + dòng trỏ
res = s.child(MASTER_V2, "user", "Nhánh user")
check("child tạo user.xlsx", os.path.exists(os.path.join(root, "user.xlsx")))
check("child trỏ đúng", res["file"] == "user.xlsx")

# 6. FULL 10 dòng → FullError
for n in range(20):
    try:
        s.child(MASTER_V2, f"z{n}", f"zone {n}")
    except FullError:
        break
rows_root = s.load_rows(MASTER_V2)
check("gốc tối đa 10", len(rows_root) == MAX_ROWS, len(rows_root))
try:
    s.add(MASTER_V2, title="tràn"); ok = False
except FullError as e:
    ok = "10/10" in str(e)
check("đầy → FullError ép merge/child", ok)

# 7. node con cũng đầy 10
sub = s.child(MASTER_V2 if False else "user.xlsx", "a", "sub")  # user.xlsx mới trống
for n in range(9):
    try:
        s.add("user.xlsx", title=f"m{n}")
    except FullError:
        pass
try:
    s.add("user.xlsx", title="overflow"); ok = False
except FullError:
    ok = True
check("file nhánh cũng cap 10", ok)

# 8. move giữa file (dọn 1 slot gốc bằng rm — test FullError đã đầy)
z_last = max(r["id"] for r in s.load_rows(MASTER_V2))
s.rm(MASTER_V2, z_last)
dst = s.child(MASTER_V2, "pref", "nhánh pref")
moved = s.move(MASTER_V2, [i1], "pref.xlsx")
check("move 1 dòng", moved == 1 and len(s.load_rows(MASTER_V2)) < MAX_ROWS)

# 9. merge: 3 dòng → 1
c1 = s.add("pref.xlsx", title="c1", content="a")
c2 = s.add("pref.xlsx", title="c2", content="b")
c3 = s.add("pref.xlsx", title="c3", content="c")
before = len(s.load_rows("pref.xlsx"))
mid = s.merge("pref.xlsx", [c1, c2, c3], "c gộp", "nén 3 còn 1", "merge")
after = len(s.load_rows("pref.xlsx"))
check("merge giải phóng 2 slot", after == before - 2 and mid > 0, f"{before}->{after}")

# 10. promote lá .py → node .xlsx + ID không tái sử dụng sau xóa
script_rel = "skill/run_report.py"
p = os.path.join(root, "skill"); os.makedirs(p, exist_ok=True)
open(os.path.join(p, "run_report.py"), "w").write("print('v1')")
psid = s.add(MASTER_V2, title="tool report", content="script nhỏ",
             link="skill/run_report.py")
pres = s.promote(MASTER_V2, psid, "report-tool")
node_rows = s.load_rows("report-tool.xlsx")
check("promote: node mới có lá trỏ asset", node_rows and node_rows[0]["branch"].endswith("run_report.py"), node_rows[:1])
check("promote: asset dời vào folder node", os.path.exists(os.path.join(root, "report-tool", "run_report.py")))
check("promote: dòng cũ biến mất", not any(r["id"] == psid for r in s.load_rows(MASTER_V2)))
# id monotonic: thêm dòng sau khi xóa không được reuse id đã cấp
z_last2 = max(r["id"] for r in s.load_rows(MASTER_V2))
s.rm(MASTER_V2, z_last2)   # giải phóng 1 slot
tmp_id = s.add(MASTER_V2, title="tam", content="x")
s.rm(MASTER_V2, tmp_id)
new_id = s.add(MASTER_V2, title="moi", content="y")
check("ID không tái sử dụng sau xóa", new_id > tmp_id, f"{tmp_id} -> {new_id}")

# 11. tree + find + path_of
tr = s.tree()
check("tree có children", len(tr["children"]) >= 1)
f = s.find(mid)
check("find(id) xác định branch", f and f["branch"] == "pref.xlsx")
pth = s.path_of(mid)
check("path_of có chuỗi cây", pth == "brain.xlsx › pref.xlsx › #" + str(mid), pth)

# 12. search
h = s.search("nén 3 còn 1")
check("search toàn cây", h and h[0]["id"] == mid)

print(f"\n{sum(1 for _, ok in T if ok)}/{len(T)} passed")
sys.exit(0 if all(ok for _, ok in T) else 1)
