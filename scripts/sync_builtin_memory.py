"""sync_builtin_memory.py — mirror Hermes builtin MEMORY.md/USER.md into the v2 Excel tree.

LOCAL ONLY: uses BrainStore direct writes — no LLM, no network calls.
Strategy:
  1. Parse the live builtin entries (split on the section separator).
  2. Remove STALE v1 mirror rows (titles like "[MEMORY (default profile)] #N: ..."
     or "[USER (default profile)] #N: ...") — they are obsolete 50-char truncations.
     Other-profile rows ([...(gemini-agent profile)] etc.) are left untouched.
  3. For each current entry: if a live (non-stale) row already holds it
     (key = first 45 normalized chars), keep as-is; otherwise add it:
       MEMORY.md -> knowledge/misc.xlsx
       USER.md   -> user/<tạp-10 or best-fit>.xlsx
     Overflow auto-splits into a child 'mirror' branch.
  4. Re-run the coverage check. Print the resulting tree.
"""
import sys, os

sys.path.insert(0, r"C:\Users\Admin\AppData\Local\hermes\plugins\excel_line")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from brain_store import BrainStore, _flatten_rows, FullError, BadBranch

ROOT = r"C:\Users\Admin\AppData\Local\hermes\excel_line"
HOME_MEM = r"C:\Users\Admin\AppData\Local\hermes\memories"
STALE_PREFIXES = ("[memory (default profile)]", "[user (default profile)]")
TITLE_CAP, CONTENT_CAP = 50, 250

store = BrainStore(ROOT)

def norm(s):
    return " ".join(str(s or "").lower().split())

def entries(fn):
    txt = open(os.path.join(HOME_MEM, fn), encoding="utf-8").read()
    return [e.strip() for e in txt.split("\n\u00a7\n") if e.strip()]

mem_entries = entries("MEMORY.md")
usr_entries = entries("USER.md")
print(f"builtin: MEMORY.md={len(mem_entries)} USER.md={len(usr_entries)}")

# ---- 1. delete stale v1 mirror rows (this profile only)
rows = _flatten_rows(store)
stale = [r for r in rows if norm(r.get("title")).startswith(STALE_PREFIXES)]
print(f"stale v1 mirror rows to remove: {len(stale)}")
for r in stale:
    store.rm(r["_branch"], r["id"])

# ---- 2. add missing current entries
rows = _flatten_rows(store)  # refresh after deletion

def already_live(entry):
    key = norm(entry)[:45]
    for r in rows:
        blob = norm(r.get("title")) + " " + norm(r.get("content"))
        if key in blob:
            return r
    return None

def add_entry(entry, branch):
    title = entry[:TITLE_CAP]
    content = entry[:CONTENT_CAP]
    try:
        mid = store.add(branch, title=title, content=content, tags="memory-md,builtin-sync")
        return mid, branch
    except FullError:
        # split: child mirror branch under the owning zone file
        zone_file = os.path.dirname(branch) + "/" + "mirror.xlsx" if "/" in branch else "mirror.xlsx"
        parent = os.path.basename(branch)
        res = store.child(branch, "mirror", title="Builtin memory mirror overflow",
                          content="Auto-split mirror branch", tags="mirror")
        mid = store.add(res["file"], title=title, content=content, tags="memory-md,builtin-sync")
        return mid, res["file"]

added = skipped = 0
for e in mem_entries:
    hit = already_live(e)
    if hit:
        skipped += 1
        continue
    mid, br = add_entry(e, "knowledge/misc.xlsx")
    rows.append({"_branch": br, "title": e[:TITLE_CAP], "content": e[:CONTENT_CAP]})
    added += 1
    print(f"  +M #{mid} -> {br}: {e[:52]}")

for e in usr_entries:
    hit = already_live(e)
    if hit:
        skipped += 1
        continue
    mid, br = add_entry(e, "user/t\u1ea1p-10.xlsx")
    rows.append({"_branch": br, "title": e[:TITLE_CAP], "content": e[:CONTENT_CAP]})
    added += 1
    print(f"  +U #{mid} -> {br}: {e[:52]}")

print(f"added={added} skipped(already live)={skipped}")

# ---- 3. coverage re-check
rows = _flatten_rows(store)
blob_all = " || ".join(norm(r.get("title")) + " " + norm(r.get("content")) for r in rows)
def coverage(label, es):
    ok = 0
    for e in es:
        if norm(e)[:45] in blob_all:
            ok += 1
        else:
            print("  STILL MISSING:", e[:70])
    print(f"coverage {label}: {ok}/{len(es)}")
coverage("MEMORY.md", mem_entries)
coverage("USER.md", usr_entries)
print("total rows in tree:", store.count())
print()
print(store.tree_text())
