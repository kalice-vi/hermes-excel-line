"""brain_server.py — live mind-map server, two-way synced with the excel_line v2 tree.

Endpoints:
  GET  /api/data     : jsMind tree rendered straight from Excel (no cache)
  GET  /api/search   : ?q=... keyword search across the whole tree
  GET  /api/mmd      : mermaid .mmd source + revision hash (client change detect)
  POST /api/mmd_sync : {mmd: edited text} -> diff vs Excel -> apply ops -> canonical mmd
  POST /api/ops      : [ {op...}, ... ] applied directly through BrainStore (locked)
  static             : serves web/ assets (editor html + mermaid lib)

Excel is the single source of truth: every write goes through BrainStore —
the server never touches .xlsx files by a side path. User-created map
memories carry the tag 'map-added' so the agent curator leaves them alone.
"""
from __future__ import annotations
import json, os, sys, threading, urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))            # .../scripts
PLUGIN = os.path.dirname(HERE)                               # .../excel_line (plugin root)
STATIC = os.path.join(PLUGIN, "web")
sys.path.insert(0, PLUGIN)

from brain_store import BrainStore, FullError, BadBranch, MASTER_V2  # noqa: E402
from brain_map import build_tree                                  # noqa: E402
from mermaid_map import build_mmd, diff_ops                       # noqa: E402

ROOT = os.environ.get("EXCEL_LINE_ROOT",
                      os.path.join(os.path.expanduser("~"),
                                   "AppData", "Local", "hermes", "excel_line"))
store = BrainStore(ROOT)
_lock = threading.Lock()

# ------------------------------------------------------------------ ops

def op_add(o):
    branch = o.get("branch") or MASTER_V2
    topic = (o.get("topic") or "New").strip()[:50]
    tags = "map-added" + ("," + str(o.get("tag")).strip() if o.get("tag") else "")
    try:
        mid = store.add(branch, title=topic, content=topic, tags=tags)
    except FullError as e:
        return {"error": str(e) + " — merge rows on the map or drag them to another branch"}
    except BadBranch as e:
        return {"error": str(e)}
    return {"id": mid} if mid and mid > 0 else {"error": "store.add failed"}

def op_rename(o):
    rid = int(o["id"])
    f = store.find(rid)
    if not f:
        return {"error": "not found: #" + str(rid)}
    ok = store.set(f["branch"], rid, title=(o.get("topic") or "").strip()[:50])
    return {"ok": ok} if ok else {"error": "update failed"}

def op_retag(o):
    rid = int(o["id"])
    f = store.find(rid)
    if not f:
        return {"error": "not found: #" + str(rid)}
    old = str(f["row"].get("tags") or "")
    base = [t for t in old.split(",") if t.strip() and t.strip() != "map-added"]
    new_tag = (o.get("tag") or "").strip()
    tags = ",".join((["map-added"] if "map-added" in old else []) +
                    ([new_tag] if new_tag else base))
    ok = store.set(f["branch"], rid, tags=tags)
    return {"ok": ok} if ok else {"error": "update failed"}

def op_move(o):
    """Move a memory row between branch workbooks (drag on the map)."""
    rid = int(o["id"])
    src = o.get("src"); dst = o.get("dst") or MASTER_V2
    f = store.find(rid)
    if not f:
        return {"error": "not found: #" + str(rid)}
    src = src or f["branch"]
    if src == dst:
        return {"ok": True, "noop": True}
    r = f["row"]
    try:
        n = store.add(dst, title=r["title"], content=r["content"],
                      tags=r["tags"], link=r.get("branch") or "")
    except FullError as e:
        return {"error": str(e)}
    except BadBranch as e:
        return {"error": str(e)}
    store.rm(src, rid)
    return {"moved": True, "new_id": n}

def op_delete(o):
    rid = int(o["id"])
    f = store.find(rid)
    if not f:
        return {"error": "not found: #" + str(rid)}
    ok = store.rm(f["branch"], rid)
    return {"ok": ok} if ok else {"error": "delete failed"}

def op_child(o):
    """Create a new sub-branch workbook under a parent branch."""
    parent = o.get("parent") or MASTER_V2
    name = (o.get("name") or "branch").strip()
    try:
        res = store.child(parent, name,
                          title=(o.get("topic") or name)[:50])
    except (FullError, BadBranch) as e:
        return {"error": str(e)}
    return res

OPS = {"add": op_add, "rename": op_rename, "retag": op_retag, "move": op_move,
       "delete": op_delete, "child": op_child}

def apply_ops(ops):
    results = []
    for o in ops:
        fn = OPS.get(o.get("op"))
        if not fn:
            results.append({"op": o, "error": "unknown op: " + str(o.get("op"))})
            continue
        try:
            r = fn(o)
        except Exception as e:
            r = {"error": f"exception: {e}"}
        results.append({"op": o.get("op"), "id": o.get("id"), **r})
    return results

# ------------------------------------------------------------------ http

class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def _send(self, code, body, ctype="application/json"):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype + "; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        if u.path == "/api/data":
            try:
                tree = build_tree(ROOT)
                tree["meta"]["rev"] = store.count()
            except Exception as e:
                return self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
            return self._send(200, json.dumps(tree, ensure_ascii=False))
        if u.path == "/api/detail":
            rid_str = urllib.parse.parse_qs(u.query).get("id", [""])[0]
            if rid_str.isdigit():
                f = store.find(int(rid_str))
                if f:
                    return self._send(200, json.dumps({
                        "id": f["row"]["id"],
                        "title": f["row"]["title"],
                        "content": f["row"]["content"],
                        "tags": f["row"]["tags"],
                        "branch": f["branch"],
                        "updated": f["row"]["updated"],
                        "path": store.path_of(f["row"]["id"])
                    }, ensure_ascii=False))
            return self._send(404, json.dumps({"error": "not found"}, ensure_ascii=False))
        if u.path == "/api/search":
            q = urllib.parse.parse_qs(u.query).get("q", [""])[0]
            hits = store.search_index(q, limit=50) if q else []
            return self._send(200, json.dumps({"hits": hits}, ensure_ascii=False))
        if u.path == "/api/mmd":
            import hashlib
            try:
                mmd = build_mmd(ROOT)
            except Exception as e:
                return self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
            rev = hashlib.md5(mmd.encode("utf-8")).hexdigest()[:12]
            return self._send(200, json.dumps({"rev": rev, "mmd": mmd}, ensure_ascii=False))
        # static — served from the plugin web/ folder
        p = u.path.lstrip("/") or "brain_mermaid.html"
        fp = os.path.normpath(os.path.join(STATIC, p))
        if fp.startswith(STATIC) and os.path.isfile(fp):
            ctype = {".html": "text/html", ".js": "text/javascript",
                     ".css": "text/css", ".map": "application/json",
                     ".mmd": "text/plain"}.get(
                os.path.splitext(fp)[1], "application/octet-stream")
            with open(fp, "rb") as f:
                return self._send(200, f.read(), ctype)
        return self._send(404, b"not found", "text/plain")

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/api/ops":
            try:
                n = int(self.headers.get("Content-Length", 0))
                payload = json.loads(self.rfile.read(n).decode("utf-8"))
                ops = payload if isinstance(payload, list) else [payload]
                with _lock:
                    results = apply_ops(ops)
                return self._send(200, json.dumps({"results": results}, ensure_ascii=False))
            except Exception as e:
                return self._send(400, json.dumps({"error": str(e)}, ensure_ascii=False))
        if path == "/api/mmd_sync":
            try:
                n = int(self.headers.get("Content-Length", 0))
                payload = json.loads(self.rfile.read(n).decode("utf-8"))
                new_mmd = payload.get("mmd", "")
                with _lock:
                    old_mmd = build_mmd(ROOT)
                    ops = diff_ops(old_mmd, new_mmd)
                    results = apply_ops(ops) if ops else []
                    final_mmd = build_mmd(ROOT)
                return self._send(200, json.dumps(
                    {"applied": len(ops), "results": results, "mmd": final_mmd},
                    ensure_ascii=False))
            except Exception as e:
                return self._send(400, json.dumps({"error": str(e)}, ensure_ascii=False))
        return self._send(404, '{"error":"?"}')

if __name__ == "__main__":
    port = int(os.environ.get("BRAIN_PORT", "8766"))
    print(f"brain server → http://127.0.0.1:{port}/brain_mermaid.html (root={ROOT})", flush=True)
    ThreadingHTTPServer(("127.0.0.1", port), H).serve_forever()
