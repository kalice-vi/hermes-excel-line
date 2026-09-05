"""__init__.py — Hermes memory provider entry point for excel_line.

INTRODUCTION
    This is the plugin's public surface: it registers the excel_line memory
    provider with Hermes and wires together the three moving parts — the tool
    handlers (add / search / read), the auto-retrieve hook (prefetch), and the
    background indexer (worker.py). It decides WHEN to persist and HOW memory is
    surfaced back into every session.

DESIGN (agent-driven, not turn-driven)
    - The agent itself decides WHEN to persist. After producing an output it deems
      worth keeping, it calls the excel_line `add` tool with the session id and the
      sequence numbers of the input/output to log (0 = omit that side).
    - Direct-store mode: when the agent passes explicit zone + brief/content, the
      record is written straight to the store with NO LLM dependency (the durable
      path that behaves like built-in memory).
    - The sub-agent (worker.py) drains the log folder continuously for lazy logs.

RELATIONSHIP TO BUILT-IN MEMORY
    Runs ALONGSIDE built-in MEMORY.md/USER.md (built-in = fast cache, excel_line =
    durable human-auditable store). Built-in MEMORY.md writes are mirrored via
    on_memory_write.

CONFIG (plugins.excel_line in config.yaml)
    root:       directory for the workbooks (default $HERMES_HOME/excel_line)
    log_dir:    shared folder where log files for the indexer live (default root/logs)
    free_model: model id for the classifier (default gemini-3.5-flash-lite)
"""

# excel_line long-term memory provider.

from __future__ import annotations

import json
import logging
import os
import re
import time
from typing import Any, Dict, List, Optional

from agent.memory_provider import MemoryProvider
from tools.registry import tool_error

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Tool schemas
# ---------------------------------------------------------------------------

EXCEL_LINE_SCHEMA = {
    "name": "excel_line",
    "description": (
        "Excel-backed hierarchical long-term memory — a TREE of workbooks where "
        "every .xlsx holds AT MOST 10 rows (ID|Title|Content|Tags|Branch|Updated). "
        "brain.xlsx is the root layer-1. A row with Branch='<x>.xlsx' is a NODE "
        "(x.xlsx lives in the same folder as its parent file's stem); a row with a "
        "non-.xlsx Branch is a LEAF asset file and must sit at the deepest level; "
        "empty Branch = plain memory. Full file => merge older rows (compress) or "
        "child() to split — never overflow. When a leaf asset grows into a project, "
        "promote() replaces it with a .xlsx node and pushes the asset behind it.\n\n"
        "ACTIONS:\n"
        "• add — store one memory: branch (file path rel root, default brain.xlsx) "
        "+ title<=50 + content<=250 + tags. Full branch => error telling you to merge/child.\n"
        "• search — keyword scan the WHOLE tree; returns hit with branch path.\n"
        "• read — rows of a branch file (branch='user.xlsx' or legacy zone name).\n"
        "• tree — ASCII rendering of the tree with per-file usage (n/10).\n"
        "• child — create a new sub-branch file under parent + the linking row.\n"
        "• move — move row ids from one file to another (rebalance).\n"
        "• merge — compress several rows (ids) into ONE new memory (frees slots).\n"
        "• promote — turn a leaf-asset row into a .xlsx node, asset moves behind it.\n"
        "• update/delete — edit/remove a row by row_id (branch optional, auto-found).\n"
        "• forget — delete all rows matching a keyword.\n"
        "• add (legacy mode) — session+input_seq/output_seq logs a turn for the "
        "free-model sub-agent curator to classify/compress into the tree.\n"
        "Prefer add with explicit branch+title+content (direct, durable). Keep "
        "memories compressed, reusable, never verbatim conversation.\\n"
        "Use search/tree to RECALL — excel_line is the long-term brain; built-in "
        "memory is only the fast cache."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "action": {
                "type": "string",
                "enum": ["add", "search", "read", "tree", "child", "move", "merge",
                         "promote", "update", "delete", "forget", "zones"],
            },
            "session": {"type": "string",
                        "description": "Legacy add: session id whose turn to log ('current')."},
            "input_seq": {"type": "integer",
                          "description": "Legacy add: 1-based input seq to log (0 omit)."},
            "output_seq": {"type": "integer",
                           "description": "Legacy add: 1-based output seq to log (0 omit)."},
            "branch": {"type": "string",
                       "description": "Target .xlsx branch rel root (e.g. 'brain.xlsx', 'skill/einvoicing.xlsx'); '' = brain.xlsx."},
            "parent": {"type": "string", "description": "child: parent .xlsx."},
            "name": {"type": "string", "description": "child/promote: new node name (no .xlsx)."},
            "ids": {"type": "array", "items": {"type": "integer"},
                    "description": "move/merge: row ids."},
            "dst": {"type": "string", "description": "move: destination .xlsx (must have room)."},
            "title": {"type": "string", "description": "<=50 chars."},
            "content": {"type": "string", "description": "<=250 chars, compressed."},
            "tags": {"type": "string", "description": "comma keywords for branch navigation."},
            "link": {"type": "string",
                     "description": "add: LEAF asset path rel root that must exist (e.g. 'skill/x.py'); omit for plain memory."},
            "row_id": {"type": "integer", "description": "update/delete/promote: row id."},
            "query": {"type": "string", "description": "search/forget keyword."},
            "zone": {"type": "string", "description": "Legacy alias of branch for read/add."},
            "brief": {"type": "string", "description": "Legacy alias: title/content for direct store."},
            "limit": {"type": "integer", "description": "Max rows (default 10)."},
        },
        "required": ["action"],
    },
}


# ---------------------------------------------------------------------------
# Config loader
# ---------------------------------------------------------------------------

def _load_plugin_config() -> dict:
    try:
        from hermes_cli.config import load_config_readonly, cfg_get
        all_config = load_config_readonly()
        return cfg_get(all_config, "plugins", "excel_line", default={}) or {}
    except Exception:
        return {}


def _root_dir(cfg: dict) -> str:
    from hermes_constants import get_hermes_home
    default = str(get_hermes_home()) + "/excel_line"
    root = cfg.get("root", default)
    root = root.replace("$HERMES_HOME", str(get_hermes_home()))
    return root


def _log_dir(cfg: dict, root: str) -> str:
    from hermes_constants import get_hermes_home
    return cfg.get("log_dir", root + "/logs").replace(
        "$HERMES_HOME", str(get_hermes_home())
    )


def _safe_sid(sid: str) -> str:
    """Sanitize a session id for use in a filename. Strips path separators on
    both POSIX and Windows (W7, r8: a sid like '..\\foo' must not escape log_dir)."""
    return sid.replace("/", "_").replace("\\", "_").replace("..", "_").strip() or "default"


# ---------------------------------------------------------------------------
# Provider
# ---------------------------------------------------------------------------

class ExcelLineProvider(MemoryProvider):
    """Excel-backed long-term memory; agent-driven logging + free-model indexer."""

    def __init__(self, config: dict | None = None, llm=None):
        self._config = config or _load_plugin_config()
        self._llm = llm  # host PluginLlm facade (ctx.llm) for classification
        self._store: Optional["ExcelLineStore"] = None
        self._root = ""
        self._log_dir = ""
        self._session_id = ""
        # Lightweight per-session transcript buffer (lookup only, never indexed).
        self._transcripts: Dict[str, List[Dict[str, str]]] = {}

    # -- required ABC ------------------------------------------------------

    @property
    def name(self) -> str:
        return "excel_line"

    def is_available(self) -> bool:
        try:
            import openpyxl  # noqa: F401
            return True
        except Exception:
            return False

    def initialize(self, session_id: str, **kwargs) -> None:
        self._root = _root_dir(self._config)
        self._log_dir = _log_dir(self._config, self._root)
        os.makedirs(self._root, exist_ok=True)
        os.makedirs(self._log_dir, exist_ok=True)
        from .brain_store import BrainStore
        self._store = BrainStore(self._root)
        self._session_id = session_id
        # Auto-restart brain server after gateway restart (port 8766)
        try:
            import socket, subprocess
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM); s.settimeout(0.3)
            try:
                s.connect(("127.0.0.1", 8766))
                s.close()  # already running
            except Exception:
                s.close()
                script_dir = os.path.join(os.path.dirname(__file__))
                server_py = os.path.join(script_dir, "scripts", "brain_server.py")
                if os.path.exists(server_py):
                    subprocess.Popen(["python", server_py], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, close_fds=True, creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
        except Exception:
            pass

    def get_tool_schemas(self) -> List[Dict[str, Any]]:
        return [EXCEL_LINE_SCHEMA]

    # -- prompt + recall ---------------------------------------------------

    def system_prompt_block(self) -> str:
        if not self._store:
            return ""
        n = self._store.count()
        zones = ", ".join(self._store.list_zones()) or "(none yet)"
        return (
            "# Excel-Line Memory\n"
            f"Active. {n} indexed memories across zones: {zones}.\n"
            "Durable, human-auditable store in Excel. Use excel_line(search/read) to "
            "recall facts, preferences, projects, contacts, and concise knowledge. "
            "To SAVE a turn, call excel_line(add, session, input_seq, output_seq) — the "
            "sub-agent will index and clean it up. Built-in memory is the fast cache; "
            "excel_line is the long-term store."
        )

    def _drain_pending_logs(self) -> None:
        """Lazily index any seq-logs left behind (e.g. indexer crashed or never
        ran) so memory is never stranded on disk. Safe to call often; it only
        acts when logs exist."""
        try:
            import os
            log_dir = getattr(self, "_log_dir", "")
            if not log_dir or not os.path.isdir(log_dir):
                return
            pending = [n for n in os.listdir(log_dir)
                       if (n.endswith(".jsonl") or n.endswith(".json"))
                       and not n.startswith("session_raw_")]
            if not pending:
                return
            try:
                from .worker import index_while_logs_present
                free_model = self._config.get("free_model", "gemini-3.5-flash-lite")
                index_while_logs_present(
                    log_dir=log_dir, store_root=self._root,
                    free_model_fn=lambda p: _ask_free_model(p, free_model, llm=self._llm),
                    store=self._store)
            except Exception as e:
                logger.debug("excel_line lazy index failed: %s", e)
        except Exception:
            pass

    def prefetch(self, query: str, *, session_id: str = "") -> str:
        if not self._store or not query:
            return ""
        try:
            q_lower = query.lower()
            # Auto-extract trigger keywords (no manual list)
            def _extract_triggers(q_text: str) -> list:
                tokens = re.findall(r"[a-záàảãạăắằẳẵặâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵđ]+", q_text.lower())
                stop = {"tôi", "là", "có", "từ", "và", "hay", "cũng", "còn", "để", "trong", "với", "về", "của", "cho", "đến", "bằng", "qua", "khi", "nếu", "thì", "mà", "nhưng", "hoặc", "vì", "vậy", "do", "đó", "như", "nên", "lại", "theo"}
                return [w for w in tokens if len(w) > 1 and w not in stop]
            triggers = _extract_triggers(query)
            want_recall = bool(triggers)
            if not want_recall:
                logger.debug("prefetch: no meaningful keywords in query; skip: %s", query)
                return ""

            # --- Tree-traversal: men từ brain.xlsx → branch → leaf ---
            def score_row(row, q):
                """Weighted keyword overlap score via substring match (fuzzy)."""
                q_words = triggers
                title = (row.get("title") or "").lower()
                tags  = (row.get("tags")  or "").lower()
                content = (row.get("content") or "").lower()
                score = 0
                for w in q_words:
                    if len(w) < 2:
                        continue
                    # Substring match (not word-boundary) so partial words score too
                    if w in title:
                        score += 3
                    if w in tags:
                        score += 2
                    if w in content:
                        score += 1
                return score

            def walk(branch: str, depth: int) -> list:
                """Traverse one best child at each level, including descendant scores."""
                try:
                    rows = self._store.load_rows(branch)
                except Exception:
                    return []
                candidates = []
                for row in rows:
                    own_score = score_row(row, query)
                    child_branch = str(row.get("branch") or "")
                    deeper = walk(child_branch, depth + 1) if child_branch.lower().endswith(".xlsx") else []
                    child_score = deeper[0][0] if deeper else 0
                    # A parent branch inherits its strongest descendant's score.
                    total_score = own_score + child_score
                    if total_score > 0:
                        candidates.append((total_score, row, branch, deeper))
                if not candidates:
                    return []
                # Pick exactly one of the <=10 rows at this level, then follow it.
                candidates.sort(key=lambda item: item[0], reverse=True)
                total_score, best_row, current_branch, deeper = candidates[0]
                return [(total_score, best_row, current_branch)] + deeper[:9]

            # Bắt đầu từ brain.xlsx (layer 1 root)
            try:
                from .brain_store import MASTER_V2
            except ImportError:
                from brain_store import MASTER_V2
            tree_hits = walk(MASTER_V2, 0)
            if not tree_hits:
                return ""

            lines = ["## Excel-Line Memory (tree path)"]
            seen_ids = set()
            for entry in tree_hits:
                score, row, branch = entry[0], entry[1], entry[2] if len(entry) > 2 else "brain.xlsx"
                rid = row["id"]
                if rid in seen_ids:
                    continue
                seen_ids.add(rid)
                branch_short = os.path.splitext(os.path.basename(branch))[0]
                title = row.get("title") or ""
                content = row.get("content") or ""
                tags = row.get("tags") or ""
                if content:
                    lines.append(f"  • [{branch_short} #{rid}] {title}: {content}")
                else:
                    lines.append(f"  • [{branch_short} #{rid}] {title}  (tags: {tags})")
            return "\n".join(lines)
        except Exception as e:
            logger.debug("excel_line prefetch failed: %s", e)
            return ""

    def sync_turn(self, user_content: str, assistant_content: str, *,
                  session_id: str = "", messages=None) -> None:
        """Attach this turn's OUTPUT to the transcript entry that on_turn_start
        already opened for the current turn. The input was recorded up front so
        the excel_line `add` tool can reference the CURRENT turn (not just past
        ones). This buffer is NEVER auto-indexed — only explicit `add` calls
        create logs."""
        if not user_content and not assistant_content:
            return
        sid = session_id or self._session_id or "default"
        buf = self._transcripts.setdefault(sid, [])
        if buf:
            # Update the most recent (current-turn) entry's output in place.
            buf[-1]["output"] = assistant_content or buf[-1].get("output", "")
        else:
            buf.append({
                "input": user_content or "",
                "output": assistant_content or "",
                "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
            })

    def on_turn_start(self, turn_number: int, message: str, **kwargs) -> None:
        """Record the user's input at the START of the turn so the excel_line
        `add` tool can reference the current turn (input is available before the
        agent even replies). Output is filled in later by sync_turn."""
        if not message:
            return
        sid = kwargs.get("session_id") or self._session_id or "default"
        buf = self._transcripts.setdefault(sid, [])
        # If the last entry has no input yet (fresh turn), reuse it; else append.
        if buf and buf[-1].get("input") == "" and buf[-1].get("output") == "":
            buf[-1]["input"] = message
            buf[-1]["ts"] = time.strftime("%Y-%m-%d %H:%M:%S")
        else:
            buf.append({
                "input": message,
                "output": "",
                "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
            })

    def handle_tool_call(self, tool_name: str, args: Dict[str, Any], **kwargs) -> str:
        if tool_name != "excel_line":
            return tool_error(f"Unknown tool: {tool_name}")
        return self._handle(args)

    def shutdown(self) -> None:
        self._store = None
        self._transcripts.clear()

    # -- optional hooks ----------------------------------------------------

    def on_memory_write(self, action: str, target: str, content: str,
                        metadata=None) -> None:
        """Mirror built-in memory writes into the excel_line store."""
        if action == "add" and self._store and content:
            try:
                zone = "user" if target == "user" else "knowledge"
                self._store.add(
                    zone=zone,
                    brief=content[:120],
                    content=content[:300],
                    title="mirrored:" + target,
                    tags="memory-md",
                )
            except Exception as e:
                logger.debug("excel_line mirror failed: %s", e)

    # -- indexer (sub-agent) ----------------------------------------------

    def _run_indexer(self) -> int:
        """Drain the log folder completely: keep indexing while logs remain."""
        from .worker import index_while_logs_present
        free_model = self._config.get("free_model", "gemini-3.5-flash-lite")

        def _call_free(prompt: str) -> str:
            return _ask_free_model(prompt, free_model, llm=self._llm)

        return index_while_logs_present(
            log_dir=self._log_dir,
            store_root=self._root,
            free_model_fn=_call_free,
            store=self._store,
        )

    # -- auto-extract (backup pass on session end) ------------------------

    def on_session_end(self, messages: List[Dict[str, Any]]) -> None:
        """Backup pass: the free-model sub-agent reads the session transcript,
        decides which facts are worth keeping, and writes them to Excel — so
        memory persists even when the agent did not call excel_line add.
        If extraction fails entirely, the raw transcript is saved to a log so a
        later run (or the lazy-indexer) can still recover it."""
        if not self._store:
            return
        try:
            stored = self._auto_extract(messages)
            if stored == 0:
                # LLM produced nothing usable — keep the raw transcript so it
                # is not silently lost (mirrors worker's no-drop guarantee).
                self._save_raw_transcript(messages)
        except Exception as e:
            logger.debug("excel_line auto_extract failed: %s", e)
            self._save_raw_transcript(messages)
        # Free per-session transcript buffer so a long-running process with many
        # sessions does not accumulate unbounded RAM (Gemini round-3 WARN).
        self._transcripts.pop(self._session_id or "default", None)

    def _save_raw_transcript(self, messages) -> None:
        """Persist the raw transcript to the log dir so it can be indexed later
        (lazy indexer / background job) instead of being dropped."""
        try:
            import os, time as _t, json as _json
            log_dir = getattr(self, "_log_dir", "")
            if not log_dir or not os.path.isdir(log_dir):
                return
            sid = self._session_id or "default"
            turns = self._transcript_from_messages(messages)
            if not turns:
                turns = [f"User: {t['input']}\nAssistant: {t['output']}"
                         for t in self._transcripts.get(sid, []) if t.get("input") or t.get("output")]
            if not turns:
                return
            stamp = _t.strftime("%Y%m%d-%H%M%S") + f"{int(_t.time()*1000)%1000:03d}"
            path = os.path.join(log_dir, f"session_raw_{_safe_sid(sid)}_{stamp}.jsonl")
            with open(path, "w", encoding="utf-8") as f:
                for line in turns:
                    # Schema matches the standard log format worker.py consumes
                    # (session, input, output, ts) so the lazy indexer can read it.
                    if line.startswith("User:") or line.startswith("Assistant:"):
                        role, _, txt = line.partition(":")
                        f.write(_json.dumps({"session": sid, "input": txt.strip() if role == "User" else "",
                                             "output": txt.strip() if role == "Assistant" else "",
                                             "ts": stamp}) + "\n")
                    else:
                        f.write(_json.dumps({"session": sid, "input": line, "output": "", "ts": stamp}) + "\n")
        except Exception as e:
            logger.debug("excel_line raw transcript backup failed: %s", e)

    def _auto_extract(self, messages) -> int:
        # Build the transcript text: prefer the raw messages Hermes passes,
        # fall back to our own per-session buffer.
        turns = self._transcript_from_messages(messages)
        if not turns:
            sid = self._session_id or "default"
            turns = [f"User: {t['input']}\nAssistant: {t['output']}"
                     for t in self._transcripts.get(sid, []) if t.get("input") or t.get("output")]
        if not turns:
            return 0

        transcript = "\n\n".join(turns)
        prompt = (
            "You are a memory curator. From the conversation transcript below, "
            "extract ONLY durable, reusable facts worth remembering long-term "
            "(user preferences, assets, decisions, contacts, project context, "
            "how-tos). Ignore chitchat, greetings, and ephemeral task output.\n"
            "Reply with strict JSON array (max 8 items), each: "
            '{"zone":"user|project|pref|task|knowledge|contact",'
            '"brief":"1-sentence summary (max 120 chars)",'
            '"content":"concise knowledge (max 300 chars)",'
            '"tags":"comma keywords"}\n\n'
            f"TRANSCRIPT:\n{transcript[:6000]}"
        )
        raw = _ask_free_model(prompt, self._config.get("free_model", "gemini-3.5-flash-lite"),
                              llm=self._llm)
        facts = self._parse_facts(raw)
        if not facts:
            # LLM down or returned garbage: do NOT dump raw turns into the
            # store (2026-09 audit: _fallback_store polluted 5.8k rows).
            # on_session_end keeps the raw transcript as a session_raw_ log
            # for a future, properly-classified pass instead.
            return 0
        stored = 0
        for f in facts:
            zone = f.get("zone", "knowledge")
            brief = (f.get("brief") or "")[:120]
            content = (f.get("content") or "")[:300]
            tags = f.get("tags", "")
            if not (brief or content):
                continue
            # The sub-agent already curated these facts; write them straight to
            # the Excel store (this IS the sub-agent's write pass). No log file
            # needed, so nothing to delete later.
            try:
                rid = self._store.add(zone=zone, brief=brief, content=content,
                                     title=brief[:40], tags=tags or "auto-extract")
                if rid and rid > 0:
                    stored += 1
            except Exception as e:
                logger.debug("excel_line auto_extract store failed: %s", e)
        return stored

    @staticmethod
    def _transcript_from_messages(messages) -> List[str]:
        out = []
        if not isinstance(messages, list):
            return out
        for m in messages:
            role = m.get("role")
            content = m.get("content")
            if isinstance(content, str) and content.strip():
                out.append(f"{role.capitalize()}: {content.strip()}")
        return out

    @staticmethod
    def _parse_facts(raw: str) -> List[Dict]:
        if not raw:
            return []
        raw = raw.strip()
        if raw.startswith("```"):
            raw = raw.strip("`")
            s, e = raw.find("["), raw.rfind("]")
            if s != -1 and e != -1:
                raw = raw[s:e + 1]
        try:
            data = json.loads(raw)
            return data if isinstance(data, list) else []
        except Exception:
            # Last-resort: extract the first [...] JSON array even if the model
            # wrapped it in prose or used wrong fences (Gemini review #8).
            import re
            m = re.search(r"\[.*\]", raw, re.DOTALL)
            if m:
                try:
                    data = json.loads(m.group(0))
                    return data if isinstance(data, list) else []
                except Exception:
                    pass
            return []

    def _schedule_index(self) -> None:
        """Run the drain-indexer in a daemon thread (non-blocking)."""
        import threading
        def _job():
            try:
                self._run_indexer()
            except Exception as e:
                logger.debug("excel_line background index failed: %s", e)
        threading.Thread(target=_job, daemon=True).start()

    # -- tool dispatch -----------------------------------------------------

    def _handle(self, args: dict) -> str:
        if not self._store:
            return tool_error("excel_line not initialized")
        action = args.get("action")
        try:
            if action == "add":
                return self._handle_add(args)
            if action == "tree":
                return json.dumps({"tree": self._store.tree_text()})
            if action == "child":
                try:
                    res = self._store.child(args.get("parent", "brain.xlsx"),
                                            args.get("name", "sub"),
                                            title=args.get("title", ""),
                                            content=(args.get("content") or "")[:250],
                                            tags=args.get("tags", ""))
                    return json.dumps({"status": "created", **res})
                except Exception as e:
                    return tool_error(str(e))
            if action == "move":
                try:
                    n = self._store.move(args.get("branch", "brain.xlsx"),
                                         [int(i) for i in args.get("ids", [])],
                                         args.get("dst", "brain.xlsx"))
                    return json.dumps({"status": "moved", "count": n})
                except Exception as e:
                    return tool_error(str(e))
            if action == "merge":
                try:
                    mid = self._store.merge(args.get("branch", "brain.xlsx"),
                                            [int(i) for i in args.get("ids", [])],
                                            args.get("title", "merged"),
                                            (args.get("content") or "")[:250],
                                            args.get("tags", ""))
                    return json.dumps({"status": "merged", "id": mid})
                except Exception as e:
                    return tool_error(str(e))
            if action == "promote":
                try:
                    rid = int(args.get("row_id", 0))
                    br = args.get("branch")
                    if not br:
                        f = self._store.find(rid)
                        br = f["branch"] if f else "brain.xlsx"
                    res = self._store.promote(br, rid, args.get("name", "tool"))
                    return json.dumps({"status": "promoted", **res})
                except Exception as e:
                    return tool_error(str(e))
            if action == "search":
                hits = (self._store.search(args.get("query", ""),
                                           limit=int(args.get("limit", 10)))
                        if hasattr(self._store, "search")
                        else self._store.search_index(args.get("query", ""),
                                                      limit=int(args.get("limit", 10))))
                return json.dumps({"results": hits, "count": len(hits)})
            if action == "read":
                br = args.get("branch") or args.get("zone") or "brain.xlsx"
                if not str(br).endswith(".xlsx"):
                    br = str(br) + ".xlsx"
                try:
                    rows = self._store.load_rows(br)
                except Exception:
                    rows = []
                return json.dumps({"branch": br, "rows": rows, "count": len(rows),
                                   "usage": f"{len(rows)}/10"})
            if action == "zones":
                return json.dumps({"zones": self._store.list_zones()})
            if action == "update":
                zone = (args.get("zone") or "knowledge").strip()
                rid = int(args.get("row_id", 0) or 0)
                if rid == 0:
                    return tool_error("update requires row_id")
                ok = self._store.update(
                    zone=zone, row_id=rid,
                    brief=(args.get("brief") or "").strip()[:120],
                    content=(args.get("content") or "").strip()[:300],
                    title=(args.get("title") or "").strip()[:40],
                    tags=(args.get("tags") or "").strip())
                return json.dumps({"status": "updated" if ok else "not_found", "id": rid, "zone": zone})
            if action == "delete":
                zone = (args.get("zone") or "knowledge").strip()
                rid = int(args.get("row_id", 0) or 0)
                if rid == 0:
                    return tool_error("delete requires row_id")
                ok = self._store.delete(zone=zone, row_id=rid)
                return json.dumps({"status": "deleted" if ok else "not_found", "id": rid, "zone": zone})
            if action == "forget":
                q = (args.get("query") or "").strip()
                if not q:
                    return tool_error("forget requires query")
                n = self._store.forget(q)
                return json.dumps({"status": "forgotten", "removed": n, "query": q})
            return tool_error(f"Unknown action: {action}")
        except Exception as exc:
            return tool_error(str(exc))

    def _handle_add(self, args: dict) -> str:
        """Save a turn. Three modes:
        - Direct-store mode: agent passes explicit `zone` + (`brief`/`content`/
          `title`) -> written STRAIGHT to the store, no LLM needed. This is the
          reliable path that behaves like built-in memory (always persists).
        - Sequence mode: session + input_seq/output_seq (>0) -> looked up from
          the transcript; classified by the free-model sub-agent via a log file.
        - Direct-log mode: input_text / output_text (no explicit zone) -> written
          to a log for the indexer to classify + store.
        Direct-store mode is preferred for durability; the other two depend on
        the indexer and are best-effort."""
        sid = args.get("session") or "current"
        if sid == "current":
            sid = self._session_id or "default"

        # --- Direct-store mode: write immediately, no LLM dependency ---
        # v2 tree params: branch/title/content/tags/link ; legacy: zone/brief
        zone = (args.get("zone") or "").strip()
        branch = (args.get("branch") or "").strip()
        brief = (args.get("brief") or "").strip()
        content = (args.get("content") or "").strip()
        title = (args.get("title") or brief[:50]).strip()
        tags = (args.get("tags") or "").strip()
        link = (args.get("link") or "").strip()
        if (branch or zone) and (title or content or brief):
            from .brain_store import FullError
            try:
                if branch or link:
                    mid = self._store.add(branch or zone or "brain.xlsx",
                                          title=title[:50], content=content[:250],
                                          tags=tags, link=link)
                else:
                    mid = self._store.add(zone=zone, brief=brief[:250],
                                          content=content[:250], title=title[:50],
                                          tags=tags)
                if not mid or mid < 0:
                    return json.dumps({
                        "status": "error", "id": mid, "branch": branch or zone,
                        "note": "Store write failed; memory was NOT persisted.",
                    })
                return json.dumps({
                    "status": "stored", "id": mid,
                    "branch": branch or zone,
                    "path": self._store.path_of(mid) if hasattr(self._store, "path_of") else "",
                    "note": "Written to the Excel tree (no indexer needed).",
                })
            except Exception as e:
                if type(e).__name__ == "FullError":
                    return json.dumps({
                        "status": "branch_full", "error": str(e),
                        "note": "File is full (10/10). Please merge/compress older rows (select similar IDs) OR create a child() branch and retry.",
                    })
                return tool_error(f"Direct store failed: {e}")

        # --- Sequence / direct-log modes: go through the log + indexer ---
        direct_in = (args.get("input_text") or "").strip()
        direct_out = (args.get("output_text") or "").strip()
        if direct_in or direct_out:
            in_text, out_text = direct_in, direct_out
            in_seq = out_seq = 0
        else:
            buf = self._transcripts.get(sid)
            if not buf:
                return tool_error(f"No transcript for session '{sid}'. Pass zone+brief+content "
                                  f"for direct store, or input_text/output_text.")
            in_seq = int(args.get("input_seq", 0) or 0)
            out_seq = int(args.get("output_seq", 0) or 0)
            if in_seq == 0 and out_seq == 0:
                return tool_error("Provide zone+brief+content, input_seq/output_seq (>0), "
                                  "or input_text/output_text.")
            in_text = buf[in_seq - 1]["input"] if 0 < in_seq <= len(buf) else ""
            out_text = buf[out_seq - 1]["output"] if 0 < out_seq <= len(buf) else ""
            if not in_text and not out_text:
                return tool_error("Requested sequence numbers are out of range for this session.")

        # Write a log file for the indexer (filename encodes the source so it is
        # unique and never collides with another pending log).
        stamp = time.strftime("%Y%m%d-%H%M%S") + f"{int(time.time()*1000)%1000:03d}"
        log_name = f"{_safe_sid(sid)}_i{in_seq}_o{out_seq}_{stamp}.jsonl"
        log_path = os.path.join(self._log_dir, log_name)
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(json.dumps({
                "session": sid,
                "input_seq": in_seq,
                "output_seq": out_seq,
                "input": in_text,
                "output": out_text,
                "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
            }, ensure_ascii=False) + "\n")

        # Trigger the continuous indexer (background). It will process this log
        # and delete it; it keeps running while any log remains.
        self._schedule_index()
        return json.dumps({
            "status": "logged",
            "session": sid,
            "input_seq": in_seq,
            "output_seq": out_seq,
            "log": log_name,
            "note": "Sub-agent will index and delete this log automatically.",
        })


# ---------------------------------------------------------------------------
# Keyless free-model rotation (OpenCode Zen free tier) + host default
# ---------------------------------------------------------------------------
# Verified keyless slugs on the opencode-free provider (no API key ever).
# Live probe 03/09/2026: laguna OK 35s; muse-spark OK-ish (transient 5xx);
# hy3-free/x-preview-f-free currently 401; nemotrons time out host-side.
# Dead slugs stay LAST — they may recover, and rotation reaches them only
# when everything above fails.
FREE_ROTATION = [
    "laguna-s-2.1-free",
    "muse-spark-1.2-contributor-free",
    "nemotron-3.5-lightning-free",
    "nemotron-3-ultra-free",
    "hy3-free",
    "x-preview-f-free",
]
_choice_path = None          # module-level, set by _init_choice_path()
_preferred = None            # pinned model name (from /excel-line model), or None


def _load_pref() -> dict:
    """Read the user's model choice from <root>/model_choice.json."""
    global _choice_path, _preferred
    try:
        root = _root_dir(_load_plugin_config())
        _choice_path = os.path.join(root, "model_choice.json")
        if os.path.exists(_choice_path):
            with open(_choice_path, encoding="utf-8") as f:
                _preferred = json.load(f).get("preferred")
    except Exception:
        _preferred = None
    return {"preferred": _preferred}


def _save_pref(preferred):
    global _preferred
    _preferred = preferred
    try:
        if _choice_path:
            with open(_choice_path, "w", encoding="utf-8") as f:
                json.dump({"preferred": preferred}, f, ensure_ascii=False)
    except Exception:
        pass


def _call_provider(provider, model, prompt, timeout=20):
    """One keyless/known-provider completion attempt via the host client.
    Returns text or raises."""
    from agent.auxiliary_client import call_llm
    resp = call_llm(task=None, provider=provider, model=model,
                    messages=[{"role": "user", "content": prompt}],
                    timeout=timeout)
    try:
        return resp.choices[0].message.content or ""
    except Exception:
        return str(getattr(resp, "content", resp) or "")


def _rotation_chain(free_model):
    """Ordered (provider, model) attempts honoring the user's preference.

    ('host', None) sentinel resolved by caller → host default via llm facade.
    """
    pref = _preferred
    if pref == "host":
        return []                      # llm facade (host default) handles it
    chain = []
    if pref:
        if pref in FREE_ROTATION:
            chain.append(("opencode-free", pref))
        elif "/" in pref:
            p, m = pref.split("/", 1)
            chain.append((p, m))
        else:
            chain.append((None, pref))
    else:
        # no pin: configured free_model first (provider-agnostic resolution)
        if free_model:
            chain.append((None, free_model))
    # keyless rotation after the pin/config model
    for slug in FREE_ROTATION:
        if ("opencode-free", slug) not in chain:
            chain.append(("opencode-free", slug))
    return chain


def _ask_free_model(prompt: str, model: str, llm=None) -> str:
    """Classify a record via a ROTATING keyless free-model chain.

    Order: user-pinned model (from /excel-line model) -> configured free_model
    -> OpenCode-Zen free rotation (no API key needed at install time) ->
    host default model via the ctx.llm facade. Any failure (timeout, 401,
    UA-gate, trust-gate) moves to the next candidate; '' only when ALL fail
    (worker then keeps the log retryable and never writes to the store).
    """
    # 1-2) provider chain incl. keyless opencode-free rotation
    for provider, mdl in _rotation_chain(model):
        try:
            text = _call_provider(provider, mdl, prompt).strip()
            if text:
                return text
        except Exception:
            continue
    # 3) host default model through the plugin facade (last resort)
    if llm is not None:
        try:
            res = llm.complete(messages=[{"role": "user", "content": prompt}],
                               purpose="excel_line-classify")
            return (getattr(res, "text", "") or "").strip()
        except Exception:
            return ""
    # Legacy one-shot helper (existed in older runtimes; gone from core).
    try:
        from agent.run_agent import quick_completion  # type: ignore
        return (quick_completion(prompt, model=model) or "").strip()
    except Exception:
        return ""


def _model_command(raw_args: str) -> str:
    """/excel-line [model ...] — picker for providers/models integrated in Hermes.

    Usage:
      /excel-line model              list current choice + available models
      /excel-line model <n>          pick from the listed number
      /excel-line model host         use the host's default agent model
      /excel-line model auto         back to rotation (opencode-free keyless)
    """
    parts = (raw_args or "").strip().lower().split()
    _load_pref()
    arg = parts[0] if parts else ""
    if arg in ("model", "") or (arg == "model" and len(parts) == 1):
        arg = parts[1] if len(parts) > 1 and arg == "model" else ""
    if not arg:
        # ---- list ----
        lines = ["🧠 **excel_line classifier model**",
                 f"Currently selected: **{_preferred or 'auto (free rotation + host default)'}**", ""]
        lines.append("Keyless — OpenCode Zen free (no API key required):")
        n = 1
        idx = {}
        for slug in FREE_ROTATION:
            ok = " ⭐" if _preferred == slug else ""
            lines.append(f"  {n}. `{slug}`{ok}")
            idx[n] = slug; n += 1
        lines.append(f"  {n}. `host` — Hermes main agent model{'' if _preferred!='host' else ' ⭐'}")
        idx[n] = "host"; n += 1
        lines.append(f"  {n}. `auto` — rotate all keyless models{'' if _preferred else ' ⭐'}")
        idx[n] = "auto"
        # extra: a couple of known no-key providers already configured
        try:
            from hermes_cli.config import load_config_readonly, cfg_get
            cfg = load_config_readonly()
            prov = cfg.get("model", {}).get("provider")
            mdl = cfg.get("model", {}).get("model")
            if prov and mdl:
                lines.append(f"\nConfigured Hermes provider: `{prov}/{mdl}`")
                lines.append(f"Use: `/excel-line model {prov}/{mdl}`")
        except Exception:
            pass
        lines.append("\nSelect by number: `/excel-line model <number>`")
        _save_pref(_preferred)  # ensures file exists; keep value
        _model_command._idx = idx
        return "\n".join(lines)

    if arg == "auto":
        _save_pref(None)
        return "✅ Reset model preference: free model rotation (OpenCode Zen keyless) → host default."
    if arg == "host":
        _save_pref("host")
        return "✅ excel_line will classify memories using your main Hermes model."
    idxmap = getattr(_model_command, "_idx", {})
    if arg.isdigit() and int(arg) in idxmap:
        arg = idxmap[int(arg)]
    if arg in FREE_ROTATION:
        _save_pref(arg)
        return f"✅ Pinned `{arg}` (OpenCode Zen free, keyless) for excel_line memory."
    if "/" in arg:
        _save_pref(arg)
        return f"✅ Pinned `{arg}`."
    _save_pref(arg)
    return f"✅ Set model to `{arg}`."


# ---------------------------------------------------------------------------
# Plugin entry point
# ---------------------------------------------------------------------------

def register(ctx) -> None:
    config = _load_plugin_config()
    _load_pref()   # model_choice.json của user (từ /excel-line model)
    provider = ExcelLineProvider(config=config, llm=getattr(ctx, "llm", None))
    ctx.register_memory_provider(provider)
    try:
        ctx.register_command(
            "excel-line", _model_command,
            description="Chọn model phân loại memory excel_line (free/keyless rotation)",
            args_hint="[model] [tên|số|host|auto]")
    except Exception as e:
        logging.getLogger(__name__).debug("excel_line command register failed: %s", e)
