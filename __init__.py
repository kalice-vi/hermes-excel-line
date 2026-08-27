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

from __future__ import annotations

import json
import logging
import os
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
        "Excel-backed long-term memory. The agent keeps durable, human-auditable "
        "knowledge in Excel workbooks: a master index (brief + path) plus per-zone "
        "files (user, project, pref, task, knowledge...).\n\n"
        "ACTIONS:\n"
        "• add — Log a specific turn for the sub-agent to index. Provide session id, "
        "the input sequence number to log (0 to omit), and the output sequence number "
        "to log (0 to omit). The plugin finds that input/output in the session, writes "
        "a log, and the free-model sub-agent indexes then deletes it.\n"
        "• search — keyword search the master index (fast scan of brief + tags).\n"
        "• read — open a zone file and return its concise knowledge rows.\n"
        "• zones — list existing zone workbooks.\n"
        "• update — edit an existing memory row (needs row_id + zone; pass brief/content/title/tags to change).\n"
        "• delete — remove a memory row by row_id + zone.\n"
        "• forget — delete all memories matching a keyword (e.g. forget 'old project').\n"
        "Use excel_line to recall durable facts, preferences, projects, contacts, and "
        "concise knowledge the user expects you to remember across sessions."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "action": {
                "type": "string",
                "enum": ["add", "search", "read", "zones", "update", "delete", "forget"],
            },
            "session": {
                "type": "string",
                "description": "Session id whose turn to log (use 'current' for the active session).",
            },
            "input_seq": {
                "type": "integer",
                "description": "1-based sequence number of the input to log (0 to omit input).",
            },
            "output_seq": {
                "type": "integer",
                "description": "1-based sequence number of the output to log (0 to omit output).",
            },
            "query": {"type": "string", "description": "Keyword for 'search'."},
            "zone": {
                "type": "string",
                "description": "Zone name for 'read'/'add' (e.g. user, project, knowledge).",
            },
            "brief": {"type": "string", "description": "1-sentence summary for 'add' (goes to master index)."},
            "content": {"type": "string", "description": "Concise knowledge for 'add' (goes to the zone file)."},
            "title": {"type": "string", "description": "Short title for 'add'."},
            "tags": {"type": "string", "description": "Comma-separated tags."},
            "limit": {"type": "integer", "description": "Max rows (default 10)."},
            "row_id": {"type": "integer", "description": "Memory row id for 'update'/'delete'."},
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


# ---------------------------------------------------------------------------
# Provider
# ---------------------------------------------------------------------------

class ExcelLineProvider(MemoryProvider):
    """Excel-backed long-term memory; agent-driven logging + free-model indexer."""

    def __init__(self, config: dict | None = None):
        self._config = config or _load_plugin_config()
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
        from .store import ExcelLineStore
        self._store = ExcelLineStore(self._root)
        self._session_id = session_id

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
                       if n.endswith(".jsonl") or n.endswith(".json")]
            if not pending:
                return
            try:
                from .worker import index_while_logs_present
                free_model = self._config.get("free_model", "gemini-3.5-flash-lite")
                index_while_logs_present(
                    log_dir=log_dir, store_root=self._root,
                    free_model_fn=lambda p: _ask_free_model(p, free_model),
                    store=self._store)
            except Exception as e:
                logger.debug("excel_line lazy index failed: %s", e)
        except Exception:
            pass

    def prefetch(self, query: str, *, session_id: str = "") -> str:
        if not self._store or not query:
            return ""
        try:
            # NOTE: we intentionally do NOT call _drain_pending_logs() here.
            # That would run a synchronous LLM call on every user turn and block
            # the agent's response (Gemini review #4). Lazy indexing is instead
            # driven by the background indexer (_schedule_index / on_session_end),
            # so prefetch stays cheap and never calls the model.
            hits = self._store.search_index(query, limit=8)
            if not hits:
                return ""
            lines = ["## Excel-Line Memory (index matches)"]
            for h in hits:
                lines.append(f"- [{h['zone']}] {h['brief']}  (tags: {h['tags']})")
            for h in hits[:3]:
                details = self._store.read_zone(h["path"], limit=3)
                for d in details:
                    if d.get("content"):
                        lines.append(f"    • {d.get('title','')}: {d['content']}")
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
            return _ask_free_model(prompt, free_model)

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
            path = os.path.join(log_dir, f"session_raw_{sid}_{stamp}.jsonl")
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
        raw = _ask_free_model(prompt, self._config.get("free_model", "gemini-3.5-flash-lite"))
        facts = self._parse_facts(raw)
        if not facts:
            # LLM unavailable or returned no structured facts: fall back to a
            # safe, no-LLM backup — persist raw user turns so memory is never
            # lost even without the classifier.
            return self._fallback_store(turns)
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
                self._store.add(zone=zone, brief=brief, content=content,
                                title=brief[:40], tags=tags or "auto-extract")
                stored += 1
            except Exception as e:
                logger.debug("excel_line auto_extract store failed: %s", e)
        return stored

    def _fallback_store(self, turns: List[str]) -> int:
        """No-LLM backup: store each user turn verbatim into the knowledge zone
        so nothing is silently dropped when the classifier is unavailable."""
        stored = 0
        for t in turns:
            if not t.lower().startswith("user:"):
                continue
            text = t[len("user:"):].strip()
            if len(text) < 10:
                continue
            try:
                self._store.add(zone="knowledge", brief=text[:120],
                                content=text[:300], title=text[:40],
                                tags="auto-backup")
                stored += 1
            except Exception as e:
                logger.debug("excel_line fallback store failed: %s", e)
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
            if action == "search":
                hits = self._store.search_index(
                    args.get("query", ""), limit=int(args.get("limit", 10)))
                return json.dumps({"results": hits, "count": len(hits)})
            if action == "read":
                zone = args.get("zone", "knowledge")
                rows = self._store.read_zone(
                    self._store.zone_path(zone), limit=int(args.get("limit", 20)))
                return json.dumps({"zone": zone, "rows": rows, "count": len(rows)})
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
        zone = (args.get("zone") or "").strip()
        brief = (args.get("brief") or "").strip()
        content = (args.get("content") or "").strip()
        title = (args.get("title") or brief[:40]).strip()
        tags = (args.get("tags") or "").strip()
        if zone and (brief or content):
            try:
                mid = self._store.add(zone=zone, brief=brief[:120],
                                      content=content[:300], title=title[:40],
                                      tags=tags)
                return json.dumps({
                    "status": "stored", "id": mid, "zone": zone,
                    "note": "Written directly to the Excel store (no indexer needed).",
                })
            except Exception as e:
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
        log_name = f"{sid.replace('/', '_')}_i{in_seq}_o{out_seq}_{stamp}.jsonl"
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


def _ask_free_model(prompt: str, model: str) -> str:
    """Call a free Hermes model in-process to classify a record."""
    try:
        from agent.run_agent import quick_completion
        return quick_completion(prompt, model=model) or ""
    except Exception:
        return json.dumps({"zone": "knowledge", "brief": prompt[:80],
                           "title": "", "content": "", "tags": ""})


# ---------------------------------------------------------------------------
# Plugin entry point
# ---------------------------------------------------------------------------

def register(ctx) -> None:
    config = _load_plugin_config()
    provider = ExcelLineProvider(config=config)
    ctx.register_memory_provider(provider)
