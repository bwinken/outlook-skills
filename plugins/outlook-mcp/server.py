#!/usr/bin/env python3
"""READ-ONLY MCP server for a local Windows Outlook (Classic) mailbox, stdio transport.

Standard library only (plus pywin32 for Outlook COM). It wraps the same scripts as the `outlook`
skills plugin: scripts/ next to this file is a verbatim copy of plugins/outlook/scripts kept in
sync by tools/sync_scripts.py. Every tool's input schema is derived from that script's argparse
parser, so the tools cannot drift from the scripts.

    python server.py                                   # MCP over stdin/stdout (JSON-RPC 2.0, one message per line)
    python server.py --list                            # print the tool list and exit, no Outlook access
    python server.py --call search_mail '{"from": "alice", "max": 5}'   # one tool call, for debugging

Nothing here writes to Outlook: no Save, Send, Move, Delete, no property setters, items are never
marked read. The only files written are the ones the user asks for (find_attachments saveto,
parse_msg_file extract_to). See README.md.
"""
import argparse
import json
import os
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "scripts"))

import outlook_com as oc  # noqa: E402
import outlook_attachments  # noqa: E402
import outlook_calendar  # noqa: E402
import outlook_followup  # noqa: E402
import outlook_meeting_prep  # noqa: E402
import outlook_overview  # noqa: E402
import outlook_search  # noqa: E402
import outlook_status  # noqa: E402
import outlook_thread  # noqa: E402
import read_msg  # noqa: E402
import rerank  # noqa: E402
import settings as ps  # noqa: E402

SERVER_NAME = "outlook"
SERVER_VERSION = "0.1.0"
PROTOCOL_VERSIONS = ("2025-06-18", "2025-03-26", "2024-11-05")  # newest first; the client's choice is echoed when known

INSTRUCTIONS = (
    "Every tool only reads: nothing in Outlook is marked read, moved, sent or deleted. "
    "Needs Windows with Classic Outlook and pywin32; parse_msg_file works on any OS. "
    "Dates are local time, YYYY-MM-DD or YYYY-MM-DDTHH:MM; 'before' is exclusive. "
    "Folder names may be English (Inbox, Sent Items) or as shown in a localized Outlook (收件匣, 寄件備份). "
    "When mail lives in a .pst, pass store (its display name from get_status) or allstores=true; "
    "a `store` default from ~/.outlook-skills/settings.json is applied automatically when neither is given. "
    "EntryID values from search_mail go to get_thread; appointment EntryIDs from list_calendar go to prepare_meeting. "
    "Prefer narrow filters (from, after, anyof) and a small max over includebody. "
    "search_mail accepts a natural-language `query`: when the user configured a reranker gateway "
    "(settings rerank.gateway or OUTLOOK_RERANK_URL) the candidates are reranked by it automatically; "
    "otherwise the query is ignored and plain substring results are returned (see the Rerank field)."
)


class ToolError(Exception):
    pass


# ---------------------------------------------------------------- settings / rerank
def _settings() -> dict:
    try:
        return ps.resolve()[0] or {}
    except Exception:
        return {}


_probe_cache = {}  # base url -> (endpoint, url) or ("error", message, timestamp)
PROBE_RETRY_SECONDS = 300


def _rerank_gateway():
    """Only an explicitly configured gateway counts: settings rerank.gateway or OUTLOOK_RERANK_URL.
    The ANTHROPIC_BASE_URL fallback that rerank.py offers on the command line is deliberately not
    used here, so mail previews are never sent to a gateway the user did not pick for reranking."""
    url = os.environ.get("OUTLOOK_RERANK_URL") or rerank._plugin_settings().get("OUTLOOK_RERANK_URL")
    if not url:
        return None
    base, model, key, _ = rerank.resolve_config(argparse.Namespace(gateway=url, model=None, api_key=None))
    cached = _probe_cache.get(base)
    if cached and cached[0] == "error" and time.time() - cached[2] < PROBE_RETRY_SECONDS:
        raise ToolError(cached[1])
    if not cached or cached[0] == "error":
        try:
            ep, ep_url, _ = rerank.probe_endpoints(base, model, key, None, 15.0)
        except SystemExit as e:
            _probe_cache[base] = ("error", str(e), time.time())
            raise ToolError(str(e))
        _probe_cache[base] = (ep, ep_url)
        cached = _probe_cache[base]
    return {"base": base, "endpoint": cached[0], "url": cached[1], "model": model, "key": key}


def _rerank_results(query: str, results: list, top: int, doc_chars: int = 600, batch: int = 30):
    gw = _rerank_gateway()
    if gw is None:
        return None, {"Applied": False, "Reason": "no reranker gateway configured (settings rerank.gateway or OUTLOOK_RERANK_URL)"}
    docs = [rerank.doc_text(m, doc_chars) for m in results]
    scores = []
    for i in range(0, len(docs), batch):
        scores.extend(rerank.call_gateway(gw["url"], gw["model"], gw["key"], query, docs[i:i + batch], gw["endpoint"], 60.0, 2))
    ranked = []
    for m, s in zip(results, scores):
        if s is None:
            continue
        item = dict(m)
        item["Score"] = round(s, 4)
        ranked.append(item)
    ranked.sort(key=lambda x: x["Score"], reverse=True)
    info = {"Applied": True, "Query": query, "Gateway": gw["url"], "Endpoint": gw["endpoint"], "Model": gw["model"],
            "Candidates": len(results), "Batches": (len(docs) + batch - 1) // batch}
    return ranked[:top], info


# ---------------------------------------------------------------- tool wrappers
def run_search(a, arguments):
    query = (getattr(a, "query", "") or "").strip()
    if not query:
        return outlook_search.run(a)
    cfg = _settings().get("search") or {}
    top = a.max if "max" in arguments else 10
    a.max = max(int(cfg.get("max_candidates") or 300), top)
    a.previewlength = max(a.previewlength, 500)
    out = outlook_search.run(a)
    try:
        ranked, info = _rerank_results(query, out["Results"], top)
    except (ToolError, SystemExit) as e:
        ranked, info = None, {"Applied": False, "Reason": str(e)}
    if ranked is not None:
        out["Results"], out["Count"] = ranked, len(ranked)
    out["Rerank"] = info
    return out


def run_parse_msg(a, arguments):
    r = read_msg.run(a)
    return r[0] if len(r) == 1 else r


def _search_parser():
    ap = outlook_search.parser()
    ap.add_argument("--query", dest="query", default="", help="what the user is looking for, in their own words; used only to rerank the candidates when a reranker gateway is configured (then max = how many reranked results to return, default 10). Still narrow with from / after / anyof")
    return ap


TOOLS = [
    {"name": "search_mail", "parser": _search_parser, "run": run_search, "writes": False,
     "description": "Search mail by sender, recipient, subject, body text, date range, unread state or attachments. Newest first, "
                    "previews only unless includebody. Returns EntryID, folder, from, subject, dates and attachment names. "
                    "With query and a configured reranker gateway the results are reranked by relevance (Score field)."},
    {"name": "get_thread", "parser": outlook_thread.parser, "run": lambda a, _: outlook_thread.run(a), "writes": False,
     "description": "Every message of one conversation, oldest first, with full bodies and recipients. Select by entryid (from search_mail), "
                    "conversationid, or a subject substring (newest match in folder anchors the thread)."},
    {"name": "list_calendar", "parser": outlook_calendar.parser, "run": lambda a, _: outlook_calendar.run(a), "writes": False,
     "description": "Calendar items in a date range (default today), recurrences expanded, overlapping items listed in Conflicts, "
                    "response status per meeting. Items marked Free are left out unless includefree. Use for agendas and for computing free slots."},
    {"name": "list_followups", "parser": outlook_followup.parser, "run": lambda a, _: outlook_followup.run(a), "writes": False,
     "description": "Mails waiting for a reply. direction=sent: the user wrote and nobody answered for `days` (default 3). "
                    "direction=received: someone wrote to the user and the user has not answered (default 2 days); LooksLikeQuestion flags requests."},
    {"name": "prepare_meeting", "parser": outlook_meeting_prep.parser, "run": lambda a, _: outlook_meeting_prep.run(a), "writes": False,
     "description": "Briefing before a meeting: the appointment, its attendees, recent mail exchanged with each attendee, mail about the "
                    "meeting's subject and attachments seen along the way. Default: the next upcoming meeting; or subject / entryid."},
    {"name": "find_attachments", "parser": outlook_attachments.parser, "run": lambda a, _: outlook_attachments.run(a), "writes": True,
     "skip": ("hasattachments", "includebody"),
     "description": "Find attachments across mails by file name, extension or size, combined with every search_mail filter. Rows are attachments. "
                    "saveto copies the matching files into that folder on disk (ask the user for the folder first); Outlook is never modified."},
    {"name": "mailbox_overview", "parser": outlook_overview.parser, "run": lambda a, _: outlook_overview.run(a), "writes": False,
     "description": "Aggregates over the last `days`: folders with counts, top senders and recipients, frequent topics, likely newsletters, "
                    "recurring meetings. No message bodies. Use to learn who matters and where mail lives."},
    {"name": "get_status", "parser": outlook_status.parser, "run": lambda a, _: outlook_status.run(a), "writes": False,
     "description": "Outlook version, profiles, accounts, stores with their .pst/.ost paths and sizes, top-level folder counts, warnings. "
                    "skipcom=true inspects only the registry and disk without talking to Outlook."},
    {"name": "parse_msg_file", "parser": read_msg.parser, "run": run_parse_msg, "writes": True, "skip": ("format",),
     "description": "Parse .msg or .eml files on disk without Outlook: subject, addresses, date, body, transport headers (headers=true) and the "
                    "attachment list. extract_to copies the attachments into a folder. Works on any OS."},
]
TOOL_BY_NAME = {t["name"]: t for t in TOOLS}
_SKIP_DESTS = ("out_file", "help")


def _prop_name(dest: str) -> str:
    return dest.rstrip("_")  # argparse dest "from_" (keyword) -> "from"


def _actions(parser, tool):
    skip = _SKIP_DESTS + tuple(tool.get("skip", ()))
    return [a for a in parser._actions if not isinstance(a, argparse._HelpAction) and a.dest not in skip and a.help != argparse.SUPPRESS]


def input_schema(tool) -> dict:
    props, required = {}, []
    for act in _actions(tool["parser"](), tool):
        p = {}
        if isinstance(act, argparse._StoreTrueAction):
            p["type"] = "boolean"
        elif act.nargs in ("+", "*"):
            p["type"] = "array"
            p["items"] = {"type": "string"}
        elif act.type is int:
            p["type"] = "integer"
        elif act.type is float:
            p["type"] = "number"
        else:
            p["type"] = "string"
        if act.choices:
            p["enum"] = list(act.choices)
        desc = (act.help or "").strip()
        if act.default not in (None, "", False, []) and not act.choices:
            desc = f"{desc} (default {act.default})" if desc else f"default {act.default}"
        elif act.choices and act.default:
            desc = f"{desc} (default {act.default})" if desc else f"default {act.default}"
        if desc:
            p["description"] = desc
        if act.required:
            required.append(_prop_name(act.dest))
        props[_prop_name(act.dest)] = p
    schema = {"type": "object", "properties": props, "additionalProperties": False}
    if required:
        schema["required"] = required
    return schema


def tool_descriptor(tool) -> dict:
    return {
        "name": tool["name"],
        "description": tool["description"],
        "inputSchema": input_schema(tool),
        "annotations": {"readOnlyHint": not tool["writes"], "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
    }


def _coerce(act, value):
    name = _prop_name(act.dest)
    if isinstance(act, argparse._StoreTrueAction):
        if isinstance(value, str):
            return value.strip().lower() in ("1", "true", "yes", "y")
        return bool(value)
    if act.nargs in ("+", "*"):
        if isinstance(value, str):
            value = [value]
        if not isinstance(value, list):
            raise ToolError(f"{name}: expected a list of strings")
        return [str(v) for v in value]
    if act.type is int:
        try:
            return int(value)
        except (TypeError, ValueError):
            raise ToolError(f"{name}: expected an integer, got {value!r}")
    if act.type is float:
        try:
            return float(value)
        except (TypeError, ValueError):
            raise ToolError(f"{name}: expected a number, got {value!r}")
    if value is None:
        return act.default
    if isinstance(value, (dict, list)):
        raise ToolError(f"{name}: expected a string")
    value = str(value)
    if act.choices and value not in act.choices:
        raise ToolError(f"{name}: must be one of {', '.join(act.choices)}")
    return value


def build_args(tool, arguments: dict):
    """Namespace with the parser's defaults, overridden by the JSON arguments (no argv round trip)."""
    parser = tool["parser"]()
    ns = argparse.Namespace()
    by_name = {}
    for act in parser._actions:
        if isinstance(act, argparse._HelpAction):
            continue
        setattr(ns, act.dest, act.default)
        if act.dest not in _SKIP_DESTS + tuple(tool.get("skip", ())):
            by_name[_prop_name(act.dest)] = act
    for k, v in (arguments or {}).items():
        act = by_name.get(k)
        if act is None:
            raise ToolError(f"unknown argument '{k}' for {tool['name']}; known: {', '.join(sorted(by_name))}")
        setattr(ns, act.dest, _coerce(act, v))
    for act in by_name.values():
        if act.required and getattr(ns, act.dest) in (None, "", []):
            raise ToolError(f"missing required argument '{_prop_name(act.dest)}'")
    # the user's default store (mail in a .pst) unless the call names a store or asks for every store
    if hasattr(ns, "store") and not ns.store and not getattr(ns, "allstores", False):
        st = _settings().get("store")
        if st:
            ns.store = str(st)
    return ns


def _is_com_error(e: Exception) -> bool:
    return type(e).__name__ == "com_error" or type(e).__module__ in ("pywintypes", "pythoncom")


def call_tool(name: str, arguments: dict):
    tool = TOOL_BY_NAME.get(name)
    if tool is None:
        raise ToolError(f"unknown tool '{name}'; tools: {', '.join(TOOL_BY_NAME)}")
    a = build_args(tool, arguments)
    try:
        return tool["run"](a, arguments or {})
    except Exception as e:
        if not _is_com_error(e):
            raise
        oc.reset()  # Outlook closed or restarted under us: reconnect once
        a = build_args(tool, arguments)
        return tool["run"](a, arguments or {})


# ---------------------------------------------------------------- JSON-RPC
def _result(id_, result):
    return {"jsonrpc": "2.0", "id": id_, "result": result}


def _error(id_, code, message, data=None):
    err = {"code": code, "message": message}
    if data is not None:
        err["data"] = data
    return {"jsonrpc": "2.0", "id": id_, "error": err}


def _tool_result(name, arguments):
    try:
        out = call_tool(name, arguments)
        return {"content": [{"type": "text", "text": json.dumps(out, ensure_ascii=False)}], "isError": False}
    except (ToolError, SystemExit) as e:  # the scripts report usage errors and missing Outlook with SystemExit
        return {"content": [{"type": "text", "text": f"{name}: {e}"}], "isError": True}
    except Exception as e:
        sys.stderr.write(traceback.format_exc())
        return {"content": [{"type": "text", "text": f"{name}: {type(e).__name__}: {e}"}], "isError": True}


def handle(msg):
    """One JSON-RPC message in, a response dict out (None for notifications)."""
    if not isinstance(msg, dict):
        return _error(None, -32600, "invalid request")
    method, id_, params = msg.get("method"), msg.get("id"), msg.get("params") or {}
    if method is None:
        return None  # a response to something we never sent
    if id_ is None:
        return None  # notifications (initialized, cancelled, progress) need no reply
    if method == "initialize":
        want = str(params.get("protocolVersion") or "")
        return _result(id_, {
            "protocolVersion": want if want in PROTOCOL_VERSIONS else PROTOCOL_VERSIONS[0],
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
            "instructions": INSTRUCTIONS,
        })
    if method == "ping":
        return _result(id_, {})
    if method == "tools/list":
        return _result(id_, {"tools": [tool_descriptor(t) for t in TOOLS]})
    if method == "tools/call":
        name = params.get("name")
        arguments = params.get("arguments") or {}
        if not isinstance(name, str) or not isinstance(arguments, dict):
            return _error(id_, -32602, "tools/call needs a string name and an object arguments")
        return _result(id_, _tool_result(name, arguments))
    return _error(id_, -32601, f"method not found: {method}")


def serve():
    inp, out = sys.stdin.buffer, sys.stdout.buffer
    sys.stdout = sys.stderr  # a stray print() anywhere must not corrupt the protocol stream

    def send(obj):
        out.write((json.dumps(obj, ensure_ascii=False) + "\n").encode("utf-8"))
        out.flush()

    for raw in iter(inp.readline, b""):
        line = raw.strip()
        if not line:
            continue
        try:
            msg = json.loads(line.decode("utf-8"))
        except Exception as e:
            send(_error(None, -32700, f"parse error: {e}"))
            continue
        if isinstance(msg, list):  # batch (protocol 2025-03-26)
            replies = [r for r in (handle(m) for m in msg) if r is not None]
            if replies:
                send(replies)
            continue
        reply = handle(msg)
        if reply is not None:
            send(reply)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--list", action="store_true", help="print the tool list as JSON and exit")
    ap.add_argument("--call", nargs=2, metavar=("TOOL", "JSON"), help="run one tool with a JSON object of arguments and exit")
    a = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if a.list:
        print(json.dumps([tool_descriptor(t) for t in TOOLS], ensure_ascii=False, indent=2))
        return
    if a.call:
        r = _tool_result(a.call[0], json.loads(a.call[1]))
        print(r["content"][0]["text"])
        sys.exit(1 if r["isError"] else 0)
    serve()


if __name__ == "__main__":
    main()
