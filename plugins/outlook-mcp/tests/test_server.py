"""MCP server tests: in-process against the fake Outlook object model (no Windows needed), then the
same server as a subprocess over stdio, exercising initialize / tools/list / tools/call / ping."""
import json
import os
import subprocess
import sys
import tempfile
from unittest import mock

HERE = os.path.dirname(os.path.abspath(__file__))
PLUGIN = os.path.dirname(HERE)
ROOT = os.path.dirname(os.path.dirname(PLUGIN))
sys.path.insert(0, PLUGIN)
sys.path.insert(0, os.path.join(ROOT, "plugins", "outlook", "tests"))  # fake_outlook lives with the skills plugin

import server  # noqa: E402
import outlook_com as oc  # noqa: E402  (the mcp copy of scripts/, via server's sys.path)
import fake_outlook as fo  # noqa: E402

assert oc.__file__.startswith(os.path.join(PLUGIN, "scripts")), oc.__file__
oc.set_namespace_for_tests(fo.build_fixture())
server._settings = lambda: {}  # no ~/.outlook-skills on the test machine influences the calls


def rpc(method, params=None, id_=1):
    return server.handle({"jsonrpc": "2.0", "id": id_, "method": method, "params": params or {}})


def call(name, **arguments):
    r = rpc("tools/call", {"name": name, "arguments": arguments})["result"]
    text = r["content"][0]["text"]
    return (json.loads(text) if not r["isError"] else text), r["isError"]


# --- initialize: known protocol versions are echoed, unknown ones get our newest
r = rpc("initialize", {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "t", "version": "0"}})["result"]
assert r["protocolVersion"] == "2024-11-05" and r["serverInfo"]["name"] == "outlook" and "tools" in r["capabilities"], r
r = rpc("initialize", {"protocolVersion": "2099-01-01"})["result"]
assert r["protocolVersion"] == server.PROTOCOL_VERSIONS[0]
assert "read" in r["instructions"].lower()
assert server.handle({"jsonrpc": "2.0", "method": "notifications/initialized"}) is None
assert rpc("ping")["result"] == {}
assert rpc("nope")["error"]["code"] == -32601

# --- tools/list: schemas derived from the scripts' parsers
tools = {t["name"]: t for t in rpc("tools/list")["result"]["tools"]}
assert set(tools) == {"search_mail", "get_thread", "list_calendar", "list_followups", "prepare_meeting", "find_attachments", "mailbox_overview", "get_status", "parse_msg_file"}, sorted(tools)
sp = tools["search_mail"]["inputSchema"]["properties"]
assert sp["from"]["type"] == "string" and sp["max"]["type"] == "integer" and sp["unread"]["type"] == "boolean", sp
assert sp["anyof"]["type"] == "array" and "out_file" not in sp and "out" not in sp and "query" in sp, sp
assert tools["list_followups"]["inputSchema"]["properties"]["direction"]["enum"] == ["sent", "received"]
assert tools["parse_msg_file"]["inputSchema"]["required"] == ["files"] and "format" not in tools["parse_msg_file"]["inputSchema"]["properties"]
assert "hasattachments" not in tools["find_attachments"]["inputSchema"]["properties"]
for t in tools.values():
    assert t["inputSchema"]["additionalProperties"] is False and t["description"]
    assert t["annotations"]["readOnlyHint"] == (t["name"] not in ("find_attachments", "parse_msg_file")), t["name"]

# --- tools/call against the fake mailbox (same expectations as the script tests)
out, err = call("search_mail", **{"from": "Cassie", "allstores": True})
assert not err and [m["EntryID"] for m in out["Results"]] == ["id1", "id5"], out
out, err = call("search_mail", **{"from": "Cassie"})
assert not err and out["Count"] == 0  # default store only
out, err = call("search_mail", anyof="報價,quote", allstores="true", allfolders=True, max="10")  # strings are coerced
assert not err and sorted(m["EntryID"] for m in out["Results"]) == ["id3", "id4"], out
out, err = call("get_thread", subject="合約草稿", store="20230731")
assert not err and [m["EntryID"] for m in out["Messages"]] == ["id2", "id1"], out
out, err = call("list_calendar", start="2026-09-16", days=1, store="20230731")
assert not err and len(out["Conflicts"]) == 1 and out["Count"] == 3, out
out, err = call("mailbox_overview", days=3650, store="20230731")
assert not err and out["TopSenders"][0]["Key"] == "cassie.tsai@contoso.com"
out, err = call("get_status", skipcom=True)
assert not err and "Warnings" in out

# --- the user's default store from settings applies when the call names none
server._settings = lambda: {"store": "20230731"}
out, err = call("search_mail", **{"from": "Cassie"})
assert not err and [m["EntryID"] for m in out["Results"]] == ["id1", "id5"], out
out, err = call("search_mail", **{"from": "Cassie", "allstores": True})
assert not err and out["Query"]["AllStores"] is True
server._settings = lambda: {}

# --- errors come back as isError results, never as a crash or a JSON-RPC error
text, err = call("get_thread")
assert err and "Provide" in text, text
text, err = call("search_mail", bogus=1)
assert err and "unknown argument 'bogus'" in text, text
text, err = call("list_followups", direction="sideways")
assert err and "must be one of" in text, text
text, err = call("search_mail", max="ten")
assert err and "integer" in text, text
text, err = call("no_such_tool")
assert err and "unknown tool" in text, text
assert rpc("tools/call", {"name": 5})["error"]["code"] == -32602

# --- a COM error (Outlook restarted) resets the connection and retries once
class ComError(Exception):
    pass
ComError.__name__ = "com_error"
calls = []
def flaky(a):
    calls.append(1)
    if len(calls) == 1:
        raise ComError("RPC server unavailable")
    return {"ok": True}
with mock.patch.object(server.outlook_calendar, "run", flaky), mock.patch.object(oc, "reset") as reset:
    out, err = call("list_calendar")
assert not err and out == {"ok": True} and len(calls) == 2 and reset.call_count == 1

# --- rerank: query without a configured gateway is a plain search with Rerank.Applied false
env_clear = {k: v for k, v in os.environ.items() if not k.startswith("OUTLOOK_RERANK")}
with mock.patch.dict(os.environ, env_clear, clear=True), mock.patch.object(server.rerank, "_plugin_settings", lambda: {}):
    out, err = call("search_mail", query="供應商的報價", allstores=True, allfolders=True)
assert not err and out["Rerank"]["Applied"] is False and out["Count"] == 8, out["Rerank"]

# with a gateway (settings rerank.gateway), the candidates are scored and sorted by Score, top = max (default 10)
def fake_probe(base, model, key, path, timeout):
    assert base == "http://gw:8000/v1"
    return "rerank", base + "/rerank", {"rerank": "ok"}
def fake_gateway(url, model, key, query, docs, endpoint, timeout, retries):
    assert query == "供應商的報價" and endpoint == "rerank" and all("Subject:" in d for d in docs)
    return [0.9 if "報價" in d else 0.1 for d in docs]
with mock.patch.dict(os.environ, env_clear, clear=True), \
     mock.patch.object(server.rerank, "_plugin_settings", lambda: {"OUTLOOK_RERANK_URL": "http://gw:8000/v1"}), \
     mock.patch.object(server.rerank, "probe_endpoints", fake_probe), \
     mock.patch.object(server.rerank, "call_gateway", fake_gateway):
    server._probe_cache.clear()
    out, err = call("search_mail", query="供應商的報價", allstores=True, allfolders=True, max=2)
    assert not err and out["Rerank"]["Applied"] is True and out["Rerank"]["Candidates"] == 8, out["Rerank"]
    assert out["Count"] == 2 and out["Results"][0]["EntryID"] == "id3" and out["Results"][0]["Score"] == 0.9 and out["Results"][1]["Score"] == 0.1, out["Results"]
    assert out["Query"]["Max"] >= 300  # candidate cap, not the result cap
    out, err = call("search_mail", query="供應商的報價", allstores=True, allfolders=True)
    assert not err and out["Count"] == 8 and [m["Score"] for m in out["Results"]] == sorted((m["Score"] for m in out["Results"]), reverse=True)
    # a gateway that does not answer: plain results, reason reported, probe failure cached
    server._probe_cache.clear()
    def dead_probe(*a):
        raise SystemExit("No working reranker endpoint at http://gw:8000/v1")
    with mock.patch.object(server.rerank, "probe_endpoints", dead_probe):
        out, err = call("search_mail", query="x", allstores=True)
        assert not err and out["Rerank"]["Applied"] is False and "No working reranker" in out["Rerank"]["Reason"]
    out, err = call("search_mail", query="x", allstores=True)  # cached failure, probe not repeated
    assert out["Rerank"]["Applied"] is False and "No working reranker" in out["Rerank"]["Reason"]
    server._probe_cache.clear()
print("in-process tests passed")

# ================= over stdio, as a real subprocess
eml = os.path.join(tempfile.mkdtemp(), "hello.eml")
with open(eml, "wb") as fh:
    fh.write("From: Alice <alice@example.com>\nTo: Bob <bob@example.com>\nSubject: Hi Bob\nDate: Tue, 15 Sep 2026 10:00:00 +0800\n"
             "Message-ID: <1@example.com>\nContent-Type: text/plain; charset=utf-8\n\nhello 中文\n".encode("utf-8"))
msgs = [
    {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "test", "version": "0"}}},
    {"jsonrpc": "2.0", "method": "notifications/initialized"},
    {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
    {"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "parse_msg_file", "arguments": {"files": [eml]}}},
    {"jsonrpc": "2.0", "id": 4, "method": "tools/call", "params": {"name": "search_mail", "arguments": {"from": "alice"}}},  # no Outlook here: isError, server survives
    {"jsonrpc": "2.0", "id": 5, "method": "ping"},
]
proc = subprocess.run([sys.executable, os.path.join(PLUGIN, "server.py")], input="".join(json.dumps(m) + "\n" for m in msgs).encode("utf-8"),
                      stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
lines = [json.loads(l) for l in proc.stdout.decode("utf-8").splitlines() if l.strip()]
by_id = {l["id"]: l for l in lines}
assert set(by_id) == {1, 2, 3, 4, 5}, lines
assert by_id[1]["result"]["protocolVersion"] == "2025-06-18"
assert len(by_id[2]["result"]["tools"]) == 9
parsed = json.loads(by_id[3]["result"]["content"][0]["text"])
assert by_id[3]["result"]["isError"] is False and parsed["subject"] == "Hi Bob" and parsed["from"][0]["address"] == "alice@example.com" and "中文" in parsed["body"], parsed
assert by_id[4]["result"]["isError"] is True or json.loads(by_id[4]["result"]["content"][0]["text"])  # real Outlook on a Windows dev box would answer
assert by_id[5]["result"] == {}
print("stdio tests passed")
