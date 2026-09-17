"""MCP server tests: in-process against the fake Outlook object model (no Windows needed), then the
same server as a subprocess over stdio, exercising initialize / tools/list / tools/call / ping.

    python -m unittest discover -s plugins/outlook-mcp/tests -v
"""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from unittest import mock

HERE = os.path.dirname(os.path.abspath(__file__))
PLUGIN = os.path.dirname(HERE)
ROOT = os.path.dirname(os.path.dirname(PLUGIN))
sys.path.insert(0, PLUGIN)
sys.path.insert(0, os.path.join(ROOT, "plugins", "outlook", "tests"))  # fake_outlook lives with the skills plugin

import server  # noqa: E402
import outlook_com as oc  # noqa: E402  (the mcp copy of scripts/, via server's sys.path)
import fake_outlook as fo  # noqa: E402

ALL_TOOLS = {"search_mail", "get_thread", "list_calendar", "list_followups", "find_attachments", "mailbox_overview", "get_status", "parse_msg_file"}


def rpc(method, params=None, id_=1):
    return server.handle({"jsonrpc": "2.0", "id": id_, "method": method, "params": params or {}})


def call(name, **arguments):
    r = rpc("tools/call", {"name": name, "arguments": arguments})["result"]
    text = r["content"][0]["text"]
    return (json.loads(text) if not r["isError"] else text), r["isError"]


def ids(rows):
    return [m["EntryID"] for m in rows]


class ServerTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        assert oc.__file__.startswith(os.path.join(PLUGIN, "scripts")), oc.__file__
        oc.set_namespace_for_tests(fo.build_fixture())

    def setUp(self):
        self._settings = mock.patch.object(server, "_settings", lambda: {})  # no ~/.outlook-skills on the test machine influences the calls
        self._settings.start()
        self.addCleanup(self._settings.stop)
        server._probe_cache.clear()


class ProtocolTest(ServerTest):
    def test_initialize_echoes_known_versions(self):
        r = rpc("initialize", {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "t", "version": "0"}})["result"]
        self.assertEqual(r["protocolVersion"], "2024-11-05")
        self.assertEqual(r["serverInfo"]["name"], "outlook")
        self.assertIn("tools", r["capabilities"])
        r = rpc("initialize", {"protocolVersion": "2099-01-01"})["result"]
        self.assertEqual(r["protocolVersion"], server.PROTOCOL_VERSIONS[0])
        self.assertIn("read", r["instructions"].lower())

    def test_version_comes_from_plugin_json(self):
        with open(os.path.join(PLUGIN, ".claude-plugin", "plugin.json"), encoding="utf-8") as fh:
            self.assertEqual(server.SERVER_VERSION, json.load(fh)["version"])

    def test_notifications_ping_unknown(self):
        self.assertIsNone(server.handle({"jsonrpc": "2.0", "method": "notifications/initialized"}))
        self.assertEqual(rpc("ping")["result"], {})
        self.assertEqual(rpc("nope")["error"]["code"], -32601)
        self.assertEqual(rpc("tools/call", {"name": 5})["error"]["code"], -32602)

    def test_tools_list_schemas_derive_from_parsers(self):
        tools = {t["name"]: t for t in rpc("tools/list")["result"]["tools"]}
        self.assertEqual(set(tools), ALL_TOOLS)
        sp = tools["search_mail"]["inputSchema"]["properties"]
        self.assertEqual((sp["from"]["type"], sp["max"]["type"], sp["unread"]["type"], sp["anyof"]["type"]), ("string", "integer", "boolean", "array"))
        self.assertNotIn("out_file", sp)
        self.assertNotIn("out", sp)
        self.assertIn("query", sp)
        self.assertEqual(tools["list_followups"]["inputSchema"]["properties"]["direction"]["enum"], ["sent", "received", "both"])
        self.assertEqual(tools["parse_msg_file"]["inputSchema"]["required"], ["files"])
        self.assertNotIn("format", tools["parse_msg_file"]["inputSchema"]["properties"])
        self.assertNotIn("hasattachments", tools["find_attachments"]["inputSchema"]["properties"])
        self.assertIn("entryid", tools["find_attachments"]["inputSchema"]["properties"])
        for t in tools.values():
            self.assertIs(t["inputSchema"]["additionalProperties"], False)
            self.assertTrue(t["description"])
            self.assertEqual(t["annotations"]["readOnlyHint"], t["name"] not in ("find_attachments", "parse_msg_file"), t["name"])


class CallTest(ServerTest):
    def test_calls_against_the_fake_mailbox(self):
        out, err = call("search_mail", **{"from": "Cassie", "allstores": True})
        self.assertFalse(err)
        self.assertEqual(ids(out["Results"]), ["id1", "id5"])
        out, err = call("search_mail", **{"from": "Cassie"})
        self.assertEqual(out["Count"], 0)  # default store only
        out, err = call("search_mail", anyof="報價,quote", allstores="true", allfolders=True, max="10")  # strings are coerced
        self.assertEqual(sorted(ids(out["Results"])), ["id3", "id4"])
        out, err = call("get_thread", subject="合約草稿", store="20230731")
        self.assertEqual(ids(out["Messages"]), ["id2", "id1"])
        out, err = call("find_attachments", entryid="id7", ext="png")
        self.assertEqual((out["MailsScanned"], [r["FileName"] for r in out["Results"]]), (1, ["image001.png"]))
        out, err = call("list_calendar", start="2026-09-16", days=1, store="20230731")
        self.assertEqual((len(out["Conflicts"]), out["Count"]), (1, 3))
        out, err = call("mailbox_overview", days=3650, store="20230731")
        self.assertEqual(out["TopSenders"][0]["Key"], "cassie.tsai@contoso.com")
        out, err = call("get_status", skipcom=True)
        self.assertFalse(err)
        self.assertIn("Warnings", out)
        out, err = call("list_followups", direction="both", store="20230731", lookback=3650)
        self.assertFalse(err)
        self.assertIn("Sent", out)
        self.assertIn("Received", out)

    def test_default_store_from_settings(self):
        with mock.patch.object(server, "_settings", lambda: {"store": "20230731"}):
            out, err = call("search_mail", **{"from": "Cassie"})
            self.assertEqual(ids(out["Results"]), ["id1", "id5"])
            out, err = call("search_mail", **{"from": "Cassie", "allstores": True})
            self.assertIs(out["Query"]["AllStores"], True)

    def test_errors_are_iserror_results(self):
        text, err = call("get_thread")
        self.assertTrue(err and "Provide" in text, text)
        text, err = call("search_mail", bogus=1)
        self.assertTrue(err and "unknown argument 'bogus'" in text, text)
        text, err = call("list_followups", direction="sideways")
        self.assertTrue(err and "must be one of" in text, text)
        text, err = call("search_mail", max="ten")
        self.assertTrue(err and "integer" in text, text)
        text, err = call("no_such_tool")
        self.assertTrue(err and "unknown tool" in text, text)

    def test_com_error_resets_and_retries_once(self):
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
        self.assertFalse(err)
        self.assertEqual((out, len(calls), reset.call_count), ({"ok": True}, 2, 1))


class RerankTest(ServerTest):
    def setUp(self):
        super().setUp()
        env_clear = {k: v for k, v in os.environ.items() if not k.startswith("OUTLOOK_RERANK")}
        self.env = mock.patch.dict(os.environ, env_clear, clear=True)
        self.env.start()
        self.addCleanup(self.env.stop)

    def test_query_without_gateway_is_a_plain_search(self):
        with mock.patch.object(server.rerank, "_plugin_settings", lambda: {}):
            out, err = call("search_mail", query="供應商的報價", allstores=True, allfolders=True)
        self.assertFalse(err)
        self.assertIs(out["Rerank"]["Applied"], False)
        self.assertEqual(out["Count"], 8)

    def test_gateway_from_settings_scores_and_sorts(self):
        def fake_probe(base, model, key, path, timeout):
            self.assertEqual(base, "http://gw:8000/v1")
            return "rerank", base + "/rerank", {"rerank": "ok"}

        def fake_gateway(url, model, key, query, docs, endpoint, timeout, retries):
            self.assertEqual((query, endpoint), ("供應商的報價", "rerank"))
            self.assertTrue(all("Subject:" in d for d in docs))
            return [0.9 if "報價" in d else 0.1 for d in docs]
        with mock.patch.object(server.rerank, "_plugin_settings", lambda: {"OUTLOOK_RERANK_URL": "http://gw:8000/v1"}), \
             mock.patch.object(server.rerank, "probe_endpoints", fake_probe), \
             mock.patch.object(server.rerank, "call_gateway", fake_gateway):
            out, err = call("search_mail", query="供應商的報價", allstores=True, allfolders=True, max=2)
            self.assertFalse(err)
            self.assertIs(out["Rerank"]["Applied"], True)
            self.assertEqual(out["Rerank"]["Candidates"], 8)
            self.assertEqual((out["Count"], out["Results"][0]["EntryID"], out["Results"][0]["Score"], out["Results"][1]["Score"]), (2, "id3", 0.9, 0.1))
            self.assertGreaterEqual(out["Query"]["Max"], 300)  # candidate cap, not the result cap
            out, err = call("search_mail", query="供應商的報價", allstores=True, allfolders=True)
            self.assertEqual(out["Count"], 8)
            self.assertEqual([m["Score"] for m in out["Results"]], sorted((m["Score"] for m in out["Results"]), reverse=True))

    def test_dead_gateway_is_reported_and_cached(self):
        def dead_probe(*a):
            raise SystemExit("No working reranker endpoint at http://gw:8000/v1")
        with mock.patch.object(server.rerank, "_plugin_settings", lambda: {"OUTLOOK_RERANK_URL": "http://gw:8000/v1"}):
            with mock.patch.object(server.rerank, "probe_endpoints", dead_probe):
                out, err = call("search_mail", query="x", allstores=True)
            self.assertFalse(err)
            self.assertIs(out["Rerank"]["Applied"], False)
            self.assertIn("No working reranker", out["Rerank"]["Reason"])
            out, err = call("search_mail", query="x", allstores=True)  # cached failure, probe not repeated
            self.assertIn("No working reranker", out["Rerank"]["Reason"])


class StdioTest(unittest.TestCase):
    def test_real_subprocess(self):
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
        self.assertEqual(proc.returncode, 0, proc.stderr.decode("utf-8", "replace"))
        lines = [json.loads(line) for line in proc.stdout.decode("utf-8").splitlines() if line.strip()]
        by_id = {line["id"]: line for line in lines}
        self.assertEqual(set(by_id), {1, 2, 3, 4, 5}, lines)
        self.assertEqual(by_id[1]["result"]["protocolVersion"], "2025-06-18")
        self.assertEqual(len(by_id[2]["result"]["tools"]), len(ALL_TOOLS))
        parsed = json.loads(by_id[3]["result"]["content"][0]["text"])
        self.assertIs(by_id[3]["result"]["isError"], False)
        self.assertEqual((parsed["subject"], parsed["from"][0]["address"]), ("Hi Bob", "alice@example.com"))
        self.assertIn("中文", parsed["body"])
        self.assertTrue(by_id[4]["result"]["isError"] is True or json.loads(by_id[4]["result"]["content"][0]["text"]))  # real Outlook on a Windows dev box would answer
        self.assertEqual(by_id[5]["result"], {})


if __name__ == "__main__":
    unittest.main()
