"""outlook_send.py against the fake Outlook: drafts resolve recipients and store the exact outgoing
text; send happens only after the confirmation window returned True, with the item matching the draft."""
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
sys.path.insert(0, os.path.dirname(__file__))
import outlook_com as oc  # noqa: E402
import fake_outlook as fo  # noqa: E402
import outlook_send as snd  # noqa: E402


class SendTest(unittest.TestCase):
    def setUp(self):
        self.ns = fo.build_fixture()
        self.app = self.ns.Application
        oc.set_namespace_for_tests(self.ns, self.app)
        home = str(Path(tempfile.mkdtemp()).resolve())
        self.env = mock.patch.dict(os.environ, {"HOME": home, "USERPROFILE": home})
        self.env.start()
        self.addCleanup(self.env.stop)
        self.addCleanup(oc.set_namespace_for_tests, None, None)

    def draft(self, *argv):
        return snd.run_draft(snd.draft_parser().parse_args(list(argv)))

    def send(self, d, clicked, confirm=None):
        with mock.patch.object(snd, "confirm_dialog", return_value=clicked) as dlg:
            r = snd.run_send(snd.send_parser().parse_args([d["id"], "-Confirm", confirm or d["confirm"]]))
        return r, dlg

    # ---- draft
    def test_new_mail_resolves_names_and_addresses(self):
        d = self.draft("-To", "Cassie Tsai", "pc.liao@contoso.com", "-Cc", "ben@contoso.com, cassie.tsai@contoso.com", "-Subject", "Hi", "-Body", "一句話。")
        self.assertEqual([r["Address"] for r in d["to"]], ["cassie.tsai@contoso.com", "pc.liao@contoso.com"])
        self.assertEqual(d["to"][0]["Name"], "Cassie Tsai")
        self.assertEqual([r["Address"] for r in d["cc"]], ["ben@contoso.com"])  # duplicate of a To recipient dropped
        self.assertEqual((d["mode"], d["subject"], d["status"], d["approver"]), ("new", "Hi", "draft", "Ben"))
        self.assertEqual(d["full_body"], "一句話。\n\n--\nDrafted by Claude, reviewed and approved by Ben.")
        self.assertEqual(d["confirm"], snd.token(d))
        self.assertTrue((snd.drafts_dir() / f"{d['id']}.json").is_file())

    def test_unknown_name_is_refused(self):
        with self.assertRaises(SystemExit) as cm:
            self.draft("-To", "Nobody Known", "-Subject", "x", "-Body", "y")
        self.assertIn("Nobody Known", str(cm.exception))

    def test_reply_and_reply_all(self):
        body = tempfile.mkstemp(suffix=".txt")[1]
        Path(body).write_text("收到，週五前回覆。\r\n", encoding="utf-8")
        d = self.draft("-ReplyTo", "id6", "-BodyFile", body)
        self.assertEqual((d["mode"], d["subject"]), ("reply", "RE: AI 人才發展：可以幫我看一下名單嗎？"))
        self.assertEqual([r["Address"] for r in d["to"]], ["pc.liao@contoso.com"])
        self.assertEqual(d["cc"], [])
        self.assertTrue(d["full_body"].startswith("收到，週五前回覆。\n\n--\nDrafted by Claude"))
        self.assertIn("-----Original Message-----\nFrom: PC Liao <pc.liao@contoso.com>", d["quote"])
        self.assertIn("麻煩看一下附件名單", d["full_body"])
        # reply-all on a mail sent to Ben and me: me is dropped, extra cc added, RE: not doubled
        d = self.draft("-ReplyTo", "id1", "-ReplyAll", "-Cc", "pc.liao@contoso.com", "-Body", "ok")
        self.assertEqual((d["mode"], d["subject"]), ("reply_all", "Re: 合約草稿 v3 - 法務意見"))
        self.assertEqual([r["Address"] for r in d["to"]], ["cassie.tsai@contoso.com", "ben@contoso.com"])
        self.assertEqual([r["Address"] for r in d["cc"]], ["pc.liao@contoso.com"])

    def test_settings_footer_and_no_quote(self):
        os.makedirs(os.path.join(os.environ["HOME"], ".outlook-skills"), exist_ok=True)
        with open(os.path.join(os.environ["HOME"], ".outlook-skills", "settings.json"), "w", encoding="utf-8") as fh:
            json.dump({"send": {"approver": "Ben Chen", "footer": "由 Claude 草擬，{approver} 核准後寄出", "quote_original": False}}, fh)
        d = self.draft("-ReplyTo", "id6", "-Body", "ok")
        self.assertEqual(d["full_body"], "ok\n\n由 Claude 草擬，Ben Chen 核准後寄出")
        self.assertEqual(d["quote"], "")

    # ---- send
    def test_send_only_after_the_click(self):
        d = self.draft("-To", "Cassie Tsai", "-Cc", "pc.liao@contoso.com", "-Subject", "Hi", "-Body", "一句話。")
        with self.assertRaises(SystemExit) as cm:
            self.send(d, clicked=False)
        self.assertIn("Cancelled", str(cm.exception))
        self.assertEqual(self.app.sent, [])
        self.assertEqual(snd.load(d["id"])["status"], "cancelled")
        with self.assertRaises(SystemExit):  # a cancelled draft cannot be retried
            self.send(d, clicked=True)
        self.assertEqual(self.app.sent, [])

        d = self.draft("-To", "Cassie Tsai", "-Cc", "pc.liao@contoso.com", "-Subject", "Hi", "-Body", "一句話。")
        r, dlg = self.send(d, clicked=True)
        spec = dlg.call_args[0][0]  # what the window showed: exactly the draft
        self.assertEqual(spec["text"], d["full_body"])
        self.assertIn("cassie.tsai@contoso.com", dict(spec["rows"])["收件者 To"])
        self.assertIn("pc.liao@contoso.com", dict(spec["rows"])["副本 Cc"])
        self.assertTrue(r["Sent"])
        self.assertEqual(len(self.app.sent), 1)
        m = self.app.sent[0]
        self.assertEqual([(x.Address, x.Type) for x in m.Recipients], [("cassie.tsai@contoso.com", 1), ("pc.liao@contoso.com", 2)])
        self.assertEqual((m.Subject, m.Body, m.BodyFormat), ("Hi", d["full_body"], 1))
        self.assertEqual(snd.load(d["id"])["status"], "sent")
        with self.assertRaises(SystemExit):  # no double send
            self.send(d, clicked=True)
        self.assertEqual(len(self.app.sent), 1)

    def test_reply_send_uses_reply_item_with_draft_recipients(self):
        d = self.draft("-ReplyTo", "id1", "-ReplyAll", "-Body", "ok")
        r, _ = self.send(d, clicked=True)
        m = self.app.sent[0]
        self.assertEqual(sorted(x.Address for x in m.Recipients), ["ben@contoso.com", "cassie.tsai@contoso.com"])
        self.assertEqual(m.Body, d["full_body"])
        self.assertEqual(r["Mode"], "reply_all")

    def test_wrong_or_stale_token_never_opens_the_window(self):
        d = self.draft("-To", "Cassie Tsai", "-Subject", "Hi", "-Body", "一句話。")
        with mock.patch.object(snd, "confirm_dialog", return_value=True) as dlg:
            with self.assertRaises(SystemExit):
                snd.run_send(snd.send_parser().parse_args([d["id"], "-Confirm", "deadbeef0000"]))
            # draft file edited after it was shown: stored token no longer matches the content
            p = snd.drafts_dir() / f"{d['id']}.json"
            data = json.loads(p.read_text(encoding="utf-8"))
            data["full_body"] += " PS: something else"
            p.write_text(json.dumps(data), encoding="utf-8")
            with self.assertRaises(SystemExit):
                snd.run_send(snd.send_parser().parse_args([d["id"], "-Confirm", d["confirm"]]))
        self.assertEqual(dlg.call_count, 0)
        self.assertEqual(self.app.sent, [])

    def test_item_that_drifts_from_the_draft_is_not_sent(self):
        d = self.draft("-To", "Cassie Tsai", "-Subject", "Hi", "-Body", "一句話。")
        real_add = fo.Recipients.Add

        def altered_add(self_, text):
            return real_add(self_, "pc.liao@contoso.com")  # Outlook "resolved" to someone else
        with mock.patch.object(fo.Recipients, "Add", altered_add), self.assertRaises(SystemExit) as cm:
            self.send(d, clicked=True)
        self.assertIn("To differs", str(cm.exception))
        self.assertEqual(self.app.sent, [])

    def test_attachments_are_hashed_shown_and_rechecked(self):
        d1 = tempfile.mkdtemp()
        f1, f2 = os.path.join(d1, "報價.pdf"), os.path.join(d1, "b.xlsx")
        Path(f1).write_bytes(b"%PDF-1.4 fake"); Path(f2).write_bytes(b"x" * 2048)
        with self.assertRaises(SystemExit):
            self.draft("-To", "Cassie Tsai", "-Subject", "Hi", "-Body", "x", "-Attach", os.path.join(d1, "missing.pdf"))
        d = self.draft("-To", "Cassie Tsai", "-Subject", "Hi", "-Body", "見附件。", "-Attach", f1, f2)
        self.assertEqual([a["Name"] for a in d["attachments"]], ["報價.pdf", "b.xlsx"])
        self.assertEqual(d["attachments"][1]["Size"], 2048)
        self.assertEqual(len(d["attachments"][0]["Sha256"]), 64)
        r, dlg = self.send(d, clicked=True)
        self.assertIn("報價.pdf (0 KB); b.xlsx (2 KB)", dict(dlg.call_args[0][0]["rows"])["附件 Attachments"])
        self.assertEqual(r["Attachments"], ["報價.pdf", "b.xlsx"])
        m = self.app.sent[0]
        self.assertEqual([(a.FileName, a.PathName) for a in m.Attachments], [("報價.pdf", f1), ("b.xlsx", f2)])
        # a file changed after the draft was shown: nothing is sent
        d = self.draft("-To", "Cassie Tsai", "-Subject", "Hi", "-Body", "見附件。", "-Attach", f1)
        Path(f1).write_bytes(b"%PDF-1.4 changed")
        with self.assertRaises(SystemExit) as cm:
            self.send(d, clicked=True)
        self.assertIn("changed since the draft", str(cm.exception))
        self.assertEqual(len(self.app.sent), 1)
        # the attachment list is part of the token
        d = self.draft("-To", "Cassie Tsai", "-Subject", "Hi", "-Body", "x")
        self.assertNotEqual(d["confirm"], self.draft("-To", "Cassie Tsai", "-Subject", "Hi", "-Body", "x", "-Attach", f2)["confirm"])
        # over the size limit
        big = os.path.join(d1, "big.bin")
        with open(big, "wb") as fh:
            fh.truncate(snd.ATTACH_MAX_MB * 1024 * 1024 + 1)
        with self.assertRaises(SystemExit) as cm:
            self.draft("-To", "Cassie Tsai", "-Subject", "Hi", "-Body", "x", "-Attach", big)
        self.assertIn("MB limit", str(cm.exception))

    def test_list_show_discard(self):
        d = self.draft("-To", "Cassie Tsai", "-Subject", "Hi", "-Body", "x")
        lst = snd.run_list(None)
        self.assertEqual((lst["Count"], lst["Drafts"][0]["id"]), (1, d["id"]))
        self.assertEqual(snd.run_show(type("A", (), {"id": d["id"]})())["subject"], "Hi")
        self.assertEqual(snd.run_discard(type("A", (), {"id": d["id"]})())["Status"], "discarded")
        with self.assertRaises(SystemExit):
            self.send(d, clicked=True)


if __name__ == "__main__":
    unittest.main()
