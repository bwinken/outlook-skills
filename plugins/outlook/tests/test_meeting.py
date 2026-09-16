"""outlook_meeting.py against the fake Outlook: drafts resolve attendees, flag overlaps and store the
exact invitation; the item is created only after the confirmation window returned True."""
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
import outlook_meeting as mtg  # noqa: E402


class MeetingTest(unittest.TestCase):
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
        return mtg.run_draft(mtg.draft_parser().parse_args(list(argv)))

    def send(self, d, clicked, confirm=None):
        with mock.patch.object(snd, "confirm_dialog", return_value=clicked) as dlg:
            r = mtg.run_send(mtg.send_parser().parse_args([d["id"], "-Confirm", confirm or d["confirm"]]))
        return r, dlg

    def test_draft_resolves_attendees_and_finds_overlaps(self):
        d = self.draft("-Subject", "Q3 review", "-Start", "2026-09-16T14:30", "-Duration", "45", "-Attendees", "Cassie Tsai", "me@contoso.com",
                       "-Optional", "pc.liao@contoso.com", "cassie.tsai@contoso.com", "-Location", "Room 3", "-Body", "議程：預算", "-Store", "20230731")
        self.assertEqual((d["kind"], d["mode"], d["start"], d["end"], d["duration_minutes"]), ("meeting", "meeting", "2026-09-16T14:30:00", "2026-09-16T15:15:00", 45))
        self.assertEqual([r["Address"] for r in d["required"]], ["cassie.tsai@contoso.com"])  # me dropped
        self.assertEqual([r["Address"] for r in d["optional"]], ["pc.liao@contoso.com"])  # already required -> dropped
        self.assertEqual(d["full_body"], "議程：預算\n\n--\nDrafted by Claude, reviewed and approved by Ben.")
        self.assertEqual(sorted(c["Subject"] for c in d["conflicts"]), ["1:1 with Bob", "供應商簡報"])
        self.assertEqual(d["confirm"], snd.token(d))
        self.assertTrue((snd.drafts_dir() / f"{d['id']}.json").is_file())

    def test_appointment_without_attendees_and_default_duration(self):
        d = self.draft("-Subject", "Focus", "-Start", "2026-09-17T09:00", "-Store", "20230731")
        self.assertEqual((d["mode"], d["end"], d["conflicts"], d["required"]), ("appointment", "2026-09-17T10:00:00", [], []))
        self.assertEqual(d["full_body"], "--\nDrafted by Claude, reviewed and approved by Ben.")

    def test_bad_input(self):
        for argv in (["-Subject", "x"], ["-Start", "2026-09-17T09:00"], ["-Subject", "x", "-Start", "2026-09-17T09:00", "-End", "2026-09-17T08:00"],
                     ["-Subject", "x", "-Start", "2026-09-17T09:00", "-Attendees", "Nobody Known"]):
            with self.assertRaises(SystemExit):
                self.draft(*argv)

    def test_meeting_is_sent_only_after_the_click(self):
        d = self.draft("-Subject", "Q3 review", "-Start", "2026-09-18T14:00", "-Attendees", "Cassie Tsai", "-Optional", "pc.liao@contoso.com", "-Location", "Teams", "-Body", "議程")
        with self.assertRaises(SystemExit) as cm:
            self.send(d, clicked=False)
        self.assertIn("Cancelled", str(cm.exception))
        self.assertEqual((self.app.sent, self.app.saved), ([], []))
        self.assertEqual(snd.load(d["id"])["status"], "cancelled")

        d = self.draft("-Subject", "Q3 review", "-Start", "2026-09-18T14:00", "-Attendees", "Cassie Tsai", "-Optional", "pc.liao@contoso.com", "-Location", "Teams", "-Body", "議程")
        r, dlg = self.send(d, clicked=True)
        spec = dlg.call_args[0][0]
        self.assertEqual(dict(spec["rows"])["主旨 Subject"], "Q3 review")
        self.assertIn("2026-09-18 14:00 到 15:00（60 分鐘）", dict(spec["rows"])["時間 When"])
        self.assertEqual(spec["text"], d["full_body"])
        self.assertTrue(r["Created"] and r["Mode"] == "meeting")
        self.assertEqual(len(self.app.sent), 1)
        ap = self.app.sent[0]
        self.assertEqual((ap.Subject, ap.Location, ap.Body, ap.MeetingStatus), ("Q3 review", "Teams", d["full_body"], 1))
        self.assertEqual((ap.Start.isoformat(), ap.End.isoformat()), ("2026-09-18T14:00:00", "2026-09-18T15:00:00"))
        self.assertEqual([(x.Address, x.Type) for x in ap.Recipients], [("cassie.tsai@contoso.com", 1), ("pc.liao@contoso.com", 2)])
        self.assertEqual(snd.load(d["id"])["status"], "sent")
        with self.assertRaises(SystemExit):  # no second send
            self.send(d, clicked=True)
        self.assertEqual(len(self.app.sent), 1)

    def test_appointment_is_saved_not_sent(self):
        d = self.draft("-Subject", "Focus", "-Start", "2026-09-17T09:00", "-Duration", "30")
        r, dlg = self.send(d, clicked=True)
        self.assertEqual(dlg.call_args[0][0]["ok"], "建立 Create")
        self.assertEqual((r["Mode"], len(self.app.saved), len(self.app.sent)), ("appointment", 1, 0))
        self.assertEqual(self.app.saved[0].MeetingStatus, 0)

    def test_wrong_kind_and_stale_token(self):
        mail = snd.run_draft(snd.draft_parser().parse_args(["-To", "Cassie Tsai", "-Subject", "Hi", "-Body", "x"]))
        with mock.patch.object(snd, "confirm_dialog", return_value=True) as dlg:
            with self.assertRaises(SystemExit) as cm:
                mtg.run_send(mtg.send_parser().parse_args([mail["id"], "-Confirm", mail["confirm"]]))
            self.assertIn("mail draft", str(cm.exception))
            d = self.draft("-Subject", "Focus", "-Start", "2026-09-17T09:00")
            p = snd.drafts_dir() / f"{d['id']}.json"
            data = json.loads(p.read_text(encoding="utf-8"))
            data["start"] = "2026-09-17T03:00:00"
            p.write_text(json.dumps(data), encoding="utf-8")
            with self.assertRaises(SystemExit):
                mtg.run_send(mtg.send_parser().parse_args([d["id"], "-Confirm", d["confirm"]]))
            with self.assertRaises(SystemExit):
                snd.run_send(snd.send_parser().parse_args([d["id"], "-Confirm", d["confirm"]]))  # meeting draft through the mail sender
        self.assertEqual(dlg.call_count, 0)
        self.assertEqual((self.app.sent, self.app.saved), ([], []))

    def test_item_drift_is_not_sent(self):
        d = self.draft("-Subject", "Q3 review", "-Start", "2026-09-18T14:00", "-Attendees", "Cassie Tsai")
        real_add = fo.Recipients.Add
        with mock.patch.object(fo.Recipients, "Add", lambda self_, text: real_add(self_, "pc.liao@contoso.com")), self.assertRaises(SystemExit) as cm:
            self.send(d, clicked=True)
        self.assertIn("required attendees differ", str(cm.exception))
        self.assertEqual(self.app.sent, [])


if __name__ == "__main__":
    unittest.main()
