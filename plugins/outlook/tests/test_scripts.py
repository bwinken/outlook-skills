"""Script tests against the fake Outlook object model (no Windows needed).

    python -m unittest discover -s plugins/outlook/tests -v
"""
import datetime as dt
import json
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
sys.path.insert(0, os.path.dirname(__file__))
import outlook_com as oc  # noqa: E402
import fake_outlook as fo  # noqa: E402
import outlook_search, outlook_thread, outlook_calendar, outlook_overview  # noqa: E402
import outlook_followup, outlook_attachments, outlook_meeting_prep, outlook_style  # noqa: E402


class _Now(dt.datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 9, 14, 12, 0)


def search(*argv):
    return outlook_search.run(outlook_search.parser().parse_args(list(argv)))


def ids(rows):
    return [m["EntryID"] for m in rows]


class FakeOutlookTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        oc.set_namespace_for_tests(fo.build_fixture())


class FolderTest(FakeOutlookTest):
    def test_english_and_localized_names_inside_a_pst(self):
        self.assertEqual(oc.get_folder("Inbox", "20230731").Name, "收件匣")
        self.assertEqual(oc.get_folder("收件匣/人才", "20230731").Name, "人才")
        self.assertEqual(oc.get_folder("Sent Items", "20230731").Name, "寄件備份")
        self.assertEqual(oc.get_folder("\\\\20230731\\Inbox\\人才").Name, "人才")
        self.assertEqual(oc.get_folder("").Name, "Inbox")


class SearchTest(FakeOutlookTest):
    def test_default_store_vs_allstores(self):
        self.assertEqual(search("-From", "Cassie")["Count"], 0)
        self.assertEqual(ids(search("-From", "Cassie", "-AllStores")["Results"]), ["id1", "id5"])

    def test_filters(self):
        r = search("-From", "Cassie", "-Store", "20230731", "-After", "2026-01-01")
        self.assertEqual(ids(r["Results"]), ["id1"])
        self.assertTrue(r["Results"][0]["HasAttachments"])
        self.assertTrue(r["Results"][0]["Attachments"][0]["FileName"].startswith("合約"))
        self.assertEqual(sorted(ids(search("-AnyOf", "報價,quote", "-AllStores", "-AllFolders")["Results"])), ["id3", "id4"])
        self.assertEqual(ids(search("-Unread", "-AllStores")["Results"]), ["id1"])
        self.assertEqual(search("-Subject", "契約", "-AllStores")["Count"], 0)
        self.assertEqual(ids(search("-Store", "20230731", "-HighImportance")["Results"]), ["id1"])

    def test_date_range_goes_into_the_dasl_filter(self):
        r = search("-Store", "20230731", "-Before", "2026-09-10", "-After", "2026-09-01", "-IncludeBody")
        self.assertEqual(ids(r["Results"]), ["id4"])
        self.assertIn("Body", r["Results"][0])
        self.assertIn("datereceived\" >= '2026-09-01 00:00'", r["Query"]["Dasl"])
        self.assertIn("datereceived\" < '2026-09-10 00:00'", r["Query"]["Dasl"])
        json.dumps(r, ensure_ascii=False)

    def test_preview_length_zero_skips_the_body(self):
        r = search("-Store", "20230731", "-PreviewLength", "0", "-Max", "1")
        self.assertEqual(r["Results"][0]["BodyPreview"], "")
        self.assertNotIn("Body", r["Results"][0])


class ThreadTest(FakeOutlookTest):
    def test_by_subject_via_get_conversation(self):
        t = outlook_thread.run(outlook_thread.parser().parse_args(["-Subject", "合約草稿", "-Store", "20230731"]))
        self.assertEqual(t["Method"], "GetConversation")
        self.assertEqual(ids(t["Messages"]), ["id2", "id1"])
        self.assertEqual(t["Messages"][1]["ToRecipients"][0]["Address"], "ben@contoso.com")
        json.dumps(t)

    def test_fallback_by_topic(self):
        t = outlook_thread.run(outlook_thread.parser().parse_args(["-EntryID", "id3"]))
        self.assertEqual(t["Method"], "ConversationTopic")
        self.assertEqual(ids(t["Messages"]), ["id3", "s1"])


class CalendarTest(FakeOutlookTest):
    def test_range_free_and_conflicts(self):
        c = outlook_calendar.run(outlook_calendar.parser().parse_args(["-Start", "2026-09-16", "-Days", "1", "-Store", "20230731"]))
        self.assertEqual([i["Subject"] for i in c["Items"]], ["每日站會", "供應商簡報", "1:1 with Bob"])
        self.assertEqual(len(c["Conflicts"]), 1)
        self.assertEqual(c["Conflicts"][0]["B"], "1:1 with Bob")
        self.assertEqual(c["Items"][2]["BusyStatus"], "Tentative")
        self.assertTrue(c["Items"][0]["IsRecurring"])
        json.dumps(c)
        c = outlook_calendar.run(outlook_calendar.parser().parse_args(["-Start", "2026-09-16", "-Store", "20230731", "-IncludeFree"]))
        self.assertEqual(len(c["Items"]), 4)
        c = outlook_calendar.run(outlook_calendar.parser().parse_args(["-Start", "2026-09-17", "-Store", "20230731"]))
        self.assertEqual(c["Count"], 0)


class OverviewTest(FakeOutlookTest):
    def test_pst_overview(self):
        o = outlook_overview.run(outlook_overview.parser().parse_args(["-Days", "3650", "-Store", "20230731"]))
        self.assertEqual(o["TopSenders"][0]["Key"], "cassie.tsai@contoso.com")
        self.assertEqual(o["TopSenders"][0]["Count"], 2)
        self.assertEqual(o["Newsletters"][0]["Key"], "news@example.com")
        self.assertEqual({r["Key"] for r in o["TopRecipients"]}, {"david.chen@contoso.com", "pc.liao@contoso.com", "cassie.tsai@contoso.com"})
        self.assertTrue(any(f["Path"].endswith("人才") for f in o["Folders"]))
        self.assertIn(o["TopTopics"][0]["Key"], ("合約草稿 v3 - 法務意見", "Q3 預算討論"))
        json.dumps(o)


class FollowupTest(FakeOutlookTest):
    def run_followup(self, *argv):
        with mock.patch.object(outlook_followup.dt, "datetime", _Now):
            return outlook_followup.run(outlook_followup.parser().parse_args(list(argv)))

    def test_sent(self):
        # s2 (2026-09-05, PC never replied) waits; s1: id3 from David is EARLIER, no later mail from others -> waits; s3 is <3 days old
        f = self.run_followup("-Direction", "sent", "-Store", "20230731", "-Days", "3")
        self.assertEqual(ids(f["Results"]), ["s2", "s1"])
        self.assertEqual(f["Results"][0]["WaitingDays"], 9)
        self.assertEqual(f["Results"][0]["Counterparts"][0]["Address"], "pc.liao@contoso.com")

    def test_received(self):
        # id6 (question from PC, no reply from me) waits; id1 (Cassie) was answered by s3; id7 FYI is not a question
        f = self.run_followup("-Direction", "received", "-Store", "20230731", "-Days", "2", "-Lookback", "30")
        self.assertIn("id6", ids(f["Results"]))
        self.assertNotIn("id1", ids(f["Results"]))
        q = next(r for r in f["Results"] if r["EntryID"] == "id6")
        self.assertTrue(q["LooksLikeQuestion"] and q["DirectToMe"])
        self.assertEqual(q["WaitingDays"], 4)
        f = self.run_followup("-Direction", "received", "-Store", "20230731", "-QuestionsOnly", "-Lookback", "30")
        self.assertEqual(ids(f["Results"]), ["id6"])

    def test_both_in_one_scan(self):
        f = self.run_followup("-Direction", "both", "-Store", "20230731", "-Lookback", "30")
        self.assertEqual(f["Direction"], "both")
        self.assertEqual((f["Sent"]["Days"], f["Received"]["Days"]), (3, 2))
        self.assertEqual(ids(f["Sent"]["Results"]), ["s2", "s1"])
        self.assertIn("id6", ids(f["Received"]["Results"]))
        self.assertNotIn("Results", f)


class AttachmentsTest(FakeOutlookTest):
    def test_filter_sort_save(self):
        at = outlook_attachments.run(outlook_attachments.parser().parse_args(["-Store", "20230731", "-Sort", "size"]))
        self.assertEqual([r["FileName"] for r in at["Results"]], ["Q3_report.pptx", "合約草稿_v3_legal.docx", "人才名單.xlsx", "image001.png"])
        at = outlook_attachments.run(outlook_attachments.parser().parse_args(["-Store", "20230731", "-Ext", "xlsx,docx", "-MinSizeKB", "100"]))
        self.assertEqual([r["FileName"] for r in at["Results"]], ["合約草稿_v3_legal.docx"])
        at = outlook_attachments.run(outlook_attachments.parser().parse_args(["-Store", "20230731", "-Name", "名單"]))
        self.assertEqual((at["Count"], at["Results"][0]["From"]), (1, "PC Liao"))
        d = tempfile.mkdtemp()
        at = outlook_attachments.run(outlook_attachments.parser().parse_args(["-Store", "20230731", "-Ext", "pptx", "-SaveTo", d]))
        with open(at["Results"][0]["SavedTo"], "rb") as fh:
            self.assertEqual(fh.read(), b"pptx")


class MeetingPrepTest(FakeOutlookTest):
    def test_next_meeting_with_attendees(self):
        with mock.patch.object(outlook_meeting_prep.dt, "datetime", _Now):
            mp = outlook_meeting_prep.run(outlook_meeting_prep.parser().parse_args(["-Subject", "供應商", "-Store", "20230731", "-Horizon", "7", "-Days", "30"]))
        self.assertEqual(mp["Meeting"]["Subject"], "供應商簡報")
        self.assertEqual([x["Address"] for x in mp["Attendees"]], ["cassie.tsai@contoso.com", "pc.liao@contoso.com"])
        cas, pc = mp["ByAttendee"]
        self.assertTrue({"id1", "s3"} <= set(ids(cas["Mails"])))
        self.assertTrue({"id6", "s2"} <= set(ids(pc["Mails"])))
        self.assertTrue(any(a["FileName"] == "人才名單.xlsx" for a in mp["Attachments"]))


class StyleTest(FakeOutlookTest):
    def test_profile(self):
        with mock.patch.object(outlook_style.dt, "datetime", _Now):
            st = outlook_style.run(outlook_style.parser().parse_args(["-Store", "20230731", "-Days", "365"]))
        self.assertEqual(st["Window"]["SentAnalysed"], 3)
        self.assertGreaterEqual(st["Window"]["ReceivedAnalysed"], 5)
        by = {e["Address"]: e for e in st["ReplyRate"]["BySender"]}
        self.assertEqual((by["cassie.tsai@contoso.com"]["Replied"], by["cassie.tsai@contoso.com"]["Rate"]), (1, 1.0))
        self.assertEqual(by["pc.liao@contoso.com"]["Replied"], 0)
        self.assertNotIn("news@example.com", by)
        self.assertEqual(st["ReplyRate"]["Newsletters"], ["news@example.com"])
        self.assertGreaterEqual(st["ReplyLatencyHours"]["Samples"], 1)
        self.assertGreater(st["Length"]["MedianChars"], 0)
        self.assertTrue(st["Language"]["MostlyChinese"])
        json.dumps(st, ensure_ascii=False)


if __name__ == "__main__":
    unittest.main()
