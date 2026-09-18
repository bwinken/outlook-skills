"""Script tests against the fake Outlook object model (no Windows needed).

    python -m unittest discover -s plugins/outlook/tests -v
"""
import datetime as dt
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
import outlook_search, outlook_thread, outlook_calendar, outlook_overview  # noqa: E402
import outlook_followup, outlook_attachments, outlook_style  # noqa: E402


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


class EntryIdTest(FakeOutlookTest):
    def test_one_mail_by_entryid_skips_folders_and_filters(self):
        r = search("-EntryID", "id7", "-From", "nobody", "-HasAttachments")  # the other filters are ignored
        self.assertEqual((r["Count"], r["Query"]["EntryID"], r["Query"]["Dasl"], r["Query"]["Folders"]), (1, "id7", "", ["\\\\20230731\\收件匣"]))
        self.assertEqual([a["FileName"] for a in r["Results"][0]["Attachments"]], ["Q3_report.pptx", "image001.png"])
        with self.assertRaises(SystemExit):
            search("-EntryID", "no-such-id")


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
        # one mail by EntryID: no search filter, so a mail Outlook would not flag as having attachments is reached too
        at = outlook_attachments.run(outlook_attachments.parser().parse_args(["-EntryID", "id7", "-Ext", "png", "-SaveTo", d]))
        self.assertEqual((at["MailsScanned"], [r["FileName"] for r in at["Results"]]), (1, ["image001.png"]))
        with open(at["Results"][0]["SavedTo"], "rb") as fh:
            self.assertEqual(fh.read(), b"png")


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


class TableScanTest(FakeOutlookTest):
    """Folder scans go through Folder.GetTable; an item is opened only for its body or its attachments."""

    def opened(self):
        ns = oc._namespace
        real, log = ns.GetItemFromID, []
        patch = mock.patch.object(ns, "GetItemFromID", side_effect=lambda eid: (log.append(eid), real(eid))[1])
        return patch, log

    def test_default_search_opens_no_item(self):
        patch, log = self.opened()
        with patch:
            r = search("-Subject", "Q3", "-Store", "20230731", "-AllFolders")  # id3 has no attachments
        self.assertEqual((ids(r["Results"]), log), (["id3"], []))
        m = r["Results"][0]
        self.assertEqual((m["FromAddress"], m["Folder"], m["BodyPreview"], m["Unread"], m["Size"]), ("david.chen@contoso.com", "\\\\20230731\\收件匣\\人才", "", False, 1024))
        self.assertNotIn("Body", m)

    def test_only_previews_and_attachments_open_items(self):
        patch, log = self.opened()
        with patch:
            r = search("-Store", "20230731")
        self.assertEqual(sorted(log), ["id1", "id6", "id7"])  # the mails with attachments, for their list
        by = {m["EntryID"]: m for m in r["Results"]}
        self.assertEqual([a["FileName"] for a in by["id7"]["Attachments"]], ["Q3_report.pptx", "image001.png"])
        self.assertEqual((by["id2"]["Attachments"], by["id2"]["HasAttachments"]), ([], False))
        log.clear()
        with patch:
            r = search("-Store", "20230731", "-PreviewLength", "12")
        self.assertEqual(sorted(log), sorted(ids(r["Results"])))  # every result needed its body once
        self.assertEqual(r["Results"][0]["BodyPreview"], "Hi, 法務回來了 hx...")
        with patch:
            r = search("-Store", "20230731", "-IncludeBody", "-Max", "1")
        self.assertEqual(r["Results"][0]["Body"], "Hi, 法務回來了 hxxp")

    def test_non_mail_rows_are_skipped_without_opening(self):
        patch, log = self.opened()
        with patch:
            r = search("-From", "Cassie", "-Store", "20230731", "-Unread")  # mr1 is an unread meeting request from Cassie
        self.assertEqual((ids(r["Results"]), log), (["id1"], ["id1"]))
        with patch:
            t = outlook_thread.run(outlook_thread.parser().parse_args(["-Subject", "合約草稿", "-Store", "20230731"]))
        self.assertEqual(ids(t["Messages"]), ["id2", "id1"])
        self.assertNotIn("mr1", log)

    def test_second_candidate_column_when_the_store_rejects_the_first(self):
        r = search("-Store", "20230731", "-HasAttachments")
        self.assertEqual(sorted(ids(r["Results"])), ["id1", "id6", "id7"])
        self.assertTrue(all(m["HasAttachments"] for m in r["Results"]))
        self.assertIn(oc.PR_HASATTACH, oc._unsupported_columns)

    def test_row_helpers(self):
        self.assertEqual((oc._hex(b"\x00\xab"), oc._hex("ABCD"), oc._hex((0, 171)), oc._hex(None)), ("00AB", "ABCD", "00AB", ""))
        self.assertEqual(oc._as_rows(((1, 2, 3), (4, 5, 6)), 3), [(1, 2, 3), (4, 5, 6)])
        self.assertEqual(oc._as_rows(((1, 4), (2, 5), (3, 6)), 3), [(1, 2, 3), (4, 5, 6)])  # transposed array
        self.assertEqual((oc._as_rows((), 3), oc._as_rows(None, 3)), ([], []))
        self.assertEqual(oc._text(("a", "b")), "a, b")
        self.assertIsNone(oc._row_from_values(("x", "s", None, None, "IPM.Schedule.Meeting.Request"), {"EntryID": 0, "Subject": 1, "MessageClass": 4}, "f", ()))

    def test_rows_one_by_one_when_getarray_fails(self):
        def broken(self, n):
            raise Exception("GetArray not supported")
        with mock.patch.object(fo.FolderTable, "GetArray", broken):
            r = search("-Store", "20230731", "-AllFolders")
        self.assertEqual(len(r["Results"]), 7)
        self.assertEqual(r["Results"][0]["EntryID"], "id1")

    def test_namespace_date_columns_come_in_utc_and_are_converted(self):
        import time
        tz = os.environ.get("TZ")
        os.environ["TZ"] = "Asia/Taipei"
        if hasattr(time, "tzset"):
            time.tzset()

        def restore():
            if tz is None:
                os.environ.pop("TZ", None)
            else:
                os.environ["TZ"] = tz
            if hasattr(time, "tzset"):
                time.tzset()
        self.addCleanup(restore)
        self.addCleanup(oc._unsupported_columns.clear)
        with mock.patch.object(fo, "_REJECTED_COLUMNS", fo._REJECTED_COLUMNS | {"ReceivedTime", "SentOn"}):  # a store without the built-in date columns
            oc._unsupported_columns.clear()
            r = search("-Store", "20230731", "-After", "2026-09-12", "-Before", "2026-09-13")
        self.assertEqual(ids(r["Results"]), ["id1"])
        self.assertEqual((r["Results"][0]["ReceivedTime"], r["Results"][0]["SentOn"]), ("2026-09-12T16:42:00", "2026-09-12T16:42:00"))
        self.assertEqual(oc._utc_to_local(dt.datetime(2026, 9, 12, 8, 42)), dt.datetime(2026, 9, 12, 16, 42) if hasattr(time, "tzset") else oc._utc_to_local(dt.datetime(2026, 9, 12, 8, 42)))

    def test_to_me_sources(self):
        ns = oc._namespace
        rows = list(oc.scan_mail(oc.get_folder("Inbox", "20230731", ns), extra=("ToMe",)))
        by = {row.summary["EntryID"]: row for row in rows}
        me = {"me@contoso.com"}
        self.assertTrue(oc.to_me(by["id6"], me))            # delivery flag column
        by["id6"].extra["ToMe"] = None
        self.assertFalse(oc.to_me(by["id6"], me))           # no flag, To line empty, item stays closed
        self.assertTrue(oc.to_me(by["id6"], me, ns, open_if_needed=True))
        self.assertIsNotNone(by["id6"].item)


class ItemsFallbackTest(FakeOutlookTest):
    """A store without Table support gives the same JSON through the Items collection."""

    def both_ways(self, fn):
        table = fn()
        fo.TABLES_SUPPORTED = False
        try:
            items = fn()
        finally:
            fo.TABLES_SUPPORTED = True
        self.assertEqual(json.dumps(items, ensure_ascii=False, sort_keys=True), json.dumps(table, ensure_ascii=False, sort_keys=True))
        return table

    def test_search(self):
        self.both_ways(lambda: search("-Store", "20230731", "-AllFolders", "-PreviewLength", "30"))
        self.both_ways(lambda: search("-AnyOf", "報價,quote", "-AllStores", "-AllFolders", "-After", "2026-01-01"))
        r = self.both_ways(lambda: search("-Store", "20230731", "-Unread"))
        self.assertEqual(ids(r["Results"]), ["id1"])

    def test_followup_overview_style_thread(self):
        with mock.patch.object(outlook_followup.dt, "datetime", _Now):
            f = self.both_ways(lambda: outlook_followup.run(outlook_followup.parser().parse_args(["-Direction", "both", "-Store", "20230731", "-Lookback", "30"])))
        self.assertEqual(ids(f["Sent"]["Results"]), ["s2", "s1"])
        q = next(r for r in f["Received"]["Results"] if r["EntryID"] == "id6")
        self.assertTrue(q["DirectToMe"] and q["LooksLikeQuestion"])
        o = self.both_ways(lambda: outlook_overview.run(outlook_overview.parser().parse_args(["-Days", "3650", "-Store", "20230731"])))
        self.assertEqual({r["Key"] for r in o["TopRecipients"]}, {"david.chen@contoso.com", "pc.liao@contoso.com", "cassie.tsai@contoso.com"})
        with mock.patch.object(outlook_style.dt, "datetime", _Now):
            st = self.both_ways(lambda: outlook_style.run(outlook_style.parser().parse_args(["-Store", "20230731", "-Days", "365"])))
        self.assertEqual(st["Window"]["BodySamples"], 3)
        self.both_ways(lambda: outlook_thread.run(outlook_thread.parser().parse_args(["-EntryID", "id3"])))


class SettingsDefaultsTest(unittest.TestCase):
    """The command line takes store / default_folder / all_folders from settings.json when not given."""

    def setUp(self):
        home = str(Path(tempfile.mkdtemp()).resolve())
        self.env = mock.patch.dict(os.environ, {"HOME": home, "USERPROFILE": home})
        self.env.start()
        self.addCleanup(self.env.stop)
        os.makedirs(os.path.join(home, ".outlook-skills"))
        with open(os.path.join(home, ".outlook-skills", "settings.json"), "w", encoding="utf-8") as fh:
            json.dump({"store": "20230731", "search": {"default_folder": "Inbox/人才", "all_folders": True}}, fh)

    def test_applied_to_search_and_thread(self):
        a = outlook_search.parser().parse_args([])
        self.assertEqual(oc.apply_settings(a), ["store", "search.default_folder", "search.all_folders"])
        self.assertEqual((a.store, a.folder, a.allfolders), ("20230731", "Inbox/人才", True))
        a = outlook_search.parser().parse_args(["-AllStores", "-Folder", "Sent Items"])
        self.assertEqual((oc.apply_settings(a), a.store, a.folder, a.allfolders), (["search.all_folders"], "", "Sent Items", True))
        a = outlook_search.parser().parse_args(["-EntryID", "id7"])
        self.assertEqual((oc.apply_settings(a), a.folder, a.allfolders), (["store"], "", False))
        a = outlook_thread.parser().parse_args(["-Subject", "x"])
        self.assertEqual((oc.apply_settings(a), a.store), (["store"], "20230731"))
        a = outlook_calendar.parser().parse_args([])
        self.assertEqual((oc.apply_settings(a), a.store, a.previewlength), (["store"], "20230731", 0))


if __name__ == "__main__":
    unittest.main()
