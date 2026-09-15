import json, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
sys.path.insert(0, os.path.dirname(__file__))
import outlook_com as oc
import fake_outlook as fo
import outlook_search, outlook_thread, outlook_calendar, outlook_overview

ns = fo.build_fixture()
oc.set_namespace_for_tests(ns)

def search(*argv):
    return outlook_search.run(outlook_search.parser().parse_args(list(argv)))

# --- folder resolution: English and localized names inside a PST store
assert oc.get_folder("Inbox", "20230731").Name == "收件匣"
assert oc.get_folder("收件匣/人才", "20230731").Name == "人才"
assert oc.get_folder("Sent Items", "20230731").Name == "寄件備份"
assert oc.get_folder("\\\\20230731\\Inbox\\人才").Name == "人才"
assert oc.get_folder("").Name == "Inbox"

# --- search: default store only finds the Exchange mail; -AllStores finds the PST too
r = search("-From", "Cassie")
assert r["Count"] == 0, r
r = search("-From", "Cassie", "-AllStores")
assert [m["EntryID"] for m in r["Results"]] == ["id1", "id5"], r["Results"]
r = search("-From", "Cassie", "-Store", "20230731", "-After", "2026-01-01")
assert [m["EntryID"] for m in r["Results"]] == ["id1"]
assert r["Results"][0]["HasAttachments"] and r["Results"][0]["Attachments"][0]["FileName"].startswith("合約")
r = search("-AnyOf", "報價,quote", "-AllStores", "-AllFolders")
assert sorted(m["EntryID"] for m in r["Results"]) == ["id3", "id4"], r["Results"]
r = search("-Unread", "-AllStores")
assert [m["EntryID"] for m in r["Results"]] == ["id1"]
r = search("-Subject", "契約", "-AllStores")
assert r["Count"] == 0
r = search("-Store", "20230731", "-Before", "2026-09-10", "-After", "2026-09-01", "-IncludeBody")
assert [m["EntryID"] for m in r["Results"]] == ["id4"] and "Body" in r["Results"][0]

# --- thread by subject inside the PST, via GetConversation
t = outlook_thread.run(outlook_thread.parser().parse_args(["-Subject", "合約草稿", "-Store", "20230731"]))
assert t["Method"] == "GetConversation" and [m["EntryID"] for m in t["Messages"]] == ["id2", "id1"], t
assert t["Messages"][1]["ToRecipients"][0]["Address"] == "ben@contoso.com"
# fallback by topic when no conversation index
t = outlook_thread.run(outlook_thread.parser().parse_args(["-EntryID", "id3"]))
assert t["Method"] == "ConversationTopic" and [m["EntryID"] for m in t["Messages"]] == ["id3", "s1"], t

# --- calendar: recurrence-safe restrict, Free excluded, conflicts flagged
c = outlook_calendar.run(outlook_calendar.parser().parse_args(["-Start", "2026-09-16", "-Days", "1", "-Store", "20230731"]))
assert [i["Subject"] for i in c["Items"]] == ["每日站會", "供應商簡報", "1:1 with Bob"], c["Items"]
assert len(c["Conflicts"]) == 1 and c["Conflicts"][0]["B"] == "1:1 with Bob"
assert c["Items"][2]["BusyStatus"] == "Tentative" and c["Items"][0]["IsRecurring"]
c = outlook_calendar.run(outlook_calendar.parser().parse_args(["-Start", "2026-09-16", "-Store", "20230731", "-IncludeFree"]))
assert len(c["Items"]) == 4
c = outlook_calendar.run(outlook_calendar.parser().parse_args(["-Start", "2026-09-17", "-Store", "20230731"]))
assert c["Count"] == 0

# --- overview of the PST
o = outlook_overview.run(outlook_overview.parser().parse_args(["-Days", "3650", "-Store", "20230731"]))
assert o["TopSenders"][0]["Key"] == "cassie.tsai@contoso.com" and o["TopSenders"][0]["Count"] == 2, o["TopSenders"]
assert o["Newsletters"][0]["Key"] == "news@example.com"
assert {r["Key"] for r in o["TopRecipients"]} == {"david.chen@contoso.com", "pc.liao@contoso.com", "cassie.tsai@contoso.com"}
assert any(f["Path"].endswith("人才") for f in o["Folders"])
assert o["TopTopics"][0]["Key"] in ("合約草稿 v3 - 法務意見", "Q3 預算討論")

json.dumps(r, ensure_ascii=False); json.dumps(t); json.dumps(c); json.dumps(o)
print("all fake-COM tests passed")

# ================= new skills
import datetime as dt, tempfile, os
import outlook_followup, outlook_attachments, outlook_meeting_prep
from unittest import mock

class _Now(dt.datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 9, 14, 12, 0)

with mock.patch.object(outlook_followup.dt, "datetime", _Now):
    # sent: s2 (2026-09-05, PC never replied) waits; s1 got a reply? conv C2: id3 is from David but EARLIER; no later mail from others -> waits; s3 is <3 days old -> not yet
    f = outlook_followup.run(outlook_followup.parser().parse_args(["-Direction", "sent", "-Store", "20230731", "-Days", "3"]))
    ids = [r["EntryID"] for r in f["Results"]]
    assert ids == ["s2", "s1"], ids
    assert f["Results"][0]["WaitingDays"] == 9 and f["Results"][0]["Counterparts"][0]["Address"] == "pc.liao@contoso.com"
    # received: id6 (question from PC, no reply from me) waits; id1 (Cassie) was answered by s3; id7 FYI has no question; id4 newsletter is old but not a question
    f = outlook_followup.run(outlook_followup.parser().parse_args(["-Direction", "received", "-Store", "20230731", "-Days", "2", "-Lookback", "30"]))
    ids = [r["EntryID"] for r in f["Results"]]
    assert "id6" in ids and "id1" not in ids, ids
    q = next(r for r in f["Results"] if r["EntryID"] == "id6")
    assert q["LooksLikeQuestion"] and q["DirectToMe"] and q["WaitingDays"] == 4
    f = outlook_followup.run(outlook_followup.parser().parse_args(["-Direction", "received", "-Store", "20230731", "-QuestionsOnly", "-Lookback", "30"]))
    assert [r["EntryID"] for r in f["Results"]] == ["id6"], f["Results"]

# attachments: filter by ext / size / name, sort by size, save out
at = outlook_attachments.run(outlook_attachments.parser().parse_args(["-Store", "20230731", "-Sort", "size"]))
assert [r["FileName"] for r in at["Results"]] == ["Q3_report.pptx", "合約草稿_v3_legal.docx", "人才名單.xlsx", "image001.png"], at["Results"]
at = outlook_attachments.run(outlook_attachments.parser().parse_args(["-Store", "20230731", "-Ext", "xlsx,docx", "-MinSizeKB", "100"]))
assert [r["FileName"] for r in at["Results"]] == ["合約草稿_v3_legal.docx"]
at = outlook_attachments.run(outlook_attachments.parser().parse_args(["-Store", "20230731", "-Name", "名單"]))
assert at["Count"] == 1 and at["Results"][0]["From"] == "PC Liao"
d = tempfile.mkdtemp()
at = outlook_attachments.run(outlook_attachments.parser().parse_args(["-Store", "20230731", "-Ext", "pptx", "-SaveTo", d]))
assert os.path.exists(at["Results"][0]["SavedTo"]) and open(at["Results"][0]["SavedTo"], "rb").read() == b"pptx"

# search new flags
r = search("-Store", "20230731", "-HighImportance")
assert [m["EntryID"] for m in r["Results"]] == ["id1"]

# meeting prep: next meeting with attendees, mails from/to attendees, attachments
with mock.patch.object(outlook_meeting_prep.dt, "datetime", _Now):
    mp = outlook_meeting_prep.run(outlook_meeting_prep.parser().parse_args(["-Subject", "供應商", "-Store", "20230731", "-Horizon", "7", "-Days", "30"]))
    assert mp["Meeting"]["Subject"] == "供應商簡報"
    assert [x["Address"] for x in mp["Attendees"]] == ["cassie.tsai@contoso.com", "pc.liao@contoso.com"], mp["Attendees"]
    cas = mp["ByAttendee"][0]; pc = mp["ByAttendee"][1]
    assert "id1" in [m["EntryID"] for m in cas["Mails"]] and "s3" in [m["EntryID"] for m in cas["Mails"]], cas
    assert "id6" in [m["EntryID"] for m in pc["Mails"]] and "s2" in [m["EntryID"] for m in pc["Mails"]], pc
    assert any(a["FileName"] == "人才名單.xlsx" for a in mp["Attachments"])
print("new-skill tests passed")
