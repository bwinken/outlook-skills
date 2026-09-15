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
assert {r["Key"] for r in o["TopRecipients"]} == {"david.chen@contoso.com", "pc.liao@contoso.com"}
assert any(f["Path"].endswith("人才") for f in o["Folders"])
assert o["TopTopics"][0]["Key"] in ("合約草稿 v3 - 法務意見", "Q3 預算討論")

json.dumps(r, ensure_ascii=False); json.dumps(t); json.dumps(c); json.dumps(o)
print("all fake-COM tests passed")
