#!/usr/bin/env python3
"""READ-ONLY. Aggregated overview for building memory: top correspondents, folders, frequent topics,
likely newsletters, recurring meetings. No message bodies.

    python outlook_overview.py -Days 180 -OutFile overview.json
    python outlook_overview.py -Store 20230731
"""
import datetime as dt

import outlook_com as oc

_PATTERN = {0: "Daily", 1: "Weekly", 2: "Monthly", 3: "MonthNth", 5: "Yearly", 6: "YearNth"}


def _store_default(ns, store_obj, root, fid, names):
    if store_obj is not None:
        try:
            return store_obj.GetDefaultFolder(fid)
        except Exception:
            for f in root.Folders:
                if str(f.Name) in names:
                    return f
            return None
    return ns.GetDefaultFolder(fid)


def run(a, ns=None):
    ns = ns or oc.connect()
    since = dt.datetime.combine(dt.date.today(), dt.time()) - dt.timedelta(days=a.days)
    me = oc._safe(lambda: str(ns.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress), "") or oc._safe(lambda: str(ns.CurrentUser.Address), "")
    store_obj = oc.find_store(a.store, ns) if a.store else None
    root = store_obj.GetRootFolder() if store_obj else ns.GetDefaultFolder(oc.OL_FOLDER["Inbox"]).Parent

    folders = []
    for f in oc.mail_folders_recursive(root):
        newest = None
        try:
            it = f.Items
            it.Sort("[ReceivedTime]", True)
            first = it.GetFirst()
            if first is not None:
                newest = oc.iso(first.ReceivedTime)
        except Exception:
            pass
        folders.append({"Path": str(f.FolderPath), "Items": int(f.Items.Count), "Unread": oc._safe(lambda: int(f.UnReadItemCount)), "Newest": newest})

    senders, topics, scanned_in = {}, {}, 0
    inbox = _store_default(ns, store_obj, root, oc.OL_FOLDER["Inbox"], ("Inbox", "收件匣", "收件箱"))
    for f in (oc.mail_folders_recursive(inbox) if inbox is not None else []):
        items = f.Items
        items.Sort("[ReceivedTime]", True)
        for m in oc.iter_items(items):
            if scanned_in >= a.maxitems:
                break
            if int(oc._safe(lambda: m.Class, 0)) != oc.OL_MAIL_ITEM:
                continue
            rt = oc.to_datetime(m.ReceivedTime)
            if rt < since:
                break
            scanned_in += 1
            addr = oc.sender_smtp(m).lower()
            if addr:
                e = senders.setdefault(addr, {"Key": addr, "Count": 0, "Name": str(m.SenderName), "Last": oc.iso(rt), "Unsub": False, "Folder": str(f.FolderPath)})
                e["Count"] += 1
                if oc.iso(rt) > e["Last"]:
                    e["Last"] = oc.iso(rt)
                try:
                    if m.PropertyAccessor.GetProperty(oc.PR_LIST_UNSUBSCRIBE):
                        e["Unsub"] = True
                except Exception:
                    pass
            topic = str(oc._safe(lambda: m.ConversationTopic, "") or "")
            if topic:
                t = topics.setdefault(topic, {"Key": topic, "Count": 0, "Last": oc.iso(rt), "Sample": addr})
                t["Count"] += 1

    recips, scanned_out = {}, 0
    sent = _store_default(ns, store_obj, root, oc.OL_FOLDER["SentMail"], ("Sent Items", "寄件備份", "已发送邮件"))
    if sent is not None:
        items = sent.Items
        items.Sort("[SentOn]", True)
        for m in oc.iter_items(items):
            if scanned_out >= a.maxitems:
                break
            if int(oc._safe(lambda: m.Class, 0)) != oc.OL_MAIL_ITEM:
                continue
            st = oc.to_datetime(oc._safe(lambda: m.SentOn))
            if st and st < since:
                break
            scanned_out += 1
            for r in oc.recipient_list(m, 1):
                addr = r["Address"].lower()
                if addr:
                    e = recips.setdefault(addr, {"Key": addr, "Count": 0, "Name": r["Name"]})
                    e["Count"] += 1

    recurring = []
    cal = _store_default(ns, store_obj, root, oc.OL_FOLDER["Calendar"], ("Calendar", "行事曆", "日历"))
    if cal is not None:
        try:
            items = cal.Items
            items.IncludeRecurrences = False
            for it in oc.iter_items(items):
                if int(oc._safe(lambda: it.Class, 0)) != oc.OL_APPOINTMENT or not it.IsRecurring:
                    continue
                rp = it.GetRecurrencePattern()
                recurring.append({
                    "Subject": str(it.Subject), "Organizer": oc._safe(lambda: str(it.Organizer), ""),
                    "Pattern": _PATTERN.get(int(rp.RecurrenceType), "Other"), "Interval": int(rp.Interval),
                    "DayOfWeekMask": int(rp.DayOfWeekMask), "StartTime": oc.to_datetime(rp.StartTime).strftime("%H:%M"),
                    "DurationMinutes": int(rp.Duration), "PatternEnd": None if rp.NoEndDate else oc.to_datetime(rp.PatternEndDate).strftime("%Y-%m-%d"),
                    "Attendees": oc._safe(lambda: str(it.RequiredAttendees), ""),
                })
        except Exception:
            pass

    top = lambda d: sorted(d.values(), key=lambda x: -x["Count"])[:a.top]
    return {
        "GeneratedAt": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"), "Me": me,
        "Window": {"Since": since.strftime("%Y-%m-%dT%H:%M:%S"), "Days": a.days, "ScannedReceived": scanned_in, "ScannedSent": scanned_out, "MaxItems": a.maxitems},
        "Folders": folders, "TopSenders": top(senders), "TopRecipients": top(recips), "TopTopics": top(topics),
        "Newsletters": [s for s in sorted(senders.values(), key=lambda x: -x["Count"]) if s["Unsub"]][:a.top],
        "RecurringMeetings": recurring,
    }


def parser():
    ap = oc.ArgParser(description=__doc__)
    ap.opt("-Days", type=int, default=180)
    ap.opt("-MaxItems", type=int, default=3000, help="cap on mails scanned per direction, newest first")
    ap.opt("-Top", type=int, default=30)
    ap.opt("-Store", default="")
    oc.add_common_output(ap)
    return ap


def main(argv=None):
    a = parser().parse_args(argv)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
