#!/usr/bin/env python3
"""READ-ONLY. Aggregated overview for building memory: top correspondents, folders, frequent topics,
likely newsletters, recurring meetings. No message bodies.

    python outlook_overview.py -Days 180 -OutFile overview.json
    python outlook_overview.py -Store 20230731

Mail is read through Outlook's Table object (no item opened). Recipients are counted from the To
line of each sent mail; one sample mail per top recipient is opened to resolve the address.
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


def _split_names(display_to: str):
    return [n.strip() for n in (display_to or "").split(";") if n.strip()]


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
            for row in oc.scan_mail(f, limit=1):
                newest = row.summary["ReceivedTime"]
        except Exception:
            pass
        folders.append({"Path": str(f.FolderPath), "Items": int(f.Items.Count), "Unread": oc._safe(lambda: int(f.UnReadItemCount)), "Newest": newest})

    senders, topics, scanned_in = {}, {}, 0
    inbox = _store_default(ns, store_obj, root, oc.OL_FOLDER["Inbox"], ("Inbox", "收件匣", "收件箱"))
    for f in (oc.mail_folders_recursive(inbox) if inbox is not None else []):
        if scanned_in >= a.maxitems:
            break
        path = str(f.FolderPath)
        for row in oc.scan_mail(f, after=since, limit=a.maxitems - scanned_in, extra=("Unsubscribe",), folder_path=path):
            scanned_in += 1
            s = row.summary
            rt = s["ReceivedTime"]
            addr = oc.sender_address(row, ns)
            if addr:
                e = senders.setdefault(addr, {"Key": addr, "Count": 0, "Name": s["From"], "Last": rt, "Unsub": False, "Folder": path})
                e["Count"] += 1
                if rt and (e["Last"] is None or rt > e["Last"]):
                    e["Last"] = rt
                if row.extra.get("Unsubscribe"):
                    e["Unsub"] = True
            topic = s["ConversationTopic"]
            if topic:
                t = topics.setdefault(topic, {"Key": topic, "Count": 0, "Last": rt, "Sample": addr})
                t["Count"] += 1

    recips, samples, scanned_out = {}, {}, 0
    sent = _store_default(ns, store_obj, root, oc.OL_FOLDER["SentMail"], ("Sent Items", "寄件備份", "已发送邮件"))
    if sent is not None:
        for row in oc.scan_mail(sent, after=since, limit=a.maxitems, sort="SentOn"):
            scanned_out += 1
            for name in _split_names(row.summary["To"]):
                key = name.lower()
                e = recips.setdefault(key, {"Key": key, "Count": 0, "Name": name})
                e["Count"] += 1
                samples.setdefault(key, row)  # newest sent mail naming this recipient
        # addresses for the top recipients: one sample mail each, its To recipients read once
        for e in sorted(recips.values(), key=lambda x: -x["Count"])[:a.top]:
            for r in oc.recipient_list(oc.open_item(samples[e["Key"]], ns), 1):
                if r["Address"] and (r["Name"].lower() == e["Key"] or r["Address"].lower() == e["Key"]):
                    e["Key"] = r["Address"].lower()
                    break
        merged = {}
        for e in recips.values():  # two spellings of one address count once
            m = merged.setdefault(e["Key"], {"Key": e["Key"], "Count": 0, "Name": e["Name"]})
            m["Count"] += e["Count"]
        recips = merged

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
    oc.apply_settings(a)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
