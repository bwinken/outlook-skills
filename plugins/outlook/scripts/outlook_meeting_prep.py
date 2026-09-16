#!/usr/bin/env python3
"""READ-ONLY. Gather what you need before a meeting: the appointment, its attendees, recent mails
exchanged with them, mails about the meeting's subject, and attachments seen along the way.

    python outlook_meeting_prep.py -Next                       # the next upcoming meeting with attendees
    python outlook_meeting_prep.py -Subject "Q3 預算" -Days 30
    python outlook_meeting_prep.py -EntryID <appointment id> -AllStores
"""
import datetime as dt
import re

import outlook_com as oc
import outlook_calendar
import outlook_search

_STOP = {"meeting", "call", "sync", "review", "討論", "會議", "週會", "例會", "with", "and", "the", "re", "fw", "fwd"}


def _keywords(subject: str):
    words = [w for w in re.split(r"[\s\-_:/|,()（）【】\[\]]+", subject) if len(w) >= 2 and w.lower() not in _STOP]
    return words[:4]


def _find_appointment(a, ns):
    if a.entryid:
        it = ns.GetItemFromID(a.entryid)
        return oc.appointment_summary(it), it
    cal_args = outlook_calendar.parser().parse_args(["-Days", str(a.horizon)] + (["-Store", a.store] if a.store else []))
    cal = outlook_calendar.run(cal_args, ns)
    now = dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    cands = cal["Items"]
    if a.subject:
        cands = [c for c in cands if a.subject.lower() in c["Subject"].lower()]
    else:
        cands = [c for c in cands if c["End"] >= now and c["MeetingStatus"] != "NonMeeting" and c["MeetingStatus"] != "Canceled"]
    if not cands:
        raise SystemExit("No matching meeting in the next %d days." % a.horizon)
    chosen = cands[0]
    return chosen, ns.GetItemFromID(chosen["EntryID"])


def run(a, ns=None):
    ns = ns or oc.connect()
    me = oc.my_addresses(ns)
    appt, item = _find_appointment(a, ns)
    attendees = [x for x in oc.appointment_attendees(item) if x["Address"].lower() not in me]
    since = (dt.datetime.now() - dt.timedelta(days=a.days)).strftime("%Y-%m-%d")

    def search(extra):
        args = outlook_search.parser().parse_args(extra + ["-After", since, "-Max", str(a.max), "-PreviewLength", "300"] + (["-AllStores"] if a.allstores else []) + (["-Store", a.store] if a.store else []) + ["-AllFolders"])
        return outlook_search.run(args, ns)["Results"]

    seen, by_attendee = set(), []
    for att in attendees:
        key = att["Address"] or att["Name"]
        rec = search(["-From", key])
        sent = search(["-To", att["Name"] or att["Address"], "-Folder", "Sent Items"]) if (att["Name"] or att["Address"]) else []
        mails = []
        for m in rec + sent:
            if m["EntryID"] in seen:
                continue
            seen.add(m["EntryID"])
            mails.append(m)
        mails.sort(key=lambda x: x["ReceivedTime"] or "", reverse=True)
        by_attendee.append({"Attendee": att, "Count": len(mails), "Mails": mails[:a.max]})

    kws = _keywords(appt["Subject"])
    about = [m for m in (search(["-AnyOf", ",".join(kws)]) if kws else []) if m["EntryID"] not in seen]
    attachments = []
    for m in [m for grp in by_attendee for m in grp["Mails"]] + about:
        for at in m["Attachments"]:
            if int(at.get("Type", 1)) == 1:
                attachments.append({"FileName": at["FileName"], "SizeKB": round(at["Size"] / 1024, 1), "From": m["From"], "ReceivedTime": m["ReceivedTime"], "Subject": m["Subject"], "EntryID": m["EntryID"]})
    attachments.sort(key=lambda x: x["ReceivedTime"] or "", reverse=True)
    return {"Meeting": appt, "Attendees": attendees, "Since": since, "Keywords": kws,
            "ByAttendee": by_attendee, "AboutSubject": about[:a.max], "Attachments": attachments[:a.max]}


def parser():
    ap = oc.ArgParser(description=__doc__)
    ap.opt("-Subject", default="", help="meeting subject contains")
    ap.opt("-EntryID", default="", help="appointment EntryID from outlook_calendar.py")
    ap.flag("-Next", help="the next upcoming meeting (default when no selector is given)")
    ap.opt("-Horizon", type=int, default=7, help="days ahead to look for the meeting")
    ap.opt("-Days", type=int, default=30, help="how far back to look for mails")
    ap.opt("-Max", type=int, default=10, help="mails per attendee / per list")
    ap.opt("-Store", default="")
    ap.flag("-AllStores")
    oc.add_common_output(ap)
    return ap


def main(argv=None):
    a = parser().parse_args(argv)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
