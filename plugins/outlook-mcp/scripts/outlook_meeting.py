#!/usr/bin/env python3
"""Create a meeting (invitations sent to the attendees) or an appointment on the user's own calendar.
Like outlook_send.py: two steps, and the second only completes after the user clicks in a window on
their own desktop.

    python outlook_meeting.py draft --subject "Q3 review" --start 2026-09-18T14:00 [--end 2026-09-18T15:00 | --duration 60]
                                    [--attendees alice@contoso.com "PC Liao"] [--optional bob] [--location "Room 3 / Teams"]
                                    [--body-file agenda.txt | --body "..."] [--store X]
        Resolves attendees through Outlook, lists calendar items that overlap the slot (Conflicts), builds the
        exact invitation text (body + approval footer), stores the draft under ~/.outlook-skills/drafts/ and
        prints it. Nothing is created.

    python outlook_meeting.py send <id> --confirm <token>
        Window with subject, time, location, attendees and the text; after the click the appointment is
        created through COM with exactly the stored values, read back and compared, then sent to the
        attendees (Send) or, with no attendees, saved to the calendar (Save). No flag skips the window.

    python outlook_meeting.py show <id> | discard <id>      (list: outlook_send.py list)
"""
import datetime as dt

import outlook_com as oc
import outlook_calendar
import outlook_send as snd
import settings as ps

OL_ITEM_TYPE_APPOINTMENT = 1
OL_MEETING = 1          # MeetingStatus olMeeting
OL_REQUIRED, OL_OPTIONAL = 1, 2


def _fmt(d: dt.datetime) -> str:
    return d.strftime("%Y-%m-%dT%H:%M:%S")


def _conflicts(start, end, store, ns):
    args = ["-Start", start.strftime("%Y-%m-%d"), "-Days", str((end.date() - start.date()).days + 1)] + (["-Store", store] if store else [])
    try:
        items = outlook_calendar.run(outlook_calendar.parser().parse_args(args), ns)["Items"]
    except SystemExit:
        return []
    s, e = _fmt(start), _fmt(end)
    return [{"Subject": i["Subject"], "Start": i["Start"], "End": i["End"], "ResponseStatus": i["ResponseStatus"]}
            for i in items if not i["AllDayEvent"] and i["Start"] < e and i["End"] > s]


def run_draft(a, ns=None):
    ns = ns or oc.connect()
    merged = ps.resolve()[0]
    cfg, mcfg = (merged.get("send") or {}), (merged.get("meeting") or {})
    if not a.subject:
        raise SystemExit("--subject is required.")
    if not a.start:
        raise SystemExit("--start is required (YYYY-MM-DDTHH:MM).")
    start = oc.parse_date(a.start)
    if a.end:
        end = oc.parse_date(a.end)
    else:
        end = start + dt.timedelta(minutes=int(a.duration or mcfg.get("default_duration_minutes") or 60))
    if end <= start:
        raise SystemExit("--end must be after --start.")
    if a.body_file:
        with open(a.body_file, "r", encoding="utf-8-sig") as fh:
            body = fh.read()
    else:
        body = a.body or ""
    body = body.replace("\r\n", "\n").strip()
    me = oc.my_addresses(ns)
    required = [r for r in snd.resolve_recipients(a.attendees or [], ns) if not oc.is_me(r["Address"], me)]
    optional = [r for r in snd.resolve_recipients(a.optional or [], ns) if not oc.is_me(r["Address"], me) and r["Address"].lower() not in snd._addresses(required)]
    approver = snd.approver_name(ns, cfg)
    footer = snd.footer_text(cfg, approver)
    full_body = (body + "\n\n" if body else "") + footer
    d = {
        "id": snd.new_id(), "kind": "meeting",
        "status": "draft", "created": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "mode": "meeting" if (required or optional) else "appointment",
        "subject": a.subject.strip(), "start": _fmt(start), "end": _fmt(end),
        "duration_minutes": int((end - start).total_seconds() // 60),
        "location": (a.location or "").strip(),
        "required": required, "optional": optional,
        "body": body, "footer": footer, "full_body": full_body,
        "reminder_minutes": int(mcfg.get("reminder_minutes") if mcfg.get("reminder_minutes") is not None else 15),
        "store": a.store or "",
        "approver": approver,
        "conflicts": _conflicts(start, end, a.store, ns),
    }
    d["confirm"] = snd.token(d)
    snd.save(d)
    return d


def dialog_spec(d: dict) -> dict:
    when = f"{d['start'][:16].replace('T', ' ')} 到 {d['end'][11:16]}（{d['duration_minutes']} 分鐘）" if d["start"][:10] == d["end"][:10] else f"{d['start'][:16].replace('T', ' ')} 到 {d['end'][:16].replace('T', ' ')}"
    text = d["full_body"]
    if d["conflicts"]:
        text = "⚠ 撞期 / overlaps: " + "; ".join(f"{c['Subject']} {c['Start'][11:16]}-{c['End'][11:16]}" for c in d["conflicts"]) + "\n\n" + text
    is_meeting = d["mode"] == "meeting"
    return {"title": "Outlook 確認建立會議 / Confirm meeting" if is_meeting else "Outlook 確認建立行程 / Confirm appointment",
            "rows": [("主旨 Subject", d["subject"]), ("時間 When", when), ("地點 Location", d["location"] or "(none)"),
                     ("必要 Required", snd.fmt_people(d["required"])), ("選擇 Optional", snd.fmt_people(d["optional"]))],
            "text": text,
            "note": "邀請會以上面的內容寄給所有與會者。" if is_meeting else "行程會以上面的內容加進你的行事曆。",
            "ok": "送出邀請 Send" if is_meeting else "建立 Create",
            "question": "送出邀請？ Send the invitation?" if is_meeting else "建立行程？ Create the appointment?"}


def _verify(appt, d):
    got = {"Subject": str(appt.Subject), "Start": _fmt(oc.to_datetime(appt.Start)), "End": _fmt(oc.to_datetime(appt.End)),
           "Location": str(oc._safe(lambda: appt.Location, "") or ""), "Body": str(appt.Body).replace("\r\n", "\n").rstrip()}
    want = {"Subject": d["subject"], "Start": d["start"], "End": d["end"], "Location": d["location"], "Body": d["full_body"].rstrip()}
    problems = [f"{k} differs" for k in want if got[k] != want[k]]
    for rtype, key in ((OL_REQUIRED, "required"), (OL_OPTIONAL, "optional")):
        have = sorted(snd._addr_of(r).lower() or str(r.Address).lower() for r in appt.Recipients if int(r.Type) == rtype)
        if have != snd._addresses(d[key]):
            problems.append(f"{key} attendees differ: item {have} vs draft {snd._addresses(d[key])}")
    if problems:
        raise SystemExit("Item does not match the approved draft, nothing was created: " + "; ".join(problems))


def run_send(a, ns=None):
    d = snd.load_sendable(a.id, "meeting", a.confirm)
    cfg = (ps.resolve()[0].get("send") or {})
    timeout = int(cfg.get("dialog_timeout_seconds") or 300)

    # the user's click, on their own screen; there is no way around this call
    if not snd.confirm_dialog(dialog_spec(d), timeout):
        snd.cancel(d, "created")

    ns = ns or oc.connect()
    app = oc.application()
    appt = app.CreateItem(OL_ITEM_TYPE_APPOINTMENT)
    appt.Subject = d["subject"]
    appt.Start = dt.datetime.strptime(d["start"], "%Y-%m-%dT%H:%M:%S")
    appt.End = dt.datetime.strptime(d["end"], "%Y-%m-%dT%H:%M:%S")
    appt.Location = d["location"]
    appt.Body = d["full_body"]
    try:
        appt.ReminderSet = d["reminder_minutes"] > 0
        appt.ReminderMinutesBeforeStart = d["reminder_minutes"]
    except Exception:
        pass
    if d["mode"] == "meeting":
        appt.MeetingStatus = OL_MEETING
        recips = appt.Recipients
        for rtype, rows in ((OL_REQUIRED, d["required"]), (OL_OPTIONAL, d["optional"])):
            for r in rows:
                rec = recips.Add(r["Address"])
                rec.Type = rtype
        if not recips.ResolveAll():
            raise SystemExit("Outlook could not resolve every attendee; nothing was created.")
    _verify(appt, d)
    if d["mode"] == "meeting":
        appt.Send()
    else:
        appt.Save()
    d["status"], d["sent_at"] = "sent", dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    d["entry_id"] = str(oc._safe(lambda: appt.EntryID, "") or "")
    snd.save(d)
    return {"Created": True, "Mode": d["mode"], "Id": d["id"], "EntryID": d["entry_id"], "Subject": d["subject"], "Start": d["start"], "End": d["end"],
            "Location": d["location"], "Required": d["required"], "Optional": d["optional"], "SentAt": d["sent_at"], "Approver": d["approver"]}


def draft_parser():
    ap = oc.ArgParser(description="Build and store a meeting draft; creates nothing. Prints the draft with its id, confirm token and overlapping calendar items.")
    ap.opt("-Subject", default="", help="required")
    ap.opt("-Start", default="", help="required, local time, YYYY-MM-DDTHH:MM")
    ap.opt("-End", default="", help="YYYY-MM-DDTHH:MM; default start + duration")
    ap.opt("-Duration", type=int, default=0, help="minutes, when -End is not given (default: settings meeting.default_duration_minutes, 60)")
    ap.opt("-Attendees", nargs="+", default=[], help="required attendees: addresses or names Outlook can resolve. None = appointment on the user's own calendar")
    ap.opt("-Optional", nargs="+", default=[], help="optional attendees")
    ap.opt("-Location", default="", help="room, link or place; free text")
    ap.opt("-BodyFile", dest="body_file", default="", help="UTF-8 text file with the agenda / description")
    ap.opt("-Body", default="", help="agenda inline (short)")
    ap.opt("-Store", default="", help="calendar store for the conflict check")
    oc.add_common_output(ap)
    return ap


def send_parser():
    ap = oc.ArgParser(description="Show the confirmation window for a stored meeting draft and create it after the user's click.")
    ap.add_argument("id", help="draft id from `draft`")
    ap.opt("-Confirm", default="", required=True, help="confirm token printed by `draft`")
    oc.add_common_output(ap)
    return ap


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("draft", parents=[draft_parser()], add_help=False).set_defaults(fn=lambda a: run_draft(a))
    sub.add_parser("send", parents=[send_parser()], add_help=False).set_defaults(fn=lambda a: run_send(a))
    p = sub.add_parser("show"); p.add_argument("id"); p.set_defaults(fn=snd.run_show, out_file="")
    p = sub.add_parser("discard"); p.add_argument("id"); p.set_defaults(fn=snd.run_discard, out_file="")
    a = ap.parse_args(argv)
    oc.write_json(a.fn(a), getattr(a, "out_file", ""))


if __name__ == "__main__":
    main()
