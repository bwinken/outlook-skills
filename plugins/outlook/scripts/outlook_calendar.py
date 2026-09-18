#!/usr/bin/env python3
"""READ-ONLY. Lists calendar items in a date range (recurrences expanded) and flags overlaps.

    python outlook_calendar.py                      # today
    python outlook_calendar.py -Days 7
    python outlook_calendar.py -Start 2026-09-15 -End 2026-09-20 -OutFile cal.json

Recurrence expansion needs the Items collection (a Table shows only the series master), so each
appointment in the range is read item by item; its Body only with -PreviewLength.
"""
import datetime as dt

import outlook_com as oc


def run(a, ns=None):
    ns = ns or oc.connect()
    start = oc.parse_date(a.start) if a.start else dt.datetime.combine(dt.date.today(), dt.time())
    end = oc.parse_date(a.end) if a.end else (start.replace(hour=0, minute=0, second=0) + dt.timedelta(days=a.days))
    cal = oc.find_store(a.store, ns).GetDefaultFolder(oc.OL_FOLDER["Calendar"]) if a.store else ns.GetDefaultFolder(oc.OL_FOLDER["Calendar"])

    # Canonical recurrence-safe pattern: sort by Start, IncludeRecurrences, Restrict with an overlap test.
    items = cal.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True
    filt = f"[Start] < '{oc.jet_date(end)}' AND [End] > '{oc.jet_date(start)}'"
    items = items.Restrict(filt)

    appts = []
    for it in oc.iter_items(items):
        if int(oc._safe(lambda: it.Class, 0)) != oc.OL_APPOINTMENT:
            continue
        s = oc.appointment_summary(it, a.previewlength)
        if a.includefree or s["BusyStatus"] != "Free":
            appts.append(s)
    appts.sort(key=lambda x: x["Start"] or "")

    conflicts = []
    for i in range(len(appts)):
        if appts[i]["AllDayEvent"]:
            continue
        for j in range(i + 1, len(appts)):
            if appts[j]["AllDayEvent"]:
                continue
            if appts[j]["Start"] < appts[i]["End"]:
                conflicts.append({"A": appts[i]["Subject"], "AStart": appts[i]["Start"], "AEnd": appts[i]["End"],
                                  "B": appts[j]["Subject"], "BStart": appts[j]["Start"], "BEnd": appts[j]["End"]})
            else:
                break
    return {
        "Range": {"Start": start.strftime("%Y-%m-%dT%H:%M:%S"), "End": end.strftime("%Y-%m-%dT%H:%M:%S"), "Filter": filt},
        "Calendar": str(cal.FolderPath), "Count": len(appts), "Conflicts": conflicts, "Items": appts,
    }


def parser():
    ap = oc.ArgParser(description=__doc__)
    ap.opt("-Start", default="", help="ISO date; default today")
    ap.opt("-End", default="", help="ISO date, exclusive")
    ap.opt("-Days", type=int, default=1, help="range length from -Start when -End is not given")
    ap.opt("-Store", default="")
    ap.flag("-IncludeFree", help="include items whose BusyStatus is Free")
    ap.opt("-PreviewLength", type=int, default=0, help="characters of BodyPreview per item (meeting notes, links); 0 = the body is not read")
    oc.add_common_output(ap)
    return ap


def main(argv=None):
    a = parser().parse_args(argv)
    oc.apply_settings(a)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
