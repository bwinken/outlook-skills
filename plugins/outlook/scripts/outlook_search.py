#!/usr/bin/env python3
"""READ-ONLY. Searches mail by sender, recipient, subject, body, date range, attachments, unread state.

    python outlook_search.py -From alice -After 2026-09-01 -Max 20
    python outlook_search.py -Subject invoice -HasAttachments -AllFolders
    python outlook_search.py -AnyOf 報價,quote,pricing -After 2026-06-01 -AllStores -Max 300 -PreviewLength 500 -OutFile candidates.json
"""
import outlook_com as oc


def build_dasl(a, after=None, before=None) -> str:
    c = []
    # date bounds go into the filter so Outlook applies them with its index instead of the loop below
    # walking every newer item; the loop checks the dates too, and run() retries without them if a
    # store rejects the literal
    if after:
        c.append(f'"urn:schemas:httpmail:datereceived" >= \'{oc.dasl_date(after)}\'')
    if before:
        c.append(f'"urn:schemas:httpmail:datereceived" < \'{oc.dasl_date(before)}\'')
    if a.from_:
        v = oc.dasl_literal(a.from_)
        c.append(f'("urn:schemas:httpmail:fromname" LIKE \'%{v}%\' OR "urn:schemas:httpmail:fromemail" LIKE \'%{v}%\')')
    if a.to:
        v = oc.dasl_literal(a.to)
        c.append(f'("urn:schemas:httpmail:displayto" LIKE \'%{v}%\' OR "urn:schemas:httpmail:displaycc" LIKE \'%{v}%\')')
    if a.subject:
        c.append(f'"urn:schemas:httpmail:subject" LIKE \'%{oc.dasl_literal(a.subject)}%\'')
    if a.body:
        c.append(f'"urn:schemas:httpmail:textdescription" LIKE \'%{oc.dasl_literal(a.body)}%\'')
    if a.text:
        v = oc.dasl_literal(a.text)
        c.append(f'("urn:schemas:httpmail:subject" LIKE \'%{v}%\' OR "urn:schemas:httpmail:textdescription" LIKE \'%{v}%\')')
    terms = [t.strip() for chunk in (a.anyof or []) for t in chunk.split(",") if t.strip()]
    if terms:
        ors = [f'"urn:schemas:httpmail:subject" LIKE \'%{oc.dasl_literal(t)}%\' OR "urn:schemas:httpmail:textdescription" LIKE \'%{oc.dasl_literal(t)}%\'' for t in terms]
        c.append("(" + " OR ".join(ors) + ")")
    if a.hasattachments:
        c.append('"urn:schemas:httpmail:hasattachment" = 1')
    if a.unread:
        c.append('"urn:schemas:httpmail:read" = 0')
    if getattr(a, "highimportance", False):
        c.append('"urn:schemas:httpmail:importance" = 2')
    if getattr(a, "flagged", False):
        c.append('"http://schemas.microsoft.com/mapi/proptag/0x10900003" = 2')
    return ("@SQL=" + " AND ".join(c)) if c else ""


def resolve_folders(a, ns):
    folders = []
    if a.allstores:
        names = [str(s.DisplayName) for s in oc.get_stores(ns) if int(oc._safe(lambda: s.ExchangeStoreType, 3)) != 1]
    else:
        names = [a.store]
    for sn in names:
        try:
            root = oc.get_folder(a.folder, sn, ns)
        except SystemExit:
            if a.allstores:
                continue  # a store without that folder (e.g. a PST with no Inbox) is skipped
            raise
        folders.extend(oc.mail_folders_recursive(root) if a.allfolders else [root])
    return folders


def run(a, ns=None):
    ns = ns or oc.connect()
    after = oc.parse_date(a.after) if a.after else None
    before = oc.parse_date(a.before) if a.before else None
    dasl = build_dasl(a, after, before)
    folders = resolve_folders(a, ns)
    results = []
    for f in folders:
        if len(results) >= a.max:
            break
        items = f.Items
        if dasl:
            try:
                items = items.Restrict(dasl)
            except Exception:
                if not (after or before):
                    raise
                dasl = build_dasl(a)  # this store did not take the date literal: filter dates in the loop only
                if dasl:
                    items = items.Restrict(dasl)
        items.Sort("[ReceivedTime]", True)  # newest first
        for item in oc.iter_mail(items):
            if len(results) >= a.max:
                break
            rt = oc.to_datetime(item.ReceivedTime)
            if after and rt < after:
                break  # sorted desc: nothing older will match
            if before and rt >= before:
                continue
            results.append(oc.mail_summary(item, a.includebody, a.previewlength))
    return {
        "Query": {
            "From": a.from_, "To": a.to, "Subject": a.subject, "Body": a.body, "Text": a.text,
            "AnyOf": [t.strip() for chunk in (a.anyof or []) for t in chunk.split(",") if t.strip()],
            "After": after.strftime("%Y-%m-%dT%H:%M:%S") if after else None,
            "Before": before.strftime("%Y-%m-%dT%H:%M:%S") if before else None,
            "HasAttachments": a.hasattachments, "Unread": a.unread, "HighImportance": getattr(a, "highimportance", False), "Flagged": getattr(a, "flagged", False),
            "Folders": [str(f.FolderPath) for f in folders], "AllStores": a.allstores, "Dasl": dasl, "Max": a.max,
        },
        "Count": len(results), "Results": results,
    }


def parser():
    ap = oc.ArgParser(description=__doc__)
    ap.opt("-From", dest="from_", default="", help="sender name or address contains")
    ap.opt("-To", default="", help="To/CC display string contains")
    ap.opt("-Subject", default="")
    ap.opt("-Body", default="", help="plain-text body contains")
    ap.opt("-Text", default="", help="subject OR body contains")
    ap.opt("-AnyOf", nargs="+", default=[], help="several terms (space or comma separated), subject OR body, joined with OR")
    ap.opt("-After", default="", help="received on/after (ISO date)")
    ap.opt("-Before", default="", help="received before (ISO date, exclusive)")
    ap.flag("-HasAttachments")
    ap.flag("-Unread")
    ap.flag("-HighImportance", help="only high-importance mails")
    ap.flag("-Flagged", help="only flagged (follow-up) mails")
    ap.opt("-Folder", default="", help='"Inbox", "Sent Items", "Inbox/Projects", "收件匣", or "\\\\Store\\Inbox\\Sub"')
    ap.opt("-Store", default="", help="store display name (shared mailbox, PST)")
    ap.flag("-AllFolders", help="recurse every mail folder under -Folder")
    ap.flag("-AllStores", help="search every store, not just one")
    ap.opt("-Max", type=int, default=50)
    ap.flag("-IncludeBody")
    ap.opt("-PreviewLength", type=int, default=200)
    oc.add_common_output(ap)
    return ap


def main(argv=None):
    a = parser().parse_args(argv)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
