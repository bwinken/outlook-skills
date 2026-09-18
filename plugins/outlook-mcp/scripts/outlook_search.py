#!/usr/bin/env python3
"""READ-ONLY. Searches mail by sender, recipient, subject, body, date range, attachments, unread state.

    python outlook_search.py -From alice -After 2026-09-01 -Max 20
    python outlook_search.py -Subject invoice -HasAttachments -AllFolders
    python outlook_search.py -AnyOf 報價,quote,pricing -After 2026-06-01 -AllStores -Max 300 -PreviewLength 500 -OutFile candidates.json
    python outlook_search.py -EntryID <id> -IncludeBody      # exactly that mail, no folder walk and no filter

Each folder is read through Outlook's Table object, hundreds of rows per call. A mail is opened only
for a body preview (-PreviewLength), the full body (-IncludeBody) or its attachment list.
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


def mail_by_id(entry_id: str, ns=None):
    """One mail by EntryID (from an earlier result), with a plain message instead of a COM error when it is not there."""
    ns = ns or oc.connect()
    try:
        item = ns.GetItemFromID(entry_id)
    except Exception as e:
        raise SystemExit(f"No item with that EntryID in this Outlook profile ({e}). Use an EntryID from a search or thread result.")
    if int(oc._safe(lambda: item.Class, 0)) != oc.OL_MAIL_ITEM:
        raise SystemExit("That EntryID is not a mail item.")
    return item


def run(a, ns=None):
    ns = ns or oc.connect()
    after = oc.parse_date(a.after) if a.after else None
    before = oc.parse_date(a.before) if a.before else None
    entry_id = getattr(a, "entryid", "") or ""
    if entry_id:
        # exactly that mail, no folder walk and no filter: a filtered scan can miss it (a mail whose only
        # attachments are inline pictures gets no paperclip, so -HasAttachments may not match it)
        dasl, results = "", [oc.mail_summary(mail_by_id(entry_id, ns), a.includebody, a.previewlength)]
        folder_paths = [results[0]["Folder"]]
    else:
        dasl, dasl_no_dates = build_dasl(a, after, before), build_dasl(a)
        folders = resolve_folders(a, ns)
        folder_paths = [str(f.FolderPath) for f in folders]
        results = []
        for f, path in zip(folders, folder_paths):
            if len(results) >= a.max:
                break
            # newest first; the store gets the filter, the item is opened only for what the Table cannot give
            for row in oc.scan_mail(f, dasl, after=after, before=before, limit=a.max - len(results), dasl_fallback=dasl_no_dates, folder_path=path):
                results.append(oc.enrich(row, body=a.includebody, preview_length=a.previewlength, ns=ns))
    return {
        "Query": {
            "EntryID": entry_id or None,
            "From": a.from_, "To": a.to, "Subject": a.subject, "Body": a.body, "Text": a.text,
            "AnyOf": [t.strip() for chunk in (a.anyof or []) for t in chunk.split(",") if t.strip()],
            "After": after.strftime("%Y-%m-%dT%H:%M:%S") if after else None,
            "Before": before.strftime("%Y-%m-%dT%H:%M:%S") if before else None,
            "HasAttachments": a.hasattachments, "Unread": a.unread, "HighImportance": getattr(a, "highimportance", False), "Flagged": getattr(a, "flagged", False),
            "Folders": folder_paths, "AllStores": a.allstores, "Dasl": dasl, "Max": a.max,
        },
        "Count": len(results), "Results": results,
    }


def parser():
    ap = oc.ArgParser(description=__doc__)
    ap.opt("-EntryID", default="", help="read exactly this mail (EntryID from an earlier result) instead of searching; every other filter is ignored")
    ap.opt("-From", dest="from_", default="", help="sender name or address contains")
    ap.opt("-To", default="", help="To/CC display string contains")
    ap.opt("-Subject", default="")
    ap.opt("-Body", default="", help="plain-text body contains (no index: every body in the range is read by Outlook)")
    ap.opt("-Text", default="", help="subject OR body contains (body part is unindexed, see -Body)")
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
    ap.flag("-IncludeBody", help="full plain-text body of each result (opens every result)")
    ap.opt("-PreviewLength", type=int, default=0, help="characters of BodyPreview per result; 0 = no body read; use ~500 for reranking")
    oc.add_common_output(ap)
    return ap


def main(argv=None):
    a = parser().parse_args(argv)
    oc.apply_settings(a)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
