#!/usr/bin/env python3
"""READ-ONLY. Find attachments across mails; list, filter, sort, optionally copy out.

    python outlook_attachments.py -Name 合約 -After 2026-06-01
    python outlook_attachments.py -Ext pptx,xlsx -MinSizeKB 500 -Sort size -Top 20 -AllStores -AllFolders
    python outlook_attachments.py -From alice -SaveTo C:/tmp/att      # copies files out; Outlook is not modified
"""
import os
import re

import outlook_com as oc
import outlook_search


def run(a, ns=None):
    ns = ns or oc.connect()
    a.hasattachments = True
    a.includebody = False
    search = outlook_search.run(a, ns)
    exts = {e.strip().lower().lstrip(".") for chunk in (a.ext or []) for e in chunk.split(",") if e.strip()}
    name_re = re.compile(re.escape(a.name), re.I) if a.name else None
    rows = []
    # We need the live items again for SaveAsFile; mail_summary already gave us names/sizes.
    by_id = {}
    if a.saveto:
        for m in search["Results"]:
            by_id[m["EntryID"]] = ns.GetItemFromID(m["EntryID"])
    for m in search["Results"]:
        for i, att in enumerate(m["Attachments"]):
            fn = att["FileName"]
            ext = os.path.splitext(fn)[1].lower().lstrip(".")
            if int(att.get("Type", 1)) != 1 and not a.includeembedded:
                continue  # skip inline images / OLE unless asked
            if exts and ext not in exts:
                continue
            if name_re and not name_re.search(fn):
                continue
            if att["Size"] < a.minsizekb * 1024:
                continue
            row = {"FileName": fn, "Ext": ext, "SizeKB": round(att["Size"] / 1024, 1), "ReceivedTime": m["ReceivedTime"],
                   "From": m["From"], "FromAddress": m["FromAddress"], "Subject": m["Subject"], "Folder": m["Folder"], "EntryID": m["EntryID"]}
            if a.saveto:
                row["SavedTo"] = _save(by_id[m["EntryID"]], i, fn, a.saveto)
            rows.append(row)
    if a.sort == "size":
        rows.sort(key=lambda r: -r["SizeKB"])
    elif a.sort == "name":
        rows.sort(key=lambda r: r["FileName"].lower())
    else:
        rows.sort(key=lambda r: r["ReceivedTime"] or "", reverse=True)
    total_kb = sum(r["SizeKB"] for r in rows)
    return {"Query": {**search["Query"], "Name": a.name, "Ext": sorted(exts), "MinSizeKB": a.minsizekb, "Sort": a.sort, "Top": a.top, "SaveTo": a.saveto},
            "MailsScanned": search["Count"], "Count": len(rows), "TotalSizeKB": round(total_kb, 1), "Results": rows[:a.top]}


def _save(mail, index, filename, folder):
    os.makedirs(folder, exist_ok=True)
    safe = re.sub(r'[\\/:*?"<>|]', "_", filename) or "attachment"
    target = os.path.join(folder, safe)
    base, ext = os.path.splitext(target)
    n = 1
    while os.path.exists(target):
        target = f"{base}({n}){ext}"
        n += 1
    att = mail.Attachments.Item(index + 1)  # COM collections are 1-based
    att.SaveAsFile(os.path.abspath(target))   # reads from Outlook, writes only to the target folder
    return os.path.abspath(target)


def parser():
    ap = outlook_search.parser()
    ap.description = __doc__
    ap.opt("-Name", default="", help="attachment file name contains")
    ap.opt("-Ext", nargs="+", default=[], help="extensions, e.g. pdf,xlsx")
    ap.opt("-MinSizeKB", type=int, default=0)
    ap.opt("-Sort", choices=["date", "size", "name"], default="date")
    ap.opt("-Top", type=int, default=50)
    ap.flag("-IncludeEmbedded", help="include inline images and OLE objects")
    ap.opt("-SaveTo", default="", help="copy matching attachments into this folder (ask the user first)")
    ap.set_defaults(max=500, allfolders=False)
    return ap


def main(argv=None):
    a = parser().parse_args(argv)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
