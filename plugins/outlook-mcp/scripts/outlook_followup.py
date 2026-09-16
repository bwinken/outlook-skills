#!/usr/bin/env python3
"""READ-ONLY. Mails waiting for a reply.

    python outlook_followup.py -Direction sent      # I sent it, nobody answered for N days (default 3)
    python outlook_followup.py -Direction received  # they asked me something, I have not answered (default 2 days)
    python outlook_followup.py -Direction sent -Days 5 -Store 20230731 -AllStores

A mail counts as answered when a later mail in the same conversation (ConversationID, else
ConversationTopic) comes from the other side. "sent": later mail from anyone but me.
"received": later mail from me. LooksLikeQuestion is a cheap heuristic for the model to weigh.
"""
import datetime as dt
import re

import outlook_com as oc

_Q = re.compile(r"[?？]|請(?:問|回|提供|確認|協助|幫)|麻煩|煩請|可否|能否|是否|\b(please|could you|can you|would you|let me know|kindly|confirm|need your)\b", re.I)


def _conv_key(m):
    return str(oc._safe(lambda: m.ConversationID, "") or "") or ("T:" + str(oc._safe(lambda: m.ConversationTopic, "") or ""))


def _mail_folders(ns, a, which):
    """which: 'inbox' or 'sent'. Honour -Store / -AllStores."""
    fid = oc.OL_FOLDER["Inbox"] if which == "inbox" else oc.OL_FOLDER["SentMail"]
    stores = [s for s in oc.get_stores(ns) if int(oc._safe(lambda: s.ExchangeStoreType, 3)) != 1] if a.allstores else ([oc.find_store(a.store, ns)] if a.store else [None])
    folders = []
    for st in stores:
        try:
            root = st.GetDefaultFolder(fid) if st is not None else ns.GetDefaultFolder(fid)
        except Exception:
            continue
        folders.extend(oc.mail_folders_recursive(root) if which == "inbox" else [root])
    return folders


def _scan(folders, since, max_items):
    out = []
    for f in folders:
        items = f.Items
        items.Sort("[ReceivedTime]", True)
        for m in oc.iter_items(items):
            if len(out) >= max_items:
                return out
            if int(oc._safe(lambda: m.Class, 0)) != oc.OL_MAIL_ITEM:
                continue
            if oc.to_datetime(m.ReceivedTime) < since:
                break
            out.append(m)
    return out


def run(a, ns=None):
    ns = ns or oc.connect()
    if a.days is None:
        a.days = 3 if a.direction == "sent" else 2
    me = oc.my_addresses(ns)
    now = dt.datetime.now()
    cutoff = now - dt.timedelta(days=a.days)
    since = now - dt.timedelta(days=a.lookback)

    inbox_mails = _scan(_mail_folders(ns, a, "inbox"), since, a.maxitems)
    sent_mails = _scan(_mail_folders(ns, a, "sent"), since, a.maxitems)

    # conversation -> list of (time, from_address, mail)
    conv = {}
    for m in inbox_mails + sent_mails:
        conv.setdefault(_conv_key(m), []).append((oc.to_datetime(m.ReceivedTime), oc.sender_smtp(m).lower(), m))
    for k in conv:
        conv[k].sort(key=lambda t: t[0])

    results = []
    if a.direction == "sent":
        for m in sent_mails:
            t = oc.to_datetime(m.ReceivedTime)
            if t > cutoff:
                continue  # too recent to chase
            later_from_others = [x for x in conv.get(_conv_key(m), []) if x[0] > t and x[1] and x[1] not in me]
            if later_from_others:
                continue
            later_from_me = [x for x in conv.get(_conv_key(m), []) if x[0] > t and x[1] in me]
            s = oc.mail_summary(m, False, a.previewlength)
            s["Counterparts"] = oc.recipient_list(m, 1)
            s["WaitingDays"] = (now - t).days
            s["MyLaterNudges"] = len(later_from_me)
            results.append(s)
    else:
        for m in inbox_mails:
            t = oc.to_datetime(m.ReceivedTime)
            frm = oc.sender_smtp(m).lower()
            if not frm or frm in me:
                continue
            if t > cutoff:
                continue
            later_from_me = [x for x in conv.get(_conv_key(m), []) if x[0] > t and x[1] in me]
            if later_from_me:
                continue
            later_from_them = [x for x in conv.get(_conv_key(m), []) if x[0] > t and x[1] == frm]
            s = oc.mail_summary(m, False, a.previewlength)
            body = oc._safe(lambda: str(m.Body), "") or ""
            s["LooksLikeQuestion"] = bool(_Q.search(body[:2000] + " " + s["Subject"]))
            s["DirectToMe"] = any(r["Address"].lower() in me for r in oc.recipient_list(m, 1))
            s["WaitingDays"] = (now - t).days
            s["TheirLaterNudges"] = len(later_from_them)
            if a.questionsonly and not s["LooksLikeQuestion"]:
                continue
            results.append(s)
    results.sort(key=lambda x: -x["WaitingDays"])
    return {
        "Direction": a.direction, "Days": a.days, "Lookback": a.lookback, "Me": sorted(me),
        "Scanned": {"Inbox": len(inbox_mails), "Sent": len(sent_mails)},
        "Count": len(results), "Results": results[:a.max],
    }


def parser():
    ap = oc.ArgParser(description=__doc__)
    ap.opt("-Direction", choices=["sent", "received"], default="sent")
    ap.opt("-Days", type=int, default=None, help="minimum age in days before a mail counts as waiting (sent: 3, received: 2)")
    ap.opt("-Lookback", type=int, default=60, help="how far back to scan, days")
    ap.opt("-Store", default="")
    ap.flag("-AllStores")
    ap.flag("-QuestionsOnly", help="received: keep only mails that look like a question or request")
    ap.opt("-MaxItems", type=int, default=2000, help="scan cap per folder set")
    ap.opt("-Max", type=int, default=50)
    ap.opt("-PreviewLength", type=int, default=200)
    oc.add_common_output(ap)
    return ap


def main(argv=None):
    a = parser().parse_args(argv)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
