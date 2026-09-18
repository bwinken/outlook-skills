#!/usr/bin/env python3
"""READ-ONLY. Mails waiting for a reply.

    python outlook_followup.py -Direction sent      # I sent it, nobody answered for N days (default 3)
    python outlook_followup.py -Direction received  # they asked me something, I have not answered (default 2 days)
    python outlook_followup.py -Direction both      # both lists from one scan (Sent / Received keys)
    python outlook_followup.py -Direction sent -Days 5 -Store 20230731 -AllStores

A mail counts as answered when a later mail in the same conversation (ConversationID, else
ConversationTopic) comes from the other side. "sent": later mail from anyone but me.
"received": later mail from me. LooksLikeQuestion is a cheap heuristic for the model to weigh.

Inbox and Sent Items are read through Outlook's Table object (no item opened). Only the mails that
make it into a result open their item: one Body read serves the preview and the question check.
"""
import datetime as dt
import re

import outlook_com as oc

_Q = re.compile(r"[?？]|請(?:問|回|提供|確認|協助|幫)|麻煩|煩請|可否|能否|是否|\b(please|could you|can you|would you|let me know|kindly|confirm|need your)\b", re.I)


def _conv_key(row):
    s = row.summary
    return s["ConversationID"] or ("T:" + s["ConversationTopic"])


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


def _scan(folders, since, max_items, extra=()):
    out = []
    for f in folders:
        if len(out) >= max_items:
            break
        out.extend(oc.scan_mail(f, after=since, limit=max_items - len(out), extra=extra))
    return out


def _waiting_sent(sent_rows, conv, me, cutoff, now, a, ns):
    waiting = []
    for row in sent_rows:
        t = row.received
        if t is None or t > cutoff:
            continue  # too recent to chase
        later = conv.get(_conv_key(row), [])
        if any(x[0] > t and x[1] and x[1] not in me for x in later):
            continue
        waiting.append((t, row, later))
    waiting.sort(key=lambda x: x[0])  # longest waiting first
    out = []
    for t, row, later in waiting[:a.max]:
        s = oc.enrich(row, preview_length=a.previewlength, ns=ns)
        s["Counterparts"] = oc.recipient_list(oc.open_item(row, ns), 1)
        s["WaitingDays"] = (now - t).days
        s["MyLaterNudges"] = sum(1 for x in later if x[0] > t and x[1] in me)
        out.append(s)
    return out


def _waiting_received(inbox_rows, conv, me, cutoff, now, a, ns):
    waiting = []
    for row in inbox_rows:
        t = row.received
        frm = oc.sender_address(row, ns)
        if not frm or frm in me or t is None or t > cutoff:
            continue
        later = conv.get(_conv_key(row), [])
        if any(x[0] > t and x[1] in me for x in later):
            continue
        waiting.append((t, row, frm, later))
    waiting.sort(key=lambda x: x[0])
    out = []
    for t, row, frm, later in waiting:
        if len(out) >= a.max:
            break
        s = oc.enrich(row, body=True, preview_length=a.previewlength, ns=ns)  # one Body read: preview and question check
        body = s.pop("Body", "")
        s["LooksLikeQuestion"] = bool(_Q.search(body[:2000] + " " + s["Subject"]))
        if a.questionsonly and not s["LooksLikeQuestion"]:
            continue
        s["DirectToMe"] = oc.to_me(row, me, ns, open_if_needed=True)  # the item is open already for the body
        s["WaitingDays"] = (now - t).days
        s["TheirLaterNudges"] = sum(1 for x in later if x[0] > t and x[1] == frm)
        out.append(s)
    return out


def run(a, ns=None):
    """One scan of Inbox and Sent Items serves both directions; -Direction both returns them together
    (Sent / Received keys) so a morning brief needs one run instead of two."""
    ns = ns or oc.connect()
    me = oc.my_addresses(ns)
    now = dt.datetime.now()
    since = now - dt.timedelta(days=a.lookback)
    days_sent = a.days if a.days is not None else 3
    days_received = a.days if a.days is not None else 2

    inbox_rows = _scan(_mail_folders(ns, a, "inbox"), since, a.maxitems, extra=("ToMe",))
    sent_rows = _scan(_mail_folders(ns, a, "sent"), since, a.maxitems)

    # conversation -> list of (time, from_address, row)
    conv = {}
    for row in inbox_rows + sent_rows:
        if row.received is not None:
            conv.setdefault(_conv_key(row), []).append((row.received, oc.sender_address(row, ns), row))
    for k in conv:
        conv[k].sort(key=lambda t: t[0])

    out = {"Direction": a.direction, "Lookback": a.lookback, "Me": sorted(me),
           "Scanned": {"Inbox": len(inbox_rows), "Sent": len(sent_rows)}}
    if a.direction in ("sent", "both"):
        rows = _waiting_sent(sent_rows, conv, me, now - dt.timedelta(days=days_sent), now, a, ns)
        if a.direction == "sent":
            out.update({"Days": days_sent, "Count": len(rows), "Results": rows})
        else:
            out["Sent"] = {"Days": days_sent, "Count": len(rows), "Results": rows}
    if a.direction in ("received", "both"):
        rows = _waiting_received(inbox_rows, conv, me, now - dt.timedelta(days=days_received), now, a, ns)
        if a.direction == "received":
            out.update({"Days": days_received, "Count": len(rows), "Results": rows})
        else:
            out["Received"] = {"Days": days_received, "Count": len(rows), "Results": rows}
    return out


def parser():
    ap = oc.ArgParser(description=__doc__)
    ap.opt("-Direction", choices=["sent", "received", "both"], default="sent", help="both: one scan, results under Sent and Received")
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
    oc.apply_settings(a)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
