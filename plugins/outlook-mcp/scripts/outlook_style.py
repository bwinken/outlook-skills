#!/usr/bin/env python3
"""READ-ONLY. Statistics about how the user replies, from Sent Items and Inbox, for the writing profile.

    python outlook_style.py -Days 180 -OutFile style.json
    python outlook_style.py -Store 20230731

Reports: reply rate overall and per sender (who gets answered, who never does), reply latency,
typical length, language mix, common greetings and closings, the signature block, sending hours.
No bodies are returned, only aggregates and short recurring lines.

Both folders are read through Outlook's Table object (no item opened); reply rates and latency come
from those rows alone. Only the newest -BodySamples sent mails are opened, for the text statistics.
"""
import datetime as dt
import re
import statistics
from collections import Counter

import outlook_com as oc

_QUOTE_CUT = re.compile(r"^(From:|寄件者:|發件人:|-----Original Message-----|-----原始郵件-----|On .{3,80} wrote:|在 .{3,80} 寫道：)", re.M)
_SIG_HINT = re.compile(r"^(--\s*$|Best regards|Regards|Thanks|Thank you|敬祝|祝 |順頌|此致|謝謝|感謝|Sincerely|BR,|Cheers)", re.I)


def _own_text(body: str) -> str:
    m = _QUOTE_CUT.search(body or "")
    return (body[:m.start()] if m else body or "").strip()


def _lines(text: str):
    return [ln.strip() for ln in text.replace("\r", "").split("\n") if ln.strip()]


def _cjk_ratio(text: str) -> float:
    letters = [c for c in text if c.isalpha()]
    if not letters:
        return 0.0
    return round(sum(1 for c in letters if "一" <= c <= "鿿" or "㐀" <= c <= "䶿") / len(letters), 2)


def _conv_key(row):
    s = row.summary
    return s["ConversationID"] or ("T:" + s["ConversationTopic"])


def _scan(folder, since, cap, recurse=True, extra=()):
    out = []
    for f in (oc.mail_folders_recursive(folder) if recurse else [folder]):
        if len(out) >= cap:
            break
        out.extend(oc.scan_mail(f, after=since, limit=cap - len(out), extra=extra))
    return out


def run(a, ns=None):
    ns = ns or oc.connect()
    me = oc.my_addresses(ns)
    since = dt.datetime.now() - dt.timedelta(days=a.days)
    st = oc.find_store(a.store, ns) if a.store else None
    inbox = st.GetDefaultFolder(oc.OL_FOLDER["Inbox"]) if st else ns.GetDefaultFolder(oc.OL_FOLDER["Inbox"])
    sent_f = st.GetDefaultFolder(oc.OL_FOLDER["SentMail"]) if st else ns.GetDefaultFolder(oc.OL_FOLDER["SentMail"])
    received = _scan(inbox, since, a.maxitems, extra=("ToMe", "Unsubscribe"))
    sent = _scan(sent_f, since, a.maxitems, recurse=False)

    conv = {}
    for row in received + sent:
        if row.received is not None:
            conv.setdefault(_conv_key(row), []).append((row.received, oc.sender_address(row, ns), row))

    # ---- reply rate per sender
    per = {}
    newsletters = set()
    for row in received:
        frm = oc.sender_address(row, ns)
        t = row.received
        if not frm or frm in me or t is None:
            continue
        e = per.setdefault(frm, {"Address": frm, "Name": row.summary["From"], "Received": 0, "Replied": 0, "ToMe": 0})
        e["Received"] += 1
        if oc.to_me(row, me, ns):
            e["ToMe"] += 1
        if any(x[0] > t and x[1] in me for x in conv.get(_conv_key(row), [])):
            e["Replied"] += 1
        if row.extra.get("Unsubscribe"):
            newsletters.add(frm)
    for e in per.values():
        e["Rate"] = round(e["Replied"] / e["Received"], 2) if e["Received"] else 0.0
        e["Newsletter"] = e["Address"] in newsletters
    senders = sorted(per.values(), key=lambda x: -x["Received"])
    human = [e for e in senders if not e["Newsletter"]]
    tot_r = sum(e["Received"] for e in human); tot_a = sum(e["Replied"] for e in human)

    # ---- latency and sending hours: from the rows, every sent mail
    latencies, hours = [], Counter()
    for row in sent:
        t = row.received
        if t is None:
            continue
        earlier = [x for x in conv.get(_conv_key(row), []) if x[0] < t and x[1] and x[1] not in me]
        if earlier:
            latencies.append((t - max(x[0] for x in earlier)).total_seconds() / 3600)
        hours[t.hour] += 1

    # ---- length, language, greetings, closings, signature: from the bodies of the newest sent mails
    lengths, cjk = [], []
    greetings, closings, sig_blocks = Counter(), Counter(), Counter()
    sampled = sent[:max(0, a.bodysamples)]
    for row in sampled:
        own = _own_text(oc.enrich(row, body=True, ns=ns).pop("Body", ""))
        ls = _lines(own)
        if not ls:
            continue
        lengths.append(len(own))
        cjk.append(_cjk_ratio(own))
        greetings[ls[0][:40]] += 1
        # signature: trailing block starting at a signature hint, up to 6 lines
        sig_start = None
        for i in range(max(0, len(ls) - 8), len(ls)):
            if _SIG_HINT.search(ls[i]):
                sig_start = i
                break
        if sig_start is not None:
            sig_blocks[" / ".join(ls[sig_start:sig_start + 6])] += 1
            body_lines = ls[:sig_start]
        else:
            body_lines = ls
        if len(body_lines) >= 2:
            closings[body_lines[-1][:40]] += 1

    def buckets(xs):
        return {"Short<200": sum(1 for x in xs if x < 200), "Medium200-800": sum(1 for x in xs if 200 <= x <= 800), "Long>800": sum(1 for x in xs if x > 800)}

    hour_hist = {f"{h:02d}": c for h, c in sorted(hours.items())}
    busy_hours = [h for h, _ in hours.most_common(4)]
    return {
        "Window": {"Since": since.strftime("%Y-%m-%dT%H:%M:%S"), "Days": a.days, "SentAnalysed": len(sent), "ReceivedAnalysed": len(received), "BodySamples": len(sampled)},
        "Me": sorted(me),
        "ReplyRate": {
            "OverallHuman": round(tot_a / tot_r, 2) if tot_r else None,
            "BySender": human[:a.top],
            "AlwaysReplied": [e for e in human if e["Received"] >= 2 and e["Rate"] >= 0.8][:a.top],
            "NeverReplied": [e for e in human if e["Received"] >= 3 and e["Replied"] == 0][:a.top],
            "Newsletters": [e["Address"] for e in senders if e["Newsletter"]][:a.top],
        },
        "ReplyLatencyHours": {
            "Median": round(statistics.median(latencies), 1) if latencies else None,
            "P75": round(sorted(latencies)[int(len(latencies) * 0.75)], 1) if latencies else None,
            "Within1h": round(sum(1 for x in latencies if x <= 1) / len(latencies), 2) if latencies else None,
            "Samples": len(latencies),
        },
        "Length": {"MedianChars": int(statistics.median(lengths)) if lengths else None, "Buckets": buckets(lengths)},
        "Language": {"MedianCjkRatio": round(statistics.median(cjk), 2) if cjk else None, "MostlyChinese": (statistics.median(cjk) >= 0.5) if cjk else None},
        "Greetings": [{"Line": k, "Count": v} for k, v in greetings.most_common(5)],
        "Closings": [{"Line": k, "Count": v} for k, v in closings.most_common(5)],
        "Signature": [{"Block": k, "Count": v} for k, v in sig_blocks.most_common(2)],
        "SendHours": {"Histogram": hour_hist, "Typical": sorted(busy_hours)},
    }


def parser():
    ap = oc.ArgParser(description=__doc__)
    ap.opt("-Days", type=int, default=180)
    ap.opt("-MaxItems", type=int, default=3000)
    ap.opt("-BodySamples", type=int, default=300, help="newest sent mails whose text is read for length, language, greetings, closings and signature")
    ap.opt("-Top", type=int, default=20)
    ap.opt("-Store", default="")
    oc.add_common_output(ap)
    return ap


def main(argv=None):
    a = parser().parse_args(argv)
    oc.apply_settings(a)
    oc.write_json(run(a), a.out_file)


if __name__ == "__main__":
    main()
