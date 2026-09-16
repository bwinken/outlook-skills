#!/usr/bin/env python3
"""Send a new mail or a reply. The ONLY script in this plugin that writes to Outlook, and it can only
do so after the user clicks "Send" in a confirmation window on their own desktop.

Two steps, always:

    python outlook_send.py draft --to alice@contoso.com [--cc bob] --subject "..." --body-file body.txt
    python outlook_send.py draft --reply-to <EntryID> [--reply-all] [--cc extra] --body-file body.txt
        Resolves every recipient through Outlook, builds the exact outgoing text (body + footer + quoted
        original for replies), stores it as ~/.outlook-skills/drafts/<id>.json and prints it. Nothing is
        sent, nothing in Outlook changes.

    python outlook_send.py send <id> --confirm <token>
        Opens a window on the Windows desktop showing To, Cc, Subject and the full outgoing text with a
        Send and a Cancel button, waits for the click (or the timeout, which cancels), then creates the
        item through COM with exactly the stored recipients, subject and text, re-reads them from the
        item, aborts on any difference, and only then calls Send(). There is no flag, variable or setting
        that skips the window.

    python outlook_send.py show <id> | list | discard <id>

The token printed by `draft` is a hash of the outgoing content; `send` recomputes it, so a draft file
edited after it was shown cannot be sent. Attachments are not supported.
"""
import datetime as dt
import hashlib
import json
import os
import platform
import secrets

import outlook_com as oc
import settings as ps

OL_TO, OL_CC = 1, 2
OL_FORMAT_PLAIN = 1


# ---------------------------------------------------------------- storage
def drafts_dir():
    return ps.user_dir() / "drafts"


def _path(draft_id: str):
    if not draft_id or any(c in draft_id for c in "/\\"):
        raise SystemExit(f"bad draft id '{draft_id}'")
    return drafts_dir() / f"{draft_id}.json"


def load(draft_id: str) -> dict:
    p = _path(draft_id)
    if not p.is_file():
        raise SystemExit(f"no draft '{draft_id}' in {drafts_dir()}")
    with open(p, "r", encoding="utf-8") as fh:
        return json.load(fh)


def save(d: dict):
    drafts_dir().mkdir(parents=True, exist_ok=True)
    with open(_path(d["id"]), "w", encoding="utf-8") as fh:
        fh.write(json.dumps(d, ensure_ascii=False, indent=2))


def token(d: dict) -> str:
    """Hash of everything that goes out. Recomputed by `send`; a mismatch aborts."""
    key = json.dumps({"to": d["to"], "cc": d["cc"], "subject": d["subject"], "full_body": d["full_body"]}, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(key.encode("utf-8")).hexdigest()[:12]


# ---------------------------------------------------------------- recipients
def _addr_of(recipient) -> str:
    for getter in (lambda: recipient.PropertyAccessor.GetProperty(oc.PR_SMTP_ADDRESS),
                   lambda: recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress,
                   lambda: recipient.Address):
        v = oc._safe(getter, "")
        if v and "@" in str(v):
            return str(v)
    return ""


def resolve_recipients(texts, ns):
    """Each text is an address or a display name; Outlook resolves it (address book, GAL). Unresolved
    or ambiguous names abort so the user confirms real addresses, never guesses."""
    out, bad = [], []
    for raw in texts:
        for t in [x.strip() for x in raw.replace(";", ",").split(",") if x.strip()]:
            r = ns.CreateRecipient(t)
            if not oc._safe(lambda: r.Resolve(), False):
                bad.append(t)
                continue
            addr = _addr_of(r)
            if not addr:
                bad.append(t)
                continue
            out.append({"Name": str(oc._safe(lambda: r.Name, "") or t), "Address": addr})
    if bad:
        raise SystemExit("Cannot resolve recipient(s): " + ", ".join(bad) + ". Give the full address, or a name Outlook knows.")
    return _dedupe(out)


def _dedupe(rows):
    seen, out = set(), []
    for r in rows:
        k = r["Address"].lower()
        if k and k not in seen:
            seen.add(k)
            out.append({"Name": r["Name"], "Address": r["Address"]})
    return out


def _addresses(rows):
    return sorted(r["Address"].lower() for r in rows)


# ---------------------------------------------------------------- draft
def _approver(ns, cfg) -> str:
    v = cfg.get("approver")
    if v:
        return str(v)
    for getter in (lambda: ns.CurrentUser.Name, lambda: ns.Accounts[0].DisplayName):
        v = oc._safe(getter, "")
        if v:
            return str(v)
    return os.environ.get("USERNAME") or os.environ.get("USER") or "the user"


def _quote(orig) -> str:
    head = [
        "-----Original Message-----",
        f"From: {oc._safe(lambda: str(orig.SenderName), '')} <{oc.sender_smtp(orig)}>",
        f"Sent: {oc.iso(orig.ReceivedTime) or ''}",
        f"To: {oc._safe(lambda: str(orig.To), '')}",
    ]
    cc = oc._safe(lambda: str(orig.CC), "")
    if cc:
        head.append(f"Cc: {cc}")
    head.append(f"Subject: {oc._safe(lambda: str(orig.Subject), '')}")
    body = oc._safe(lambda: str(orig.Body), "") or ""
    return "\n".join(head) + "\n\n" + body.replace("\r\n", "\n").rstrip()


def _reply_subject(subject: str) -> str:
    s = (subject or "").strip()
    return s if s.lower().startswith(("re:", "回覆:", "回覆：", "答复:", "答复：")) else f"RE: {s}"


def run_draft(a, ns=None):
    ns = ns or oc.connect()
    cfg = (ps.resolve()[0].get("send") or {})
    if a.body_file:
        with open(a.body_file, "r", encoding="utf-8-sig") as fh:
            body = fh.read()
    else:
        body = a.body or ""
    body = body.replace("\r\n", "\n").strip()
    if not body:
        raise SystemExit("Empty body. Pass --body-file <file> (recommended) or --body <text>.")
    me = oc.my_addresses(ns)

    orig, quote, reply = None, "", {}
    if a.reply_to:
        orig = ns.GetItemFromID(a.reply_to)
        sender = {"Name": str(oc._safe(lambda: orig.SenderName, "") or ""), "Address": oc.sender_smtp(orig)}
        to = [sender] if sender["Address"] else []
        cc = []
        if a.reply_all:
            to += [r for r in oc.recipient_list(orig, OL_TO) if r["Address"] and not oc.is_me(r["Address"], me)]
            cc = [r for r in oc.recipient_list(orig, OL_CC) if r["Address"] and not oc.is_me(r["Address"], me)]
        to = _dedupe(to + resolve_recipients(a.to or [], ns))
        cc = _dedupe([r for r in cc + resolve_recipients(a.cc or [], ns) if r["Address"].lower() not in _addresses(to)])
        subject = a.subject or _reply_subject(str(oc._safe(lambda: orig.Subject, "") or ""))
        if cfg.get("quote_original", True):
            quote = _quote(orig)
        reply = {"EntryID": a.reply_to, "From": sender, "Subject": str(oc._safe(lambda: orig.Subject, "") or ""),
                 "ReceivedTime": oc.iso(orig.ReceivedTime), "ReplyAll": bool(a.reply_all)}
    else:
        if not a.to:
            raise SystemExit("--to is required for a new mail (or --reply-to <EntryID> for a reply).")
        if not a.subject:
            raise SystemExit("--subject is required for a new mail.")
        to = resolve_recipients(a.to, ns)
        cc = _dedupe([r for r in resolve_recipients(a.cc or [], ns) if r["Address"].lower() not in _addresses(to)])
        subject = a.subject.strip()
    if not to:
        raise SystemExit("No recipient.")

    approver = _approver(ns, cfg)
    footer = str(cfg.get("footer") or "--\nDrafted by Claude, reviewed and approved by {approver}.").replace("\\n", "\n").format(approver=approver)
    full_body = body + "\n\n" + footer + ("\n\n" + quote if quote else "")
    d = {
        "id": dt.datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + secrets.token_hex(2),
        "status": "draft", "created": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "mode": ("reply_all" if a.reply_all else "reply") if a.reply_to else "new",
        "reply_to": reply or None,
        "to": to, "cc": cc, "subject": subject,
        "body": body, "footer": footer, "quote": quote, "full_body": full_body,
        "approver": approver,
    }
    d["confirm"] = token(d)
    save(d)
    return d


# ---------------------------------------------------------------- the confirmation window
def render_text(d: dict) -> str:
    fmt = lambda rows: "; ".join(f"{r['Name']} <{r['Address']}>" if r["Name"] and r["Name"] != r["Address"] else r["Address"] for r in rows) or "(none)"
    return f"To:      {fmt(d['to'])}\nCc:      {fmt(d['cc'])}\nSubject: {d['subject']}\n\n{d['full_body']}"


def confirm_dialog(d: dict, timeout_seconds: int) -> bool:
    """A window on the user's desktop. Returns True only when the user clicks Send. Closing the window,
    Cancel, or the timeout all return False. Tests replace this function; nothing else does."""
    if platform.system() != "Windows":
        raise SystemExit("Sending needs Windows with Classic Outlook.")
    try:
        import tkinter as tk
        from tkinter import scrolledtext
    except Exception:
        return _messagebox(d)
    result = {"ok": False}
    root = tk.Tk()
    root.title("Outlook 確認寄出 / Confirm send")
    root.attributes("-topmost", True)
    head = tk.Frame(root, padx=12, pady=8)
    head.pack(fill="x")
    fmt = lambda rows: "; ".join(f"{r['Name']} <{r['Address']}>" for r in rows) or "(none)"
    for label, value in (("收件者 To", fmt(d["to"])), ("副本 Cc", fmt(d["cc"])), ("主旨 Subject", d["subject"])):
        row = tk.Frame(head)
        row.pack(fill="x")
        tk.Label(row, text=label + ":", width=12, anchor="w", font=("", 10, "bold")).pack(side="left")
        tk.Label(row, text=value, anchor="w", justify="left", wraplength=760).pack(side="left", fill="x")
    box = scrolledtext.ScrolledText(root, width=100, height=28, wrap="word", font=("", 10))
    box.insert("1.0", d["full_body"])
    box.configure(state="disabled")
    box.pack(fill="both", expand=True, padx=12)
    foot = tk.Frame(root, padx=12, pady=10)
    foot.pack(fill="x")
    tk.Label(foot, text=f"這封信會以上面的內容原樣寄出，寄出後無法收回。{timeout_seconds // 60} 分鐘沒動作視為取消。", anchor="w").pack(side="left")

    def send():
        result["ok"] = True
        root.destroy()

    tk.Button(foot, text="取消 Cancel", width=14, command=root.destroy).pack(side="right")
    tk.Button(foot, text="寄出 Send", width=14, command=send, default="active").pack(side="right", padx=8)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.after(max(10, int(timeout_seconds)) * 1000, root.destroy)
    root.lift()
    root.focus_force()
    root.mainloop()
    return result["ok"]


def _messagebox(d: dict) -> bool:
    """Fallback when tkinter is missing: a Yes/No message box (text limited, long bodies are cut)."""
    import ctypes
    text = render_text(d)
    if len(text) > 3500:
        text = text[:3500] + "\n\n[... cut for display; the mail itself is complete ...]"
    MB_YESNO, MB_ICONWARNING, MB_DEFBUTTON2, MB_SETFOREGROUND, MB_TOPMOST = 0x4, 0x30, 0x100, 0x10000, 0x40000
    r = ctypes.windll.user32.MessageBoxW(0, text + "\n\n寄出？ Send this mail?", "Outlook 確認寄出 / Confirm send", MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2 | MB_SETFOREGROUND | MB_TOPMOST)
    return r == 6  # IDYES


# ---------------------------------------------------------------- send
def _set_recipients(mail, d):
    recips = mail.Recipients
    while int(recips.Count) > 0:
        recips.Remove(1)
    for rtype, rows in ((OL_TO, d["to"]), (OL_CC, d["cc"])):
        for r in rows:
            rec = recips.Add(r["Address"])
            rec.Type = rtype
    if not recips.ResolveAll():
        raise SystemExit("Outlook could not resolve every recipient; nothing was sent.")


def _verify(mail, d):
    """Read back what Outlook holds and compare with the draft. Any difference means no Send()."""
    got_to = sorted(_addr_of(r).lower() or str(r.Address).lower() for r in mail.Recipients if int(r.Type) == OL_TO)
    got_cc = sorted(_addr_of(r).lower() or str(r.Address).lower() for r in mail.Recipients if int(r.Type) == OL_CC)
    problems = []
    if got_to != _addresses(d["to"]):
        problems.append(f"To differs: item {got_to} vs draft {_addresses(d['to'])}")
    if got_cc != _addresses(d["cc"]):
        problems.append(f"Cc differs: item {got_cc} vs draft {_addresses(d['cc'])}")
    if str(mail.Subject) != d["subject"]:
        problems.append("Subject differs")
    if str(mail.Body).replace("\r\n", "\n").rstrip() != d["full_body"].rstrip():
        problems.append("Body differs")
    if problems:
        raise SystemExit("Item does not match the approved draft, nothing was sent: " + "; ".join(problems))


def run_send(a, ns=None):
    d = load(a.id)
    if d.get("status") != "draft":
        raise SystemExit(f"draft {d['id']} is '{d.get('status')}', not sendable. Make a new draft.")
    if a.confirm != d["confirm"] or token(d) != d["confirm"]:
        raise SystemExit("Confirmation token does not match the draft as shown; nothing was sent. Re-run draft and show it again.")
    cfg = (ps.resolve()[0].get("send") or {})
    timeout = int(cfg.get("dialog_timeout_seconds") or 300)

    # the user's click, on their own screen; there is no way around this call
    if not confirm_dialog(d, timeout):
        d["status"], d["cancelled"] = "cancelled", dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        save(d)
        raise SystemExit("Cancelled in the confirmation window (or it timed out); nothing was sent. Ask the user before making a new draft.")

    ns = ns or oc.connect()
    app = oc.application()
    if d["mode"] == "new":
        mail = app.CreateItem(oc.OL_ITEM_TYPE_MAIL)
    else:
        orig = ns.GetItemFromID(d["reply_to"]["EntryID"])
        mail = orig.ReplyAll() if d["mode"] == "reply_all" else orig.Reply()
    _set_recipients(mail, d)
    mail.Subject = d["subject"]
    try:
        mail.BodyFormat = OL_FORMAT_PLAIN
    except Exception:
        pass
    mail.Body = d["full_body"]
    _verify(mail, d)
    mail.Send()
    d["status"], d["sent_at"] = "sent", dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    save(d)
    return {"Sent": True, "Id": d["id"], "Mode": d["mode"], "To": d["to"], "Cc": d["cc"], "Subject": d["subject"], "SentAt": d["sent_at"], "Approver": d["approver"]}


# ---------------------------------------------------------------- other commands
def run_show(a):
    return load(a.id)


def run_list(a):
    dd = drafts_dir()
    rows = []
    for p in sorted(dd.glob("*.json")) if dd.is_dir() else []:
        try:
            with open(p, "r", encoding="utf-8") as fh:
                d = json.load(fh)
            rows.append({k: d.get(k) for k in ("id", "status", "created", "mode", "subject", "sent_at")} | {"To": [r["Address"] for r in d.get("to", [])]})
        except Exception:
            continue
    return {"Dir": str(dd), "Count": len(rows), "Drafts": rows}


def run_discard(a):
    d = load(a.id)
    if d.get("status") == "draft":
        d["status"] = "discarded"
        save(d)
    return {"Id": d["id"], "Status": d["status"]}


# ---------------------------------------------------------------- parsers
def draft_parser():
    ap = oc.ArgParser(description="Build and store a draft; sends nothing. Prints the draft with its id and confirm token.")
    ap.opt("-To", nargs="+", default=[], help="recipients: addresses or names Outlook can resolve (comma or space separated). Required for a new mail")
    ap.opt("-Cc", nargs="+", default=[], help="Cc recipients")
    ap.opt("-Subject", default="", help="required for a new mail; replies default to RE: <original>")
    ap.opt("-ReplyTo", dest="reply_to", default="", help="EntryID of the mail to answer (from outlook_search / outlook_thread)")
    ap.flag("-ReplyAll", dest="reply_all", help="answer everyone on the original (the user's own addresses excluded)")
    ap.opt("-BodyFile", dest="body_file", default="", help="UTF-8 text file with the body (recommended: no shell quoting issues)")
    ap.opt("-Body", default="", help="body text inline (short bodies only)")
    oc.add_common_output(ap)
    return ap


def send_parser():
    ap = oc.ArgParser(description="Show the confirmation window for a stored draft and send it after the user's click.")
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
    p = sub.add_parser("show"); p.add_argument("id"); p.set_defaults(fn=run_show, out_file="")
    p = sub.add_parser("list"); p.set_defaults(fn=run_list, out_file="")
    p = sub.add_parser("discard"); p.add_argument("id"); p.set_defaults(fn=run_discard, out_file="")
    a = ap.parse_args(argv)
    oc.write_json(a.fn(a), getattr(a, "out_file", ""))


if __name__ == "__main__":
    main()
