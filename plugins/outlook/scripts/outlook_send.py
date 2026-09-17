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
edited after it was shown cannot be sent. Attachments (--attach <file> ...) are recorded with their size
and SHA-256 at draft time and re-hashed before they are attached; a changed or missing file aborts.
A reply is built on Outlook's own Reply()/ReplyAll() item so it threads with the original; the inline
pictures that item inherits from the original (cid: images, signature logos) are removed before the
draft's files are added, so the item carries exactly the approved attachments and the check stays exact.
"""
import datetime as dt
import hashlib
import json
import os
import platform
import secrets
from collections import Counter

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


def new_id() -> str:
    return dt.datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + secrets.token_hex(2)


_TOKEN_FIELDS = {"mail": ("to", "cc", "subject", "full_body", "attachments"),
                 "meeting": ("required", "optional", "subject", "start", "end", "location", "full_body")}


def token(d: dict) -> str:
    """Hash of everything that goes out. Recomputed by `send`; a mismatch aborts."""
    fields = _TOKEN_FIELDS[d.get("kind", "mail")]
    key = json.dumps({k: d[k] for k in fields}, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(key.encode("utf-8")).hexdigest()[:12]


def load_sendable(draft_id: str, kind: str, confirm: str) -> dict:
    """The draft, if it is of this kind, still a draft, and the token matches both the argument and the content."""
    d = load(draft_id)
    if d.get("kind", "mail") != kind:
        raise SystemExit(f"draft {d['id']} is a {d.get('kind', 'mail')} draft; use outlook_{'send' if d.get('kind', 'mail') == 'mail' else 'meeting'}.py")
    if d.get("status") != "draft":
        raise SystemExit(f"draft {d['id']} is '{d.get('status')}', not sendable. Make a new draft.")
    if confirm != d["confirm"] or token(d) != d["confirm"]:
        raise SystemExit("Confirmation token does not match the draft as shown; nothing was sent. Re-run draft and show it again.")
    return d


def cancel(d: dict, what: str):
    d["status"], d["cancelled"] = "cancelled", dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    save(d)
    raise SystemExit(f"Cancelled in the confirmation window (or it timed out); nothing was {what}. Ask the user before making a new draft.")


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


# ---------------------------------------------------------------- attachments
ATTACH_MAX_MB = 20   # what most mail servers accept in total


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def describe_attachments(paths) -> list:
    out, total = [], 0
    for raw in paths or []:
        p = os.path.abspath(os.path.expanduser(raw))
        if not os.path.isfile(p):
            raise SystemExit(f"Attachment not found: {raw}")
        size = os.path.getsize(p)
        total += size
        out.append({"Path": p, "Name": os.path.basename(p), "Size": size, "Sha256": _sha256(p)})
    if total > ATTACH_MAX_MB * 1024 * 1024:
        raise SystemExit(f"Attachments total {total / 1048576:.1f} MB, over the {ATTACH_MAX_MB} MB limit.")
    return out


def check_attachments(rows):
    """Before attaching: every file still there and unchanged since the draft was shown."""
    for a in rows:
        if not os.path.isfile(a["Path"]):
            raise SystemExit(f"Attachment missing, nothing was sent: {a['Path']}")
        if os.path.getsize(a["Path"]) != a["Size"] or _sha256(a["Path"]) != a["Sha256"]:
            raise SystemExit(f"Attachment changed since the draft was shown, nothing was sent: {a['Path']}")


def fmt_attachments(rows) -> str:
    return "; ".join(f"{a['Name']} ({a['Size'] / 1024:.0f} KB)" if a["Size"] < 1048576 else f"{a['Name']} ({a['Size'] / 1048576:.1f} MB)" for a in rows) or "(none)"


# ---------------------------------------------------------------- draft
def approver_name(ns, cfg) -> str:
    v = cfg.get("approver")
    if v:
        return str(v)
    for getter in (lambda: ns.CurrentUser.Name, lambda: ns.Accounts[0].DisplayName):
        v = oc._safe(getter, "")
        if v:
            return str(v)
    return os.environ.get("USERNAME") or os.environ.get("USER") or "the user"


def footer_text(cfg, approver: str) -> str:
    return str(cfg.get("footer") or "--\nDrafted by Claude, reviewed and approved by {approver}.").replace("\\n", "\n").format(approver=approver)


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

    approver = approver_name(ns, cfg)
    footer = footer_text(cfg, approver)
    full_body = body + "\n\n" + footer + ("\n\n" + quote if quote else "")
    d = {
        "id": new_id(), "kind": "mail",
        "status": "draft", "created": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "mode": ("reply_all" if a.reply_all else "reply") if a.reply_to else "new",
        "reply_to": reply or None,
        "to": to, "cc": cc, "subject": subject,
        "body": body, "footer": footer, "quote": quote, "full_body": full_body,
        "attachments": describe_attachments(a.attach),
        "approver": approver,
    }
    d["confirm"] = token(d)
    save(d)
    return d


# ---------------------------------------------------------------- the confirmation window
def fmt_people(rows) -> str:
    return "; ".join(f"{r['Name']} <{r['Address']}>" if r["Name"] and r["Name"] != r["Address"] else r["Address"] for r in rows) or "(none)"


def mail_dialog_spec(d: dict) -> dict:
    return {"title": "Outlook 確認寄出 / Confirm send",
            "rows": [("收件者 To", fmt_people(d["to"])), ("副本 Cc", fmt_people(d["cc"])), ("主旨 Subject", d["subject"]),
                     ("附件 Attachments", fmt_attachments(d.get("attachments") or []))],
            "text": d["full_body"],
            "note": "這封信會以上面的內容原樣寄出，寄出後無法收回。",
            "ok": "寄出 Send", "question": "寄出？ Send this mail?"}


def confirm_dialog(spec: dict, timeout_seconds: int) -> bool:
    """A window on the user's desktop showing spec["rows"] (label, value) and spec["text"], with an OK
    button (spec["ok"]) and Cancel. Returns True only when the user clicks OK. Closing the window,
    Cancel, or the timeout all return False. Tests replace this function; nothing else does."""
    if platform.system() != "Windows":
        raise SystemExit("This step needs Windows with Classic Outlook.")
    try:
        import tkinter as tk
        from tkinter import scrolledtext
    except Exception:
        return _messagebox(spec)
    result = {"ok": False}
    root = tk.Tk()
    root.title(spec["title"])
    root.attributes("-topmost", True)
    head = tk.Frame(root, padx=12, pady=8)
    head.pack(fill="x")
    for label, value in spec["rows"]:
        row = tk.Frame(head)
        row.pack(fill="x")
        tk.Label(row, text=label + ":", width=14, anchor="w", font=("", 10, "bold")).pack(side="left")
        tk.Label(row, text=value, anchor="w", justify="left", wraplength=740).pack(side="left", fill="x")
    box = scrolledtext.ScrolledText(root, width=100, height=24, wrap="word", font=("", 10))
    box.insert("1.0", spec["text"])
    box.configure(state="disabled")
    box.pack(fill="both", expand=True, padx=12)
    foot = tk.Frame(root, padx=12, pady=10)
    foot.pack(fill="x")
    tk.Label(foot, text=f"{spec['note']}{timeout_seconds // 60} 分鐘沒動作視為取消。", anchor="w").pack(side="left")

    def ok():
        result["ok"] = True
        root.destroy()

    tk.Button(foot, text="取消 Cancel", width=14, command=root.destroy).pack(side="right")
    tk.Button(foot, text=spec["ok"], width=14, command=ok, default="active").pack(side="right", padx=8)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.after(max(10, int(timeout_seconds)) * 1000, root.destroy)
    root.lift()
    root.focus_force()
    root.mainloop()
    return result["ok"]


def _messagebox(spec: dict) -> bool:
    """Fallback when tkinter is missing: a Yes/No message box (text limited, long bodies are cut)."""
    import ctypes
    text = "\n".join(f"{k}: {v}" for k, v in spec["rows"]) + "\n\n" + spec["text"]
    if len(text) > 3500:
        text = text[:3500] + "\n\n[... cut for display; the item itself is complete ...]"
    MB_YESNO, MB_ICONWARNING, MB_DEFBUTTON2, MB_SETFOREGROUND, MB_TOPMOST = 0x4, 0x30, 0x100, 0x10000, 0x40000
    r = ctypes.windll.user32.MessageBoxW(0, text + "\n\n" + spec["question"], spec["title"], MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2 | MB_SETFOREGROUND | MB_TOPMOST)
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


def _drop_inherited_attachments(mail) -> list:
    """Outlook's Reply()/ReplyAll() item starts with the original's inline pictures (cid: images, signature
    logos) as attachments; regular attachments are never carried over. They are not in the approved draft
    and the outgoing text is plain, so everything the fresh item holds is removed here, before the draft's
    own files are added. Returns the removed names. Anything Outlook refuses to remove stays on the item
    and _verify then aborts the send."""
    atts = mail.Attachments
    removed = []
    for i in range(int(oc._safe(lambda: atts.Count, 0) or 0), 0, -1):
        name = str(oc._safe(lambda: atts.Item(i).FileName, "") or "")
        try:
            atts.Remove(i)
        except Exception:
            continue
        removed.append(name)
    removed.reverse()
    return removed


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
    want = sorted(a["Name"].lower() for a in d.get("attachments") or [])
    got = sorted(str(att.FileName).lower() for att in oc._safe(lambda: list(mail.Attachments), []) or [])
    if got != want:
        extra = sorted((Counter(got) - Counter(want)).elements())
        missing = sorted((Counter(want) - Counter(got)).elements())
        msg = "Attachments differ"
        if extra:
            msg += f"; on the item but not in the approved draft: {extra}"
            if d["mode"] != "new":
                msg += " (inherited from the original mail; Outlook did not let this script remove it)"
        if missing:
            msg += f"; in the approved draft but not on the item: {missing}"
        problems.append(msg)
    if problems:
        raise SystemExit("Item does not match the approved draft, nothing was sent: " + "; ".join(problems))


def run_send(a, ns=None):
    d = load_sendable(a.id, "mail", a.confirm)
    cfg = (ps.resolve()[0].get("send") or {})
    timeout = int(cfg.get("dialog_timeout_seconds") or 300)

    # the user's click, on their own screen; there is no way around this call
    if not confirm_dialog(mail_dialog_spec(d), timeout):
        cancel(d, "sent")

    ns = ns or oc.connect()
    app = oc.application()
    inherited = []
    if d["mode"] == "new":
        mail = app.CreateItem(oc.OL_ITEM_TYPE_MAIL)
    else:
        orig = ns.GetItemFromID(d["reply_to"]["EntryID"])
        mail = orig.ReplyAll() if d["mode"] == "reply_all" else orig.Reply()
        inherited = _drop_inherited_attachments(mail)
    _set_recipients(mail, d)
    mail.Subject = d["subject"]
    try:
        mail.BodyFormat = OL_FORMAT_PLAIN
    except Exception:
        pass
    mail.Body = d["full_body"]
    atts = d.get("attachments") or []
    check_attachments(atts)
    for att in atts:
        mail.Attachments.Add(att["Path"])
    _verify(mail, d)
    mail.Send()
    d["status"], d["sent_at"] = "sent", dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    save(d)
    return {"Sent": True, "Id": d["id"], "Mode": d["mode"], "To": d["to"], "Cc": d["cc"], "Subject": d["subject"],
            "Attachments": [a["Name"] for a in atts], "InheritedAttachmentsRemoved": inherited,
            "SentAt": d["sent_at"], "Approver": d["approver"]}


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
            rows.append({k: d.get(k) for k in ("id", "kind", "status", "created", "mode", "subject", "sent_at", "start")} | {"To": [r["Address"] for r in d.get("to") or d.get("required") or []]})
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
    ap.opt("-Attach", nargs="+", default=[], help=f"files to attach (paths); recorded with size and SHA-256, re-checked before sending; {ATTACH_MAX_MB} MB total")
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
