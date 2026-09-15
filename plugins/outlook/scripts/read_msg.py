#!/usr/bin/env python3
"""READ-ONLY parser for Outlook .msg and standard .eml files.

Opens files in read mode only. Never modifies the input file and never touches Outlook.
Attachments are only *listed* unless --extract-to is given, in which case copies are
written to that directory (the source file is still untouched).

Usage:
    python read_msg.py message.msg
    python read_msg.py a.eml b.msg --format markdown
    python read_msg.py message.msg --headers --max-body 5000
    python read_msg.py message.msg --extract-to ./out

Dependencies: none beyond the Python standard library (.msg is read by msgfile.py next to this file).
"""
import argparse
import email
import html
import json
import os
import re
import sys
from email import policy
from email.utils import getaddresses, parsedate_to_datetime


def _html_to_text(markup: str) -> str:
    markup = re.sub(r"(?is)<(script|style).*?</\1>", "", markup)
    markup = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>|</li>", "\n", markup)
    text = re.sub(r"<[^>]+>", "", markup)
    text = html.unescape(text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _addr_list(value):
    if not value:
        return []
    if isinstance(value, str):
        pairs = getaddresses([value])
    else:
        pairs = getaddresses([str(v) for v in value])
    return [{"name": n, "address": a} for n, a in pairs if n or a]


# ---------------------------------------------------------------- .eml
def parse_eml(path: str, want_headers: bool, extract_to):
    with open(path, "rb") as fh:
        msg = email.message_from_binary_file(fh, policy=policy.default)

    body_text, body_html = "", ""
    attachments = []
    for part in msg.walk():
        if part.is_multipart():
            continue
        disp = part.get_content_disposition()
        ctype = part.get_content_type()
        filename = part.get_filename()
        if disp == "attachment" or (filename and disp != "inline") or (filename and not ctype.startswith("text/")):
            payload = part.get_payload(decode=True) or b""
            entry = {"name": filename or "(unnamed)", "size": len(payload), "mime": ctype}
            if extract_to:
                entry["saved_to"] = _save(extract_to, entry["name"], payload)
            attachments.append(entry)
            continue
        if ctype == "text/plain" and not body_text:
            body_text = part.get_content()
        elif ctype == "text/html" and not body_html:
            body_html = part.get_content()

    date = msg.get("Date")
    try:
        date_iso = parsedate_to_datetime(date).isoformat() if date else None
    except Exception:
        date_iso = None

    result = {
        "file": os.path.abspath(path),
        "format": "eml",
        "subject": msg.get("Subject", ""),
        "from": _addr_list(msg.get("From")),
        "to": _addr_list(msg.get_all("To")),
        "cc": _addr_list(msg.get_all("Cc")),
        "bcc": _addr_list(msg.get_all("Bcc")),
        "date": date_iso or date,
        "message_id": msg.get("Message-ID"),
        "in_reply_to": msg.get("In-Reply-To"),
        "references": msg.get("References"),
        "body": body_text or (_html_to_text(body_html) if body_html else ""),
        "body_source": "text" if body_text else ("html" if body_html else "none"),
        "attachments": attachments,
    }
    if want_headers:
        result["headers"] = [{"name": k, "value": str(v)} for k, v in msg.items()]
    return result


# ---------------------------------------------------------------- .msg
def parse_msg(path: str, want_headers: bool, extract_to):
    from msgfile import MsgFile  # standard-library OLE2 + MAPI reader next to this file

    m = MsgFile(path)
    attachments = []
    for att in m.attachments:
        entry = {"name": att.filename, "size": att.size, "mime": att.mime}
        if att.method == 5:
            entry["embedded_message"] = True
        if extract_to and att.data is not None:
            entry["saved_to"] = _save(extract_to, att.filename, att.data)
        attachments.append(entry)

    body, source = m.body, "text"
    if not body and m.html_body:
        body, source = _html_to_text(m.html_body), "html"

    # Prefer the Date: transport header (what the user sees in Outlook); fall back to the MAPI submit time (UTC).
    date_iso = None
    hdr = email.message_from_string(m.headers, policy=policy.default) if m.headers else None
    if hdr is not None and hdr.get("Date"):
        try:
            date_iso = parsedate_to_datetime(hdr.get("Date")).isoformat()
        except Exception:
            date_iso = None
    if not date_iso and m.date:
        date_iso = m.date.isoformat()

    def addr_list(rtype, fallback):
        rs = m.recipients_of(rtype)
        if rs:
            return [{"name": r["name"] if r["name"] != r["address"] else "", "address": r["address"]} for r in rs]
        return _addr_list(fallback) if fallback else []

    result = {
        "file": os.path.abspath(path),
        "format": "msg",
        "subject": m.subject,
        "from": [{"name": m.sender_name, "address": m.sender_email}] if (m.sender_name or m.sender_email) else [],
        "to": addr_list("to", m.to),
        "cc": addr_list("cc", m.cc),
        "bcc": addr_list("bcc", m.bcc),
        "date": date_iso,
        "message_id": m.message_id or (hdr.get("Message-ID") if hdr is not None else None),
        "in_reply_to": m.in_reply_to or (hdr.get("In-Reply-To") if hdr is not None else None),
        "references": m.references or (hdr.get("References") if hdr is not None else None),
        "body": body,
        "body_source": source if body else "none",
        "attachments": attachments,
    }
    if want_headers:
        result["headers"] = [{"name": k, "value": str(v)} for k, v in hdr.items()] if hdr is not None else []
    return result


# ---------------------------------------------------------------- helpers
def _save(directory: str, name: str, data: bytes) -> str:
    os.makedirs(directory, exist_ok=True)
    safe = re.sub(r"[\\/:*?\"<>|]", "_", name) or "attachment"
    target = os.path.join(directory, safe)
    base, ext = os.path.splitext(target)
    n = 1
    while os.path.exists(target):
        target = f"{base}({n}){ext}"
        n += 1
    with open(target, "wb") as fh:
        fh.write(data)
    return os.path.abspath(target)


def _fmt_addrs(items):
    return ", ".join(f"{i['name']} <{i['address']}>" if i["name"] else i["address"] for i in items) or "-"


def to_markdown(r: dict) -> str:
    lines = [
        f"# {r['subject'] or '(no subject)'}",
        "",
        f"- **File**: {r['file']}",
        f"- **From**: {_fmt_addrs(r['from'])}",
        f"- **To**: {_fmt_addrs(r['to'])}",
    ]
    if r["cc"]:
        lines.append(f"- **CC**: {_fmt_addrs(r['cc'])}")
    lines.append(f"- **Date**: {r['date'] or '-'}")
    if r.get("message_id"):
        lines.append(f"- **Message-ID**: {r['message_id']}")
    if r["attachments"]:
        lines.append("- **Attachments**:")
        for a in r["attachments"]:
            size = f" ({a['size']} bytes)" if a.get("size") is not None else ""
            saved = f" -> {a['saved_to']}" if a.get("saved_to") else ""
            lines.append(f"  - {a['name']}{size}{saved}")
    if r.get("headers"):
        lines += ["", "## Headers", "", "```"]
        lines += [f"{h['name']}: {h['value']}" for h in r["headers"]]
        lines.append("```")
    lines += ["", "## Body", "", r["body"] or "(empty)"]
    return "\n".join(lines)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+", help=".msg or .eml files")
    ap.add_argument("--format", choices=["json", "markdown"], default="json")
    ap.add_argument("--headers", action="store_true", help="include full transport headers")
    ap.add_argument("--max-body", type=int, default=0, help="truncate body to N chars (0 = no limit)")
    ap.add_argument("--extract-to", metavar="DIR", help="copy attachments into DIR (source file untouched)")
    args = ap.parse_args(argv)

    results = []
    for path in args.files:
        if not os.path.isfile(path):
            raise SystemExit(f"File not found: {path}")
        ext = os.path.splitext(path)[1].lower()
        if ext == ".msg":
            r = parse_msg(path, args.headers, args.extract_to)
        elif ext == ".eml":
            r = parse_eml(path, args.headers, args.extract_to)
        else:
            raise SystemExit(f"Unsupported extension '{ext}': {path} (expected .msg or .eml)")
        if args.max_body and len(r["body"]) > args.max_body:
            r["body"] = r["body"][: args.max_body] + "\n[... truncated ...]"
            r["body_truncated"] = True
        results.append(r)

    if args.format == "markdown":
        print("\n\n---\n\n".join(to_markdown(r) for r in results))
    else:
        print(json.dumps(results if len(results) > 1 else results[0], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
