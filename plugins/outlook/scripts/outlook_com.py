#!/usr/bin/env python3
"""Shared READ-ONLY helpers for the Outlook skills (Python + pywin32).

Every function here only reads from Outlook through COM automation. There is deliberately no
wrapper for Save, Send, Delete, Move, Copy, MarkAsRead/UnRead, Display, or any property setter.
Scripts that import this module follow the same rule, with one exception: outlook_send.py creates and
sends a mail, and only after the user clicked Send in a window on their own desktop.

Requires Classic Outlook (2013 ... Microsoft 365) on Windows and `pip install pywin32`.
"New Outlook" (olk.exe) has no COM object model and is not supported.
"""
import argparse
import datetime as dt
import json
import platform
import sys

# OlDefaultFolders
OL_FOLDER = {
    "DeletedItems": 3, "Outbox": 4, "SentMail": 5, "Inbox": 6, "Calendar": 9, "Contacts": 10,
    "Journal": 11, "Notes": 12, "Tasks": 13, "Drafts": 16, "Junk": 23,
}
OL_MAIL_ITEM = 43          # item.Class of a mail
OL_APPOINTMENT = 26
OL_ITEM_TYPE_MAIL = 0      # Application.CreateItem argument (OlItemType), used only by outlook_send.py

# MAPI property tags via PropertyAccessor (read-only lookups)
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
PR_LIST_UNSUBSCRIBE = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/List-Unsubscribe"

_KNOWN_FOLDER_NAMES = {
    "inbox": "Inbox", "收件匣": "Inbox", "收件箱": "Inbox",
    "sent items": "SentMail", "sentmail": "SentMail", "寄件備份": "SentMail", "已发送邮件": "SentMail",
    "deleted items": "DeletedItems", "deleteditems": "DeletedItems", "刪除的郵件": "DeletedItems", "已删除邮件": "DeletedItems",
    "drafts": "Drafts", "草稿": "Drafts",
    "junk email": "Junk", "junk": "Junk", "垃圾郵件": "Junk", "垃圾邮件": "Junk",
    "outbox": "Outbox", "寄件匣": "Outbox",
    "calendar": "Calendar", "行事曆": "Calendar", "日历": "Calendar",
    "contacts": "Contacts", "連絡人": "Contacts", "联系人": "Contacts",
    "tasks": "Tasks", "工作": "Tasks", "任务": "Tasks",
}

_namespace = None
_application = None


# ---------------------------------------------------------------- connection
def connect():
    """Return the MAPI namespace. Attaches to a running Outlook or starts one in the background."""
    global _namespace, _application
    if _namespace is not None:
        return _namespace
    if platform.system() != "Windows":
        raise SystemExit("Outlook COM automation needs Windows with Classic Outlook.")
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise SystemExit("pywin32 is not installed. Run:  pip install pywin32")
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        _application = win32com.client.Dispatch("Outlook.Application")  # single-instance: attaches if running
    except Exception as e:
        raise SystemExit(f"Cannot start Outlook COM automation. Classic Outlook must be installed; New Outlook (olk.exe) is not supported. Error: {e}")
    _namespace = _application.GetNamespace("MAPI")
    return _namespace


def application():
    if _application is None:
        connect()
    return _application


def set_namespace_for_tests(ns, app=None):
    """Tests inject a fake namespace here so the scripts run without Outlook."""
    global _namespace, _application
    _namespace, _application = ns, app


def reset():
    """Drop the cached COM objects so the next call reconnects. A long-lived process (the MCP server)
    calls this when a COM call fails because Outlook was closed or restarted underneath it."""
    global _namespace, _application
    _namespace, _application = None, None


# ---------------------------------------------------------------- dates
def to_datetime(value):
    """COM dates arrive as pywintypes datetimes in local time (sometimes tz-aware). Normalise to naive local."""
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.replace(tzinfo=None) if value.tzinfo else value
    try:
        return dt.datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)
    except Exception:
        return dt.datetime.fromisoformat(str(value)[:19])


def iso(value):
    d = to_datetime(value)
    return d.strftime("%Y-%m-%dT%H:%M:%S") if d else None


def parse_date(text: str) -> dt.datetime:
    """Accept 2026-09-14, 2026-09-14T10:00, 2026/9/14, 9/14/2026."""
    text = text.strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y/%m/%d %H:%M", "%Y/%m/%d", "%m/%d/%Y %H:%M", "%m/%d/%Y"):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise SystemExit(f"Cannot parse date '{text}'. Use YYYY-MM-DD or YYYY-MM-DD HH:MM.")


def jet_date(d: dt.datetime) -> str:
    """Date literal for Items.Restrict in Jet syntax. Outlook parses it with the user's locale, so build it
    from the Windows short-date pattern when available; ISO-like output is a safe fallback in most locales."""
    pattern = _windows_short_date_pattern()
    if pattern:
        return _format_with_pattern(d, pattern) + " " + d.strftime("%H:%M")
    return d.strftime("%Y-%m-%d %H:%M")


def _windows_short_date_pattern():
    if platform.system() != "Windows":
        return None
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(80)
        LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE = 0x0400, 0x1F
        if ctypes.windll.kernel32.GetLocaleInfoW(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, buf, 80):
            return buf.value
    except Exception:
        pass
    return None


def _format_with_pattern(d: dt.datetime, pattern: str) -> str:
    out, i = "", 0
    while i < len(pattern):
        c = pattern[i]
        if c in "yMd":
            j = i
            while j < len(pattern) and pattern[j] == c:
                j += 1
            n = j - i
            if c == "y":
                out += str(d.year) if n > 2 else f"{d.year % 100:02d}"
            elif c == "M":
                out += f"{d.month:02d}" if n >= 2 else str(d.month)
            else:
                out += f"{d.day:02d}" if n >= 2 else str(d.day)
            i = j
        elif c == "'":
            j = pattern.find("'", i + 1)
            out += pattern[i + 1:j] if j > i else ""
            i = (j + 1) if j > i else i + 1
        else:
            out += c
            i += 1
    return out


# ---------------------------------------------------------------- folders
def get_stores(ns=None):
    ns = ns or connect()
    return [s for s in ns.Stores]


def store_names(ns=None):
    return [str(s.DisplayName) for s in get_stores(ns)]


def find_store(name: str, ns=None):
    for s in get_stores(ns):
        if str(s.DisplayName) == name:
            return s
    raise SystemExit(f"Store '{name}' not found. Stores: {', '.join(store_names(ns))}")


def resolve_known_folder_name(name: str):
    n = name.strip()
    return _KNOWN_FOLDER_NAMES.get(n.lower()) or _KNOWN_FOLDER_NAMES.get(n)


def get_folder(path: str = "", store: str = "", ns=None):
    """Resolve "Inbox", "Inbox/Projects/Alpha", "收件匣", "\\\\Store\\Inbox\\Sub". Empty = default Inbox."""
    ns = ns or connect()
    path = (path or "").strip()
    if not path:
        if store:
            return find_store(store, ns).GetDefaultFolder(OL_FOLDER["Inbox"])
        return ns.GetDefaultFolder(OL_FOLDER["Inbox"])

    parts = [p for p in path.replace("\\", "/").split("/") if p]
    if path.startswith("\\\\") or store:
        if store:
            store_name = store
        else:
            store_name, parts = parts[0], parts[1:]
        st = find_store(store_name, ns)
        root = st.GetRootFolder()
        if parts:
            known = resolve_known_folder_name(parts[0])
            if known:
                try:
                    root = st.GetDefaultFolder(OL_FOLDER[known])
                    parts = parts[1:]
                except Exception:
                    pass
    else:
        known = resolve_known_folder_name(parts[0])
        if known:
            root = ns.GetDefaultFolder(OL_FOLDER[known])
            parts = parts[1:]
        else:
            root = ns.GetDefaultFolder(OL_FOLDER["Inbox"]).Parent

    current = root
    for p in parts:
        nxt = None
        for f in current.Folders:
            if str(f.Name).lower() == p.lower():
                nxt = f
                break
        if nxt is None:
            raise SystemExit(f"Folder '{p}' not found under '{current.FolderPath}'.")
        current = nxt
    return current


def mail_folders_recursive(folder):
    out = []
    try:
        if int(folder.DefaultItemType) == 0:
            out.append(folder)
    except Exception:
        out.append(folder)
    for f in folder.Folders:
        out.extend(mail_folders_recursive(f))
    return out


def iter_items(items):
    """Iterate a (possibly restricted/sorted) Items collection with GetFirst/GetNext."""
    item = items.GetFirst()
    while item is not None:
        yield item
        item = items.GetNext()


def iter_mail(items):
    """iter_items, mail items only (meeting requests, reports and other classes are skipped)."""
    for item in iter_items(items):
        if int(_safe(lambda: item.Class, 0)) == OL_MAIL_ITEM:
            yield item


def dasl_date(d: dt.datetime) -> str:
    """Date literal for a DASL (@SQL=) filter, ISO form (locale-independent in practice; outlook_search
    falls back to filtering dates in Python if a store rejects it)."""
    return d.strftime("%Y-%m-%d %H:%M")


def my_addresses(ns=None):
    """Lower-cased SMTP addresses that count as 'me' (all accounts + current user)."""
    ns = ns or connect()
    out = set()
    try:
        for a in ns.Accounts:
            v = str(_safe(lambda: a.SmtpAddress, "") or "").lower()
            if v:
                out.add(v)
    except Exception:
        pass
    for getter in (lambda: ns.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress, lambda: ns.CurrentUser.Address):
        v = str(_safe(getter, "") or "").lower()
        if v and "@" in v:
            out.add(v)
    return out


def is_me(address: str, me: set) -> bool:
    return (address or "").lower() in me


def appointment_attendees(appt):
    """Attendees with addresses where available: [{Name, Address, Type}] (Type 1 required, 2 optional, 3 resource)."""
    out = recipient_list(appt, 0)
    if out:
        return [{"Name": r["Name"], "Address": r["Address"]} for r in out]
    names = [n.strip() for n in (str(_safe(lambda: appt.RequiredAttendees, "") or "") + ";" + str(_safe(lambda: appt.OptionalAttendees, "") or "")).split(";") if n.strip()]
    return [{"Name": n, "Address": ""} for n in names]


# ---------------------------------------------------------------- item projections
def sender_smtp(mail) -> str:
    try:
        v = mail.PropertyAccessor.GetProperty(PR_SENDER_SMTP_ADDRESS)
        if v:
            return str(v)
    except Exception:
        pass
    try:
        if str(mail.SenderEmailType) == "EX":
            u = mail.Sender.GetExchangeUser()
            if u and u.PrimarySmtpAddress:
                return str(u.PrimarySmtpAddress)
    except Exception:
        pass
    try:
        return str(mail.SenderEmailAddress or "")
    except Exception:
        return ""


def recipient_list(item, rtype: int = 0):
    """rtype: 1 To, 2 CC, 3 BCC, 0 all."""
    out = []
    try:
        for r in item.Recipients:
            if rtype and int(r.Type) != rtype:
                continue
            addr = ""
            try:
                addr = str(r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) or "")
            except Exception:
                pass
            if not addr:
                try:
                    addr = str(r.Address or "")
                except Exception:
                    pass
            out.append({"Name": str(r.Name), "Address": addr})
    except Exception:
        pass
    return out


def attachment_list(item):
    out = []
    try:
        for a in item.Attachments:
            out.append({"FileName": str(a.FileName), "Size": int(a.Size), "Type": int(a.Type)})
    except Exception:
        pass
    return out


def _safe(getter, default=None):
    try:
        return getter()
    except Exception:
        return default


def mail_summary(mail, include_body: bool = False, preview_length: int = 200) -> dict:
    """Body is the most expensive property to read over COM; it is fetched only when a preview or the
    full body was asked for (preview_length 0 and include_body False skips it)."""
    body = (_safe(lambda: str(mail.Body), "") or "") if (include_body or preview_length > 0) else ""
    preview = (body[:preview_length] + "...") if len(body) > preview_length else body
    attachments = attachment_list(mail)
    obj = {
        "EntryID": str(mail.EntryID),
        "Folder": _safe(lambda: str(mail.Parent.FolderPath), ""),
        "ReceivedTime": iso(mail.ReceivedTime),
        "SentOn": _safe(lambda: iso(mail.SentOn)),
        "From": _safe(lambda: str(mail.SenderName), ""),
        "FromAddress": sender_smtp(mail),
        "To": _safe(lambda: str(mail.To), ""),
        "CC": _safe(lambda: str(mail.CC), ""),
        "Subject": _safe(lambda: str(mail.Subject), ""),
        "Unread": bool(_safe(lambda: mail.UnRead, False)),
        "HasAttachments": len(attachments) > 0,
        "Attachments": attachments,
        "Size": int(_safe(lambda: mail.Size, 0) or 0),
        "Importance": int(_safe(lambda: mail.Importance, 1) or 1),
        "FlagStatus": int(_safe(lambda: mail.FlagStatus, 0) or 0),
        "Categories": _safe(lambda: str(mail.Categories), ""),
        "ConversationID": _safe(lambda: str(mail.ConversationID), ""),
        "ConversationTopic": _safe(lambda: str(mail.ConversationTopic), ""),
        "BodyPreview": " ".join(preview.split()),
    }
    if include_body:
        obj["Body"] = body
    return obj


_BUSY = {0: "Free", 1: "Tentative", 2: "Busy", 3: "OutOfOffice", 4: "WorkingElsewhere"}
_MEETING = {0: "NonMeeting", 1: "Meeting", 3: "Received", 5: "Canceled", 7: "ReceivedAndCanceled"}
_RESPONSE = {0: "None", 1: "Organized", 2: "Tentative", 3: "Accepted", 4: "Declined", 5: "NotResponded"}


def appointment_summary(appt) -> dict:
    body = " ".join((_safe(lambda: str(appt.Body), "") or "").split())
    if len(body) > 300:
        body = body[:300] + "..."
    return {
        "EntryID": str(appt.EntryID),
        "Subject": _safe(lambda: str(appt.Subject), ""),
        "Start": iso(appt.Start),
        "End": iso(appt.End),
        "DurationMinutes": int(_safe(lambda: appt.Duration, 0) or 0),
        "AllDayEvent": bool(_safe(lambda: appt.AllDayEvent, False)),
        "Location": _safe(lambda: str(appt.Location), ""),
        "Organizer": _safe(lambda: str(appt.Organizer), ""),
        "RequiredAttendees": _safe(lambda: str(appt.RequiredAttendees), ""),
        "OptionalAttendees": _safe(lambda: str(appt.OptionalAttendees), ""),
        "BusyStatus": _BUSY.get(int(_safe(lambda: appt.BusyStatus, 2) or 0), "Unknown"),
        "MeetingStatus": _MEETING.get(int(_safe(lambda: appt.MeetingStatus, 0) or 0), "Unknown"),
        "ResponseStatus": _RESPONSE.get(int(_safe(lambda: appt.ResponseStatus, 0) or 0), "Unknown"),
        "IsRecurring": bool(_safe(lambda: appt.IsRecurring, False)),
        "Categories": _safe(lambda: str(appt.Categories), ""),
        "BodyPreview": body,
    }


# ---------------------------------------------------------------- output & args
def dasl_literal(value: str) -> str:
    return value.replace("'", "''")


def write_json(obj, out_file: str = ""):
    text = json.dumps(obj, ensure_ascii=False, indent=2)
    if out_file:
        with open(out_file, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"Written to {out_file}")
    else:
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
        print(text)


class ArgParser(argparse.ArgumentParser):
    """argparse that accepts PowerShell-style spellings (-From, -After, -HasAttachments) as well as --from."""

    def __init__(self, *a, **kw):
        kw.setdefault("allow_abbrev", False)
        kw.setdefault("formatter_class", argparse.RawDescriptionHelpFormatter)
        super().__init__(*a, **kw)

    def opt(self, ps_name: str, *, dest=None, **kw):
        long = "--" + ps_name.lstrip("-").lower()
        dest = dest or ps_name.lstrip("-").lower().replace("-", "_")
        return self.add_argument(ps_name, long, dest=dest, **kw)

    def flag(self, ps_name: str, *, dest=None, **kw):
        return self.opt(ps_name, dest=dest, action="store_true", **kw)


def add_common_output(ap: ArgParser):
    ap.opt("-OutFile", dest="out_file", default="", help="write JSON to this file instead of stdout")
    ap.add_argument("--out", dest="out_file", help=argparse.SUPPRESS)
