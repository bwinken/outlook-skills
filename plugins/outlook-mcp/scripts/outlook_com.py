#!/usr/bin/env python3
"""Shared READ-ONLY helpers for the Outlook skills (Python + pywin32).

Every function here only reads from Outlook through COM automation. There is deliberately no
wrapper for Save, Send, Delete, Move, Copy, MarkAsRead/UnRead, Display, or any property setter.
Scripts that import this module follow the same rule, with one exception: outlook_send.py creates and
sends a mail, and only after the user clicked Send in a window on their own desktop.

Cost model. Python and Outlook are separate processes, so every property read is a cross-process
call. A folder scan therefore goes through Outlook's Table object (scan_mail): one GetArray call
returns hundreds of rows with every summary field except the body and the attachment list, and the
item itself is opened only for the few rows that need those (enrich). The Items collection with
per-item reads remains as the fallback for stores without Table support.

Requires Classic Outlook (2013 ... Microsoft 365) on Windows and `pip install pywin32`.
"New Outlook" (olk.exe) has no COM object model and is not supported.
"""
import argparse
import contextlib
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

# MAPI property tags: read-only lookups through PropertyAccessor, and Table columns
_PROPTAG = "http://schemas.microsoft.com/mapi/proptag/"
PR_SENDER_SMTP_ADDRESS = _PROPTAG + "0x5D01001F"
PR_SMTP_ADDRESS = _PROPTAG + "0x39FE001F"
PR_MESSAGE_FLAGS = _PROPTAG + "0x0E070003"       # bit 0: read
PR_HASATTACH = _PROPTAG + "0x0E1B000B"
PR_MESSAGE_SIZE = _PROPTAG + "0x0E080003"
PR_IMPORTANCE = _PROPTAG + "0x00170003"
PR_FLAG_STATUS = _PROPTAG + "0x10900003"
PR_CONVERSATION_ID = _PROPTAG + "0x30130102"
PR_MESSAGE_TO_ME = _PROPTAG + "0x0057000B"       # the user is a To recipient (set on delivery)
PR_LIST_UNSUBSCRIBE = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/List-Unsubscribe"
DASL_RECEIVED = "urn:schemas:httpmail:datereceived"
DASL_SENT = "urn:schemas:httpmail:date"

# Table columns behind the summary fields: (field, candidate property names, only if this field is missing).
# EntryID, Subject and MessageClass are default columns of every Table and are not added.
# The first candidate Columns.Add accepts is used; a field no candidate provides gets its default.
# Date-time columns: a Table gives local time for a built-in name (ReceivedTime) but UTC for a
# namespace or proptag name, so the built-in name comes first and UTC_DATE_COLUMNS are converted.
TABLE_COLUMNS = (
    ("ReceivedTime", ("ReceivedTime", DASL_RECEIVED), None),
    ("SentOn", ("SentOn", DASL_SENT), None),
    ("From", ("urn:schemas:httpmail:fromname",), None),
    ("FromSmtp", (PR_SENDER_SMTP_ADDRESS,), None),
    ("FromEmail", ("urn:schemas:httpmail:fromemail",), None),
    ("To", ("urn:schemas:httpmail:displayto",), None),
    ("CC", ("urn:schemas:httpmail:displaycc",), None),
    ("MessageFlags", (PR_MESSAGE_FLAGS,), None),
    ("Read", ("urn:schemas:httpmail:read",), "MessageFlags"),
    ("HasAttachments", (PR_HASATTACH, "urn:schemas:httpmail:hasattachment"), None),
    ("Size", (PR_MESSAGE_SIZE,), None),
    ("Importance", (PR_IMPORTANCE, "urn:schemas:httpmail:importance"), None),
    ("FlagStatus", (PR_FLAG_STATUS,), None),
    ("Categories", ("urn:schemas-microsoft-com:office:office#Keywords",), None),
    ("ConversationID", (PR_CONVERSATION_ID,), None),
    ("ConversationTopic", ("urn:schemas:httpmail:thread-topic",), None),
)
# Optional columns a script asks for by name (scan_mail extra=...). Value None when unavailable.
EXTRA_COLUMNS = {
    "ToMe": (PR_MESSAGE_TO_ME,),
    "Unsubscribe": (PR_LIST_UNSUBSCRIBE,),
}
UTC_DATE_COLUMNS = {DASL_RECEIVED, DASL_SENT}
TABLE_BATCH = 500          # rows per GetArray call

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
_unsupported_columns = set()   # property names Columns.Add rejected in this process; not retried per folder


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
        _application = _dispatch_outlook(win32com.client)  # single-instance: attaches if running
    except Exception as e:
        raise SystemExit(f"Cannot start Outlook COM automation. Classic Outlook must be installed; New Outlook (olk.exe) is not supported. Error: {e}")
    _namespace = _application.GetNamespace("MAPI")
    return _namespace


def _dispatch_outlook(client):
    """Early-bound (makepy) Outlook when its type library cache can be built, else late-bound.
    Early binding knows every dispid up front: a property read is one call instead of a name lookup
    plus a call, and no type information is fetched for each of the thousands of objects a scan
    touches. The cache is generated once per machine (a few seconds), on demand per class, in
    pywin32's own gen_py folder; makepy prints nothing on stdout (the JSON channel) because stdout is
    redirected while it runs. Any failure falls back, down to a dispatch that ignores the cache."""
    try:
        from win32com.client import gencache
        with contextlib.redirect_stdout(sys.stderr):
            return gencache.EnsureDispatch("Outlook.Application")
    except Exception:
        pass
    try:
        return client.Dispatch("Outlook.Application")
    except Exception:
        from win32com.client import dynamic
        return dynamic.Dispatch("Outlook.Application")   # a broken gen_py cache cannot get in the way here


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
    _unsupported_columns.clear()


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
        try:
            return dt.datetime.fromisoformat(str(value)[:19])
        except Exception:
            return None


def iso(value):
    d = to_datetime(value)
    return d.strftime("%Y-%m-%dT%H:%M:%S") if d else None


def from_iso(text):
    return dt.datetime.strptime(text, "%Y-%m-%dT%H:%M:%S") if text else None


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
    """iter_items, mail items only (meeting requests, reports and other classes are skipped).
    One Class read per item; scan_mail avoids it by reading MessageClass from the Table."""
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


# ---------------------------------------------------------------- folder scans (Table object)
class MailRow:
    """One mail of a folder scan: the summary fields a Table column can give (no item opened), the times
    as datetimes for comparisons, the optional extra columns, and the item once something opened it."""
    __slots__ = ("summary", "received", "sent", "from_email", "extra", "item")

    def __init__(self, summary, received, sent, from_email="", extra=None, item=None):
        self.summary, self.received, self.sent = summary, received, sent
        self.from_email, self.extra, self.item = from_email, dict(extra or {}), item


def scan_mail(folder, dasl="", after=None, before=None, limit=None, sort="ReceivedTime", extra=(), dasl_fallback=None, folder_path=None):
    """Mails of one folder, newest first, as MailRow objects, without opening the items: one
    Table.GetArray call brings back hundreds of rows with every summary field except the body and the
    attachment list (enrich fetches those for the rows that need them). Stores without Table support
    get the Items collection with per-item reads instead.

    `after` / `before` are checked here as well as in the DASL, so a filter the store rejected can be
    retried without its date literals (`dasl_fallback`, the same filter built without dates).
    `sort` is ReceivedTime or SentOn; `extra` names optional columns (EXTRA_COLUMNS)."""
    folder_path = folder_path if folder_path is not None else str(_safe(lambda: folder.FolderPath, "") or "")
    filters = [dasl] + ([dasl_fallback] if dasl_fallback is not None and dasl_fallback != dasl else [])
    rows = None
    for i, filt in enumerate(filters):
        try:
            rows = _open_rows(folder, filt, sort, extra, folder_path)
            break
        except Exception:
            if i == len(filters) - 1:
                raise
    n = 0
    for row in rows:
        when = row.received if sort == "ReceivedTime" else (row.sent or row.received)
        if when is not None:
            if after and when < after:
                break  # sorted newest first: nothing older will match
            if before and when >= before:
                continue
        yield row
        n += 1
        if limit is not None and n >= limit:
            break


def _open_rows(folder, dasl, sort, extra, folder_path):
    try:
        return _table_rows(folder, dasl, sort, extra, folder_path)
    except Exception:
        return _item_rows(folder, dasl, sort, extra, folder_path)


def _table_rows(folder, dasl, sort, extra, folder_path):
    table = folder.GetTable(dasl) if dasl else folder.GetTable()
    cols = table.Columns
    n = int(cols.Count)
    index = {}   # field -> 0-based position in a row
    defaults = {"EntryID": 0, "Subject": 1, "MessageClass": 4}   # documented default columns, in order
    if n == 5:
        index.update(defaults)
    else:
        names = [str(cols.Item(i).Name).lower() for i in range(1, n + 1)]
        for field in defaults:
            if field.lower() in names:
                index[field] = names.index(field.lower())
    if "EntryID" not in index:
        raise RuntimeError("Table without an EntryID column")
    wanted = list(TABLE_COLUMNS) + [(f, EXTRA_COLUMNS[f], None) for f in extra if f in EXTRA_COLUMNS]
    used = {}   # field -> the property name that provides it
    for field, candidates, only_if_missing in wanted:
        if only_if_missing and only_if_missing in index:
            continue
        for name in candidates:
            if name in _unsupported_columns:
                continue
            try:
                cols.Add(name)
            except Exception:
                _unsupported_columns.add(name)
                continue
            index[field], used[field] = n, name
            n += 1
            break
    utc_fields = {f for f, name in used.items() if name in UTC_DATE_COLUMNS}
    table.Sort(used.get(sort) or (DASL_RECEIVED if sort == "ReceivedTime" else DASL_SENT), True)
    try:
        table.MoveToStart()
    except Exception:
        pass

    def rows():
        for values in _table_values(table, n):
            row = _row_from_values(values, index, folder_path, extra, utc_fields)
            if row is not None:
                yield row
    return rows()


def _table_values(table, ncols):
    """Row tuples of a Table, hundreds per call; one row per call if the store rejects GetArray."""
    try:
        while not bool(table.EndOfTable):
            arr = table.GetArray(TABLE_BATCH)
            if not arr:
                return
            arr = _as_rows(arr, ncols)
            for values in arr:
                yield values
            if len(arr) < TABLE_BATCH:
                return
    except Exception:
        pass
    while not bool(table.EndOfTable):
        row = table.GetNextRow()
        if row is None:
            return
        yield tuple(row.GetValues())


def _as_rows(arr, ncols):
    """GetArray comes back as a tuple of row tuples; a transposed array is put back into rows."""
    if not isinstance(arr, (tuple, list)) or not arr:
        return []
    first = arr[0]
    if not isinstance(first, (tuple, list)):
        return [tuple(arr)]
    if len(first) != ncols and len(arr) == ncols:
        return [tuple(r) for r in zip(*arr)]
    return [tuple(r) for r in arr]


def _row_from_values(values, index, folder_path, extra, utc_fields=()):
    def get(field):
        i = index.get(field)
        return values[i] if i is not None and i < len(values) else None

    def when(field):
        d = to_datetime(get(field))
        return _utc_to_local(d) if field in utc_fields else d

    cls = _text(get("MessageClass"))
    if cls and not cls.upper().startswith("IPM.NOTE"):
        return None   # meeting requests, reports, posts: not mail (item.Class 43)
    received = when("ReceivedTime")
    sent = when("SentOn")
    smtp, email = _text(get("FromSmtp")), _text(get("FromEmail"))
    flags, read = get("MessageFlags"), get("Read")
    if flags is not None:
        unread = not (int(flags) & 1)
    elif read is not None:
        unread = not bool(read)
    else:
        unread = False
    attachments = get("HasAttachments")
    summary = {
        "EntryID": _hex(get("EntryID")),
        "Folder": folder_path,
        "ReceivedTime": iso(received),
        "SentOn": iso(sent),
        "From": _text(get("From")),
        "FromAddress": smtp or (email if "@" in email else ""),
        "To": _text(get("To")),
        "CC": _text(get("CC")),
        "Subject": _text(get("Subject")),
        "Unread": bool(unread),
        "HasAttachments": bool(attachments) if attachments is not None else False,
        "Attachments": [],
        "Size": _int(get("Size"), 0),
        "Importance": _int(get("Importance"), 1),
        "FlagStatus": _int(get("FlagStatus"), 0),
        "Categories": _text(get("Categories")),
        "ConversationID": _hex(get("ConversationID")),
        "ConversationTopic": _text(get("ConversationTopic")),
        "BodyPreview": "",
    }
    return MailRow(summary, received, sent, email, {f: get(f) for f in extra})


def _item_rows(folder, dasl, sort, extra, folder_path):
    """Fallback for a store without Table support: the Items collection, one item at a time."""
    items = folder.Items
    if dasl:
        items = items.Restrict(dasl)
    items.Sort("[ReceivedTime]" if sort == "ReceivedTime" else "[SentOn]", True)

    def rows():
        for item in iter_mail(items):
            s = mail_summary(item, False, 0, folder_path=folder_path)
            ex = {}
            for f in extra:
                ex[f] = _safe(lambda: item.PropertyAccessor.GetProperty(PR_LIST_UNSUBSCRIBE)) if f == "Unsubscribe" else None
            yield MailRow(s, from_iso(s["ReceivedTime"]), from_iso(s["SentOn"]), _safe(lambda: str(item.SenderEmailAddress or ""), ""), ex, item=item)
    return rows()


def open_item(row, ns=None):
    """The MailItem behind a row, opened once (GetItemFromID) and kept on the row."""
    if row.item is None:
        row.item = (ns or connect()).GetItemFromID(row.summary["EntryID"])
    return row.item


def enrich(row, body=False, preview_length=0, attachments=True, ns=None):
    """Complete a row's summary with what only the open item has: the SMTP sender when the table had
    none, the attachment list (mails that have attachments), a body preview, the body. The item is
    opened only when one of those is needed; the summary is returned either way."""
    s = row.summary
    need_body = body or preview_length > 0
    need_att = attachments and s["HasAttachments"] and not s["Attachments"]
    need_from = not s["FromAddress"] and bool(row.from_email)
    if not (need_body or need_att or need_from):
        return s
    item = open_item(row, ns)
    if need_from:
        s["FromAddress"] = sender_smtp(item)
    if need_att:
        s["Attachments"] = attachment_list(item)
        s["HasAttachments"] = len(s["Attachments"]) > 0
    if need_body:
        text = _safe(lambda: str(item.Body), "") or ""
        s["BodyPreview"] = _preview(text, preview_length)
        if body:
            s["Body"] = text
    return s


def to_me(row, me, ns=None, open_if_needed=False) -> bool:
    """I am a To recipient of this row: the delivery flag column when the store has it, else the To
    display string, else the recipients of the item (only if it is open already, or open_if_needed)."""
    flag = row.extra.get("ToMe")
    if flag is not None:
        return bool(flag)
    to = row.summary["To"].lower()
    if any(addr in to for addr in me):
        return True
    if row.item is None and not open_if_needed:
        return False
    return any(is_me(r["Address"], me) for r in recipient_list(open_item(row, ns), 1))


def sender_address(row, ns=None) -> str:
    """Lower-cased sender address of a row; an Exchange (non-SMTP) address is resolved through the item once."""
    s = row.summary
    if not s["FromAddress"] and row.from_email:
        s["FromAddress"] = sender_smtp(open_item(row, ns))
    return s["FromAddress"].lower()


def _utc_to_local(d):
    return d.replace(tzinfo=dt.timezone.utc).astimezone().replace(tzinfo=None) if d else None


def _preview(text: str, length: int) -> str:
    if length <= 0:
        return ""
    preview = (text[:length] + "...") if len(text) > length else text
    return " ".join(preview.split())


def _text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (tuple, list)):
        return ", ".join(str(x) for x in v)
    return str(v)


def _int(v, default: int) -> int:
    try:
        return int(v) if v is not None else default
    except (TypeError, ValueError):
        return default


def _hex(v) -> str:
    """Binary MAPI values (EntryID, ConversationID) as the upper-case hex string the object model shows."""
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    if isinstance(v, (bytes, bytearray, memoryview)):
        return bytes(v).hex().upper()
    if isinstance(v, (tuple, list)):
        try:
            return bytes(v).hex().upper()
        except (TypeError, ValueError):
            return _text(v)
    return str(v)


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
    return [{"Name": r["Name"], "Address": r["Address"]} for r in _recipients(item) if not rtype or r["Type"] == rtype]


def recipients_by_type(item):
    """{1: [...To], 2: [...CC], 3: [...BCC]} from one pass over the Recipients collection."""
    out = {1: [], 2: [], 3: []}
    for r in _recipients(item):
        out.setdefault(r["Type"], []).append({"Name": r["Name"], "Address": r["Address"]})
    return out


def _recipients(item):
    out = []
    try:
        for r in item.Recipients:
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
            out.append({"Name": str(r.Name), "Address": addr, "Type": int(_safe(lambda: r.Type, 1) or 1)})
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


def mail_summary(mail, include_body: bool = False, preview_length: int = 0, attachments: bool = True, folder_path=None) -> dict:
    """Summary of one open MailItem, every field a property read. Body is the most expensive property
    and is fetched only for a preview (preview_length > 0) or the full body; pass folder_path when
    the folder is known so it is not read back from the item. Scans use scan_mail + enrich instead."""
    body = (_safe(lambda: str(mail.Body), "") or "") if (include_body or preview_length > 0) else ""
    atts = attachment_list(mail) if attachments else []
    obj = {
        "EntryID": str(mail.EntryID),
        "Folder": folder_path if folder_path is not None else _safe(lambda: str(mail.Parent.FolderPath), ""),
        "ReceivedTime": iso(mail.ReceivedTime),
        "SentOn": _safe(lambda: iso(mail.SentOn)),
        "From": _safe(lambda: str(mail.SenderName), ""),
        "FromAddress": sender_smtp(mail),
        "To": _safe(lambda: str(mail.To), ""),
        "CC": _safe(lambda: str(mail.CC), ""),
        "Subject": _safe(lambda: str(mail.Subject), ""),
        "Unread": bool(_safe(lambda: mail.UnRead, False)),
        "HasAttachments": len(atts) > 0,
        "Attachments": atts,
        "Size": int(_safe(lambda: mail.Size, 0) or 0),
        "Importance": int(_safe(lambda: mail.Importance, 1) or 1),
        "FlagStatus": int(_safe(lambda: mail.FlagStatus, 0) or 0),
        "Categories": _safe(lambda: str(mail.Categories), ""),
        "ConversationID": _safe(lambda: str(mail.ConversationID), ""),
        "ConversationTopic": _safe(lambda: str(mail.ConversationTopic), ""),
        "BodyPreview": _preview(body, preview_length),
    }
    if include_body:
        obj["Body"] = body
    return obj


_BUSY = {0: "Free", 1: "Tentative", 2: "Busy", 3: "OutOfOffice", 4: "WorkingElsewhere"}
_MEETING = {0: "NonMeeting", 1: "Meeting", 3: "Received", 5: "Canceled", 7: "ReceivedAndCanceled"}
_RESPONSE = {0: "None", 1: "Organized", 2: "Tentative", 3: "Accepted", 4: "Declined", 5: "NotResponded"}


def appointment_summary(appt, preview_length: int = 0) -> dict:
    """Body (the expensive property) is read only when a preview was asked for."""
    body = ""
    if preview_length > 0:
        body = " ".join((_safe(lambda: str(appt.Body), "") or "").split())
        if len(body) > preview_length:
            body = body[:preview_length] + "..."
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


def apply_settings(a):
    """Command-line entry points: defaults from .outlook-skills/settings.json for options the command
    line left empty, so a skill can run its first script without waiting for `settings.py show`.
    `store` when neither -Store nor -AllStores was given (the rule the MCP server applies as well);
    for the search scripts also search.default_folder and search.all_folders. Returns what was applied."""
    try:
        import settings as ps
        merged = ps.resolve()[0]
    except Exception:
        return []
    applied = []
    if hasattr(a, "store") and not a.store and not getattr(a, "allstores", False) and merged.get("store"):
        a.store = str(merged["store"])
        applied.append("store")
    search = merged.get("search") or {}
    if hasattr(a, "folder") and hasattr(a, "allfolders") and not getattr(a, "entryid", ""):
        default_folder = str(search.get("default_folder") or "")
        if not a.folder and default_folder and default_folder.strip().lower() != "inbox":
            a.folder = default_folder
            applied.append("search.default_folder")
        if not a.allfolders and search.get("all_folders") is True:
            a.allfolders = True
            applied.append("search.all_folders")
    return applied
