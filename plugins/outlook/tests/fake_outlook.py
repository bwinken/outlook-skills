"""Minimal fake of the Outlook COM object model, enough to exercise the scripts without Windows.

Folders answer GetTable like Outlook's Table object (default columns, Columns.Add, Sort, GetArray in
batches); TABLES_SUPPORTED = False makes GetTable fail so the scripts take their Items fallback."""
import datetime as dt

TABLES_SUPPORTED = True

_DEFAULT_COLUMNS = ["EntryID", "Subject", "CreationTime", "LastModificationTime", "MessageClass"]
_PT = "http://schemas.microsoft.com/mapi/proptag/"
_PR_LIST_UNSUBSCRIBE = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/List-Unsubscribe"
_REJECTED_COLUMNS = {_PT + "0x0E1B000B"}   # a store that does not take PR_HASATTACH: the next candidate column must be tried


class Attachment:
    def __init__(self, name, size=100, atype=1, data=b"data", inline=False):
        self.FileName, self.Size, self.Type, self._data = name, size, atype, data
        # inline: a cid: picture of an HTML mail, hidden and with a content id, as Outlook stores them
        self._inline = inline
        self.PropertyAccessor = PropertyAccessor({"http://schemas.microsoft.com/mapi/proptag/0x7FFE000B": True,
                                                  "http://schemas.microsoft.com/mapi/proptag/0x3712001F": f"{name}@01D0"} if inline else {})

    def SaveAsFile(self, path):
        with open(path, "wb") as fh:
            fh.write(self._data)


class Attachments(list):
    def Item(self, i):
        return self[i - 1]


class PropertyAccessor:
    def __init__(self, props):
        self.props = props

    def GetProperty(self, tag):
        if tag in self.props:
            return self.props[tag]
        raise Exception("property not found")


class Recipient:
    def __init__(self, name, address, rtype=1):
        self.Name, self.Address, self.Type = name, address, rtype
        self.PropertyAccessor = PropertyAccessor({"http://schemas.microsoft.com/mapi/proptag/0x39FE001F": address})
        self.Resolved = bool(address)

    def Resolve(self):
        return self.Resolved


class Recipients(list):
    """Recipients collection of an outgoing item: Add / Remove / Count / ResolveAll, 1-based Remove."""

    def __init__(self, directory):
        super().__init__()
        self._dir = directory

    @property
    def Count(self):
        return len(self)

    def Add(self, text):
        name, addr = self._dir.lookup(text)
        r = Recipient(name or text, addr or "", 1)
        self.append(r)
        return r

    def Remove(self, index):
        del self[index - 1]

    def ResolveAll(self):
        return all(r.Resolved for r in self)


class Directory:
    """Address book: display names seen anywhere in the fixture, plus any plain SMTP address."""

    def __init__(self):
        self.by_name, self.by_addr = {}, {}

    def learn(self, name, address):
        if name and address:
            self.by_name.setdefault(name.lower(), (name, address))
            self.by_addr.setdefault(address.lower(), (name, address))

    def lookup(self, text):
        t = (text or "").strip()
        if t.lower() in self.by_addr:
            return self.by_addr[t.lower()]
        if t.lower() in self.by_name:
            return self.by_name[t.lower()]
        if "@" in t:
            return t, t
        return None, None


class OutgoingAttachments(list):
    """Attachments of an outgoing item: Add(path), Count, Item(i) and Remove(i), 1-based like COM."""

    @property
    def Count(self):
        return len(self)

    def Item(self, i):
        return self[i - 1]

    def Add(self, path):
        import os
        a = Attachment(os.path.basename(path), os.path.getsize(path), data=open(path, "rb").read())
        a.PathName = path
        self.append(a)
        return a

    def Remove(self, index):
        del self[index - 1]


class Outgoing:
    """A MailItem made by CreateItem / Reply / ReplyAll. Send() records it on the application."""
    Class = 43

    def __init__(self, app, directory, subject="", body="", recipients=()):
        self._app = app
        self.Recipients = Recipients(directory)
        self.Recipients.extend(recipients)
        self.Attachments = OutgoingAttachments()
        self.Subject, self.Body, self.BodyFormat = subject, body, 2
        self.Sent, self.Saved = False, False

    @property
    def To(self):
        return "; ".join(r.Name for r in self.Recipients if r.Type == 1)

    @property
    def CC(self):
        return "; ".join(r.Name for r in self.Recipients if r.Type == 2)

    def Send(self):
        self.Sent = True
        self._app.sent.append(self)

    def Save(self):
        self.Saved = True


class Mail:
    Class = 43
    MessageClass = "IPM.Note"

    def __init__(self, entry_id, subject, sender, addr, received, body="", folder=None, to="", cc="", unread=False,
                 attachments=(), conv_topic=None, conv_id="", recipients=(), unsubscribe=False, importance=1):
        self.EntryID, self.Subject, self.SenderName, self.SenderEmailAddress = entry_id, subject, sender, addr
        self.SenderEmailType = "SMTP"
        self.ReceivedTime = received
        self.SentOn = received
        self.Body, self.To, self.CC, self.UnRead = body, to, cc, unread
        self.Attachments = Attachments(attachments)
        self.Size, self.Importance, self.FlagStatus, self.Categories = 1024, importance, 0, ""
        self.ConversationTopic = conv_topic if conv_topic is not None else subject.replace("Re: ", "").replace("RE: ", "")
        self.ConversationID = conv_id
        self.Recipients = list(recipients)
        self.Parent = folder
        props = {"http://schemas.microsoft.com/mapi/proptag/0x5D01001F": addr}
        if unsubscribe:
            props["http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/List-Unsubscribe"] = "<mailto:x>"
        self.PropertyAccessor = PropertyAccessor(props)
        self._conversation = None

    def GetConversation(self):
        return self._conversation

    def _reply_item(self, recipients):
        app = self.Parent.Store._app
        out = Outgoing(app, app.directory, "RE: " + self.Subject, "\n-----Original Message-----\n" + self.Body, recipients)
        # as in Outlook: the quoted original brings its inline pictures along as attachments, never its regular files
        out.Attachments.extend(Attachment(a.FileName, a.Size, a.Type, a._data, inline=True) for a in self.Attachments if a._inline)
        return out

    def Reply(self):
        return self._reply_item([Recipient(self.SenderName, self.SenderEmailAddress, 1)])

    def ReplyAll(self):
        others = [Recipient(r.Name, r.Address, r.Type) for r in self.Recipients if r.Address.lower() != "me@contoso.com"]
        return self._reply_item([Recipient(self.SenderName, self.SenderEmailAddress, 1)] + others)


class MeetingRequest(Mail):
    """A meeting invitation in a mail folder: not a MailItem (Class 53), so the scripts must skip it."""
    Class = 53
    MessageClass = "IPM.Schedule.Meeting.Request"


class Appointment:
    Class = 26
    MessageClass = "IPM.Appointment"

    def __init__(self, entry_id, subject, start, end, busy=2, response=3, meeting=3, recurring=False, allday=False, organizer="", location=""):
        self.EntryID, self.Subject, self.Start, self.End = entry_id, subject, start, end
        self.Duration = int((end - start).total_seconds() // 60)
        self.AllDayEvent, self.Location, self.Organizer = allday, location, organizer
        self.RequiredAttendees, self.OptionalAttendees = "", ""
        self.BusyStatus, self.MeetingStatus, self.ResponseStatus, self.IsRecurring = busy, meeting, response, recurring
        self.Categories, self.Body = "", ""
        self.Recipients = []


class Items:
    """Supports Restrict (subset of DASL/Jet we emit), Sort, IncludeRecurrences, GetFirst/GetNext, Count."""

    def __init__(self, items):
        self._items = list(items)
        self._i = 0
        self.IncludeRecurrences = False

    @property
    def Count(self):
        return len(self._items)

    def Sort(self, prop, desc=False):
        key = prop.strip("[]")
        attr = {"ReceivedTime": "ReceivedTime", "Start": "Start", "SentOn": "SentOn"}[key]
        self._items.sort(key=lambda x: getattr(x, attr), reverse=bool(desc))

    def Restrict(self, filt):
        return Items([it for it in self._items if _match(it, filt)])

    def GetFirst(self):
        self._i = 0
        return self.GetNext()

    def GetNext(self):
        if self._i < len(self._items):
            it = self._items[self._i]
            self._i += 1
            return it
        return None


def _match(it, filt):
    f = filt
    if f.startswith("@SQL="):
        f = f[5:]
        import re
        # evaluate each LIKE/= clause and combine with the boolean structure via Python eval on a rewritten string
        def repl_like(m):
            field, val = m.group(1), m.group(2).replace("''", "'").strip("%").lower()
            return "True" if val in (_field(it, field) or "").lower() else "False"
        def repl_eq(m):
            field, val = m.group(1), m.group(2).strip("'")
            return "True" if str(_field(it, field)) == val else "False"
        def repl_date(m):
            field, op, val = m.group(1), m.group(2), _jet(m.group(3))
            have = _field(it, field)
            return str({">=": have >= val, "<": have < val, ">": have > val, "<=": have <= val}[op])
        g = re.sub(r'"([^"]+)"\s+LIKE\s+\'([^\']*(?:\'\'[^\']*)*)\'', repl_like, f)
        g = re.sub(r'"(urn:schemas:httpmail:datereceived)"\s*(>=|<=|<|>)\s*\'([^\']*)\'', repl_date, g)
        g = re.sub(r'"([^"]+)"\s*=\s*(\'[^\']*\'|\d+)', repl_eq, g)
        g = g.replace(" AND ", " and ").replace(" OR ", " or ")
        return eval(g)
    # Jet: [Start] < 'x' AND [End] > 'y'
    import re
    m = re.match(r"\[Start\] < '(.+?)' AND \[End\] > '(.+?)'", f)
    end, start = _jet(m.group(1)), _jet(m.group(2))
    return it.Start < end and it.End > start


def _jet(s):
    for fmt in ("%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M", "%m/%d/%Y %H:%M"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            pass
    raise ValueError(s)


def _field(it, field):
    return {
        "urn:schemas:httpmail:fromname": it.SenderName, "urn:schemas:httpmail:fromemail": it.SenderEmailAddress,
        "urn:schemas:httpmail:displayto": it.To, "urn:schemas:httpmail:displaycc": it.CC,
        "urn:schemas:httpmail:subject": it.Subject, "urn:schemas:httpmail:textdescription": it.Body,
        "urn:schemas:httpmail:hasattachment": 1 if it.Attachments else 0, "urn:schemas:httpmail:read": 0 if it.UnRead else 1,
        "urn:schemas:httpmail:thread-topic": it.ConversationTopic,
        "urn:schemas:httpmail:datereceived": it.ReceivedTime,
        "urn:schemas:httpmail:importance": it.Importance,
        "http://schemas.microsoft.com/mapi/proptag/0x10900003": it.FlagStatus,
    }[field]


class Folder:
    def __init__(self, name, path, items=(), subfolders=(), item_type=0, store=None):
        self.Name, self.FolderPath, self.DefaultItemType, self.Store = name, path, item_type, store
        self._items = list(items)
        for it in self._items:
            it.Parent = self
        self.Folders = list(subfolders)
        self.Parent = None
        for f in self.Folders:
            f.Parent = self

    @property
    def Items(self):
        return Items(self._items)

    def GetTable(self, filt=""):
        if not TABLES_SUPPORTED:
            raise Exception("GetTable is not supported by this store")
        return FolderTable(self._items, filt)

    @property
    def UnReadItemCount(self):
        return sum(1 for m in self._items if getattr(m, "UnRead", False))


class Store:
    def __init__(self, name, root, defaults, file_path="", store_type=3):
        self.DisplayName, self._root, self._defaults = name, root, defaults
        self.FilePath, self.ExchangeStoreType, self.IsDataFileStore, self.IsCachedExchange = file_path, store_type, store_type == 3, store_type == 0
        for f in _walk(root):
            f.Store = self

    def GetRootFolder(self):
        return self._root

    def GetDefaultFolder(self, fid):
        if fid in self._defaults:
            return self._defaults[fid]
        raise Exception("no such default folder")


def _walk(f):
    yield f
    for s in f.Folders:
        yield from _walk(s)


def _column_value(it, name):
    """What Outlook's Table returns for a column: the values the object model shows for the item."""
    if name == "EntryID":
        return it.EntryID
    if name == "Subject":
        return it.Subject
    if name in ("CreationTime", "LastModificationTime", "ReceivedTime"):
        return it.ReceivedTime
    if name == "SentOn":
        return it.SentOn
    if name == "urn:schemas:httpmail:datereceived":
        return _utc(it.ReceivedTime)   # a namespace date column comes back in UTC
    if name == "MessageClass":
        return it.MessageClass
    if name == _PT + "0x5D01001F":
        return it.PropertyAccessor.props.get(name)
    if name == "urn:schemas:httpmail:date":
        return _utc(it.SentOn)
    if name == _PT + "0x0E070003":
        return 0 if it.UnRead else 1
    if name == _PT + "0x0E080003":
        return it.Size
    if name == _PT + "0x00170003":
        return it.Importance
    if name == "urn:schemas-microsoft-com:office:office#Keywords":
        return it.Categories
    if name == _PT + "0x30130102":
        return it.ConversationID
    if name == _PT + "0x0057000B":
        return any(r.Type == 1 and r.Address.lower() == "me@contoso.com" for r in it.Recipients)
    if name == _PR_LIST_UNSUBSCRIBE:
        return it.PropertyAccessor.props.get(name)
    return _field(it, name)


def _utc(local):
    return local.astimezone(dt.timezone.utc).replace(tzinfo=None) if local else None


class Column:
    def __init__(self, name):
        self.Name = name


class Columns:
    def __init__(self, names):
        self._cols = [Column(n) for n in names]

    @property
    def Count(self):
        return len(self._cols)

    def Item(self, i):
        return self._cols[i - 1]

    def Add(self, name):
        if name in _REJECTED_COLUMNS:
            raise Exception(f'The property "{name}" is unknown or cannot be found.')
        self._cols.append(Column(name))
        return self._cols[-1]

    def RemoveAll(self):
        self._cols = []

    def names(self):
        return [c.Name for c in self._cols]


class Row:
    def __init__(self, names, values):
        self._names, self._values = names, values

    def Item(self, key):
        return self._values[key - 1] if isinstance(key, int) else self._values[self._names.index(key)]

    def GetValues(self):
        return tuple(self._values)


class FolderTable:
    """Folder.GetTable(filter): rows of column values, GetArray up to n rows per call, as Outlook does."""

    def __init__(self, items, filt=""):
        self._items = [it for it in items if not filt or _match(it, filt)]
        self.Columns = Columns(_DEFAULT_COLUMNS)
        self._i = 0

    def Sort(self, prop, desc=False):
        attr = {"urn:schemas:httpmail:datereceived": "ReceivedTime", "urn:schemas:httpmail:date": "SentOn", "ReceivedTime": "ReceivedTime", "SentOn": "SentOn"}[prop]
        self._items.sort(key=lambda x: getattr(x, attr), reverse=bool(desc))
        self._i = 0

    def MoveToStart(self):
        self._i = 0

    @property
    def EndOfTable(self):
        return self._i >= len(self._items)

    def GetRowCount(self):
        return len(self._items)

    def _values(self, it):
        return tuple(_column_value(it, c.Name) for c in self.Columns._cols)

    def GetArray(self, n):
        rows = tuple(self._values(it) for it in self._items[self._i:self._i + n])
        self._i += len(rows)
        return rows

    def GetNextRow(self):
        it = self._items[self._i]
        self._i += 1
        return Row(self.Columns.names(), self._values(it))


class ConversationTable:
    """Conversation.GetTable(): one row per item of the conversation, default columns only."""

    def __init__(self, ids, lookup):
        self._ids, self._lookup, self._i = ids, lookup, 0
        self.Columns = Columns(_DEFAULT_COLUMNS)

    @property
    def EndOfTable(self):
        return self._i >= len(self._ids)

    def GetNextRow(self):
        it = self._lookup[self._ids[self._i]]
        self._i += 1
        return Row(_DEFAULT_COLUMNS, tuple(_column_value(it, n) for n in _DEFAULT_COLUMNS))


class Conversation:
    def __init__(self, ids):
        self._ids, self._lookup = ids, {}

    def GetTable(self):
        return ConversationTable(self._ids, self._lookup)


class OutgoingAppointment:
    """An AppointmentItem made by CreateItem(1). Send() (meeting) or Save() (appointment) records it."""
    Class = 26

    def __init__(self, app, directory):
        self._app = app
        self.Recipients = Recipients(directory)
        self.Subject, self.Body, self.Location = "", "", ""
        self.Start = self.End = None
        self.MeetingStatus, self.ReminderSet, self.ReminderMinutesBeforeStart = 0, True, 15
        self.EntryID = "new-appt"
        self.Sent = self.Saved = False

    def Send(self):
        self.Sent = True
        self._app.sent.append(self)

    def Save(self):
        self.Saved = True
        self._app.saved.append(self)


class Application:
    """Only what outlook_send.py / outlook_meeting.py need: CreateItem, and what Send() / Save() were called on."""

    def __init__(self, directory):
        self.directory, self.sent, self.saved, self.Version = directory, [], [], "16.0.fake"

    def CreateItem(self, item_type):
        if item_type == 1:
            return OutgoingAppointment(self, self.directory)
        assert item_type == 0, item_type
        return Outgoing(self, self.directory)


class Namespace:
    def __init__(self, stores, default_store):
        self.Stores = stores
        self._default = default_store
        self.Accounts = [type("Acc", (), {"SmtpAddress": "me@contoso.com", "DisplayName": "me", "UserName": "me", "AccountType": 0})()]
        self._by_id = {}
        self.directory = Directory()
        self.Application = Application(self.directory)
        for s in stores:
            s._app = self.Application
            for f in _walk(s.GetRootFolder()):
                for it in f._items:
                    self._by_id[it.EntryID] = it
                    if getattr(it, "_conversation", None) is not None:
                        it._conversation._lookup = self._by_id
                    self.directory.learn(getattr(it, "SenderName", ""), getattr(it, "SenderEmailAddress", ""))
                    for r in getattr(it, "Recipients", []):
                        self.directory.learn(r.Name, r.Address)
        self.CurrentUser = type("U", (), {"Name": "Ben", "Address": "me@contoso.com", "AddressEntry": None})()

    def CreateRecipient(self, text):
        name, addr = self.directory.lookup(text)
        return Recipient(name or text, addr or "", 1)

    def GetDefaultFolder(self, fid):
        return self._default.GetDefaultFolder(fid)

    def GetItemFromID(self, eid):
        return self._by_id[eid]


def build_fixture():
    """Two stores: an almost-empty Exchange mailbox and a PST '20230731' with localized folder names."""
    d = dt.datetime
    a = Attachment("合約草稿_v3_legal.docx", 184320)
    m1 = Mail("id1", "Re: 合約草稿 v3 - 法務意見", "Cassie Tsai", "cassie.tsai@contoso.com", d(2026, 9, 12, 16, 42), body="Hi, 法務回來了 hxxp", unread=True, attachments=[a], conv_id="C1", importance=2, recipients=[Recipient("Ben", "ben@contoso.com")])
    m2 = Mail("id2", "合約草稿 v3 - 法務意見", "Ben", "ben@contoso.com", d(2026, 9, 11, 9, 0), body="請法務看一下", conv_id="C1")
    m3 = Mail("id3", "Q3 預算討論", "David WY Chen", "david.chen@contoso.com", d(2026, 9, 8, 9, 12), body="三個方案 A B C 報價", conv_id="C2")
    m4 = Mail("id4", "Weekly newsletter", "News", "news@example.com", d(2026, 9, 1, 8, 0), body="quote of the week", unsubscribe=True)
    m6 = Mail("id6", "AI 人才發展：可以幫我看一下名單嗎？", "PC Liao", "pc.liao@contoso.com", d(2026, 9, 10, 11, 0), body="Ben，麻煩看一下附件名單，週五前回我好嗎？", conv_id="C6", recipients=[Recipient("Ben", "me@contoso.com")], attachments=[Attachment("人才名單.xlsx", 51200, data=b"xlsx")])
    m7 = Mail("id7", "FYI: 季報", "David WY Chen", "david.chen@contoso.com", d(2026, 9, 11, 9, 0), body="供參考，不用回", conv_id="C7", recipients=[Recipient("Ben", "me@contoso.com")], attachments=[Attachment("Q3_report.pptx", 4823040, data=b"pptx"), Attachment("image001.png", 9120, atype=1, data=b"png", inline=True)])
    m5 = Mail("id5", "Old mail", "Cassie Tsai", "cassie.tsai@contoso.com", d(2024, 1, 5, 10, 0), body="old")
    mr1 = MeetingRequest("mr1", "合約討論會議", "Cassie Tsai", "cassie.tsai@contoso.com", d(2026, 9, 13, 9, 0), body="invite", conv_id="C1", unread=True)
    conv = Conversation(["id1", "id2", "mr1"])
    m1._conversation = conv; m2._conversation = conv; mr1._conversation = conv
    sub = Folder("人才", "\\\\20230731\\收件匣\\人才", items=[m3])
    inbox = Folder("收件匣", "\\\\20230731\\收件匣", items=[m1, m2, m4, m5, m6, m7, mr1], subfolders=[sub])
    s1 = Mail("s1", "Re: Q3 預算討論", "Me", "me@contoso.com", d(2026, 9, 9, 10, 0), conv_id="C2", to="David WY Chen; PC Liao", recipients=[Recipient("David WY Chen", "david.chen@contoso.com"), Recipient("PC Liao", "pc.liao@contoso.com")])
    s2 = Mail("s2", "報價單請確認", "Me", "me@contoso.com", d(2026, 9, 5, 10, 0), conv_id="C5", to="PC Liao", recipients=[Recipient("PC Liao", "pc.liao@contoso.com")], body="請確認報價")
    s3 = Mail("s3", "Re: 合約草稿 v3 - 法務意見", "Me", "me@contoso.com", d(2026, 9, 12, 18, 0), conv_id="C1", to="Cassie Tsai", recipients=[Recipient("Cassie Tsai", "cassie.tsai@contoso.com")], body="收到，謝謝")
    sent = Folder("寄件備份", "\\\\20230731\\寄件備份", items=[s1, s2, s3])
    ap1 = Appointment("a1", "每日站會", d(2026, 9, 16, 9, 30), d(2026, 9, 16, 10, 0), recurring=True, organizer="Cassie Tsai", location="Teams")
    ap2 = Appointment("a2", "供應商簡報", d(2026, 9, 16, 14, 0), d(2026, 9, 16, 15, 0), organizer="Cassie Tsai")
    ap2.Recipients = [Recipient("Cassie Tsai", "cassie.tsai@contoso.com"), Recipient("PC Liao", "pc.liao@contoso.com"), Recipient("Ben", "me@contoso.com")]
    ap2.RequiredAttendees = "Cassie Tsai; PC Liao; Ben"
    ap3 = Appointment("a3", "1:1 with Bob", d(2026, 9, 16, 14, 30), d(2026, 9, 16, 15, 30), busy=1, response=2)
    ap4 = Appointment("a4", "Free block", d(2026, 9, 16, 11, 0), d(2026, 9, 16, 12, 0), busy=0)
    cal = Folder("行事曆", "\\\\20230731\\行事曆", items=[ap1, ap2, ap3, ap4], item_type=1)
    pst_root = Folder("20230731", "\\\\20230731", subfolders=[inbox, sent, cal])
    pst = Store("20230731", pst_root, {6: inbox, 5: sent, 9: cal}, file_path="D:\\Mailbox\\20230731.pst", store_type=3)
    ex_inbox = Folder("Inbox", "\\\\ben@contoso.com\\Inbox", items=[Mail("e1", "Ping", "PC Liao", "pc.liao@contoso.com", d(2026, 9, 14, 8, 0), body="ping")])
    ex_cal = Folder("Calendar", "\\\\ben@contoso.com\\Calendar", items=[], item_type=1)
    ex_root = Folder("ben@contoso.com", "\\\\ben@contoso.com", subfolders=[ex_inbox, ex_cal])
    ex = Store("ben@contoso.com", ex_root, {6: ex_inbox, 5: Folder("Sent Items", "\\\\ben@contoso.com\\Sent Items"), 9: ex_cal}, store_type=0)
    return Namespace([ex, pst], ex)
