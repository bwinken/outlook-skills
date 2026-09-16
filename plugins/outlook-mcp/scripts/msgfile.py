#!/usr/bin/env python3
"""Standard-library reader for Outlook .msg files (OLE2 compound file + MAPI property streams).

READ-ONLY: the file is opened in binary read mode and never modified.

    from msgfile import MsgFile
    m = MsgFile("mail.msg")
    m.subject, m.sender_name, m.sender_email, m.to, m.cc, m.bcc, m.body, m.html_body,
    m.headers (raw transport headers), m.message_id, m.date (datetime, UTC), m.attachments
    (list of Attachment with .filename, .data, .mime), m.recipients (list of dicts)

Covers what the skills need. Not a full MAPI implementation: named properties, embedded
message attachments (returned with data=None) and RTF-compressed bodies are out of scope.
"""
import datetime as dt
import struct

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FILETIME_EPOCH = dt.datetime(1601, 1, 1, tzinfo=dt.timezone.utc)


# ---------------------------------------------------------------- OLE2 compound file
class _Entry:
    __slots__ = ("name", "type", "left", "right", "child", "start", "size", "children")

    def __init__(self, name, etype, left, right, child, start, size):
        self.name, self.type, self.left, self.right, self.child, self.start, self.size = name, etype, left, right, child, start, size
        self.children = {}


class CompoundFile:
    def __init__(self, path):
        with open(path, "rb") as fh:
            self.data = fh.read()
        d = self.data
        if d[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("not an OLE2 compound file (bad magic)")
        self.sector_size = 1 << struct.unpack_from("<H", d, 0x1E)[0]
        self.mini_size = 1 << struct.unpack_from("<H", d, 0x20)[0]
        n_fat = struct.unpack_from("<I", d, 0x2C)[0]
        first_dir = struct.unpack_from("<I", d, 0x30)[0]
        self.mini_cutoff = struct.unpack_from("<I", d, 0x38)[0]
        first_minifat = struct.unpack_from("<I", d, 0x3C)[0]
        n_minifat = struct.unpack_from("<I", d, 0x40)[0]
        first_difat = struct.unpack_from("<I", d, 0x44)[0]
        n_difat = struct.unpack_from("<I", d, 0x48)[0]

        difat = list(struct.unpack_from("<109I", d, 0x4C))
        sec = first_difat
        per = self.sector_size // 4 - 1
        for _ in range(n_difat):
            if sec in (ENDOFCHAIN, FREESECT):
                break
            raw = self._sector(sec)
            entries = struct.unpack_from(f"<{per + 1}I", raw)
            difat.extend(entries[:per])
            sec = entries[per]
        self.fat = []
        for s in difat[:n_fat]:
            if s in (ENDOFCHAIN, FREESECT):
                continue
            self.fat.extend(struct.unpack_from(f"<{self.sector_size // 4}I", self._sector(s)))

        dir_bytes = self._read_chain(first_dir)
        self.entries = []
        for off in range(0, len(dir_bytes), 128):
            e = dir_bytes[off:off + 128]
            if len(e) < 128:
                break
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(0, nlen - 2)].decode("utf-16-le", "replace") if nlen >= 2 else ""
            etype = e[66]
            left, right, child = struct.unpack_from("<3I", e, 68)
            start = struct.unpack_from("<I", e, 116)[0]
            size = struct.unpack_from("<Q", e, 120)[0]
            if self.sector_size == 512:
                size &= 0xFFFFFFFF
            self.entries.append(_Entry(name, etype, left, right, child, start, size))
        self.root = self.entries[0]
        self._build_tree(self.root)

        self.minifat = []
        if n_minifat and first_minifat not in (ENDOFCHAIN, FREESECT):
            mf = self._read_chain(first_minifat)
            self.minifat = list(struct.unpack_from(f"<{len(mf) // 4}I", mf))
        self.ministream = self._read_chain(self.root.start, self.root.size) if self.root.start not in (ENDOFCHAIN, FREESECT) else b""

    def _sector(self, n):
        off = (n + 1) * self.sector_size
        return self.data[off:off + self.sector_size]

    def _read_chain(self, start, size=None):
        out, sec, seen = [], start, set()
        while sec not in (ENDOFCHAIN, FREESECT) and sec < len(self.fat) + 1 and sec not in seen:
            seen.add(sec)
            out.append(self._sector(sec))
            if sec >= len(self.fat):
                break
            sec = self.fat[sec]
        blob = b"".join(out)
        return blob[:size] if size is not None else blob

    def _read_mini(self, start, size):
        out, sec, seen = [], start, set()
        while sec not in (ENDOFCHAIN, FREESECT) and sec < len(self.minifat) and sec not in seen:
            seen.add(sec)
            off = sec * self.mini_size
            out.append(self.ministream[off:off + self.mini_size])
            sec = self.minifat[sec]
        return b"".join(out)[:size]

    def _build_tree(self, storage):
        stack = [storage.child]
        while stack:
            i = stack.pop()
            if i == FREESECT or i >= len(self.entries):
                continue
            e = self.entries[i]
            storage.children[e.name] = e
            stack.extend([e.left, e.right])
            if e.type == 1:
                self._build_tree(e)

    def stream(self, entry):
        if entry.size == 0:
            return b""
        if entry.size < self.mini_cutoff:
            return self._read_mini(entry.start, entry.size)
        return self._read_chain(entry.start, entry.size)


# ---------------------------------------------------------------- MSG properties
PT_LONG, PT_BOOLEAN, PT_SYSTIME, PT_STRING8, PT_UNICODE, PT_BINARY = 0x0003, 0x000B, 0x0040, 0x001E, 0x001F, 0x0102


class _Props:
    """Properties of one storage: fixed-size ones from __properties_version1.0, variable ones from __substg streams."""

    def __init__(self, cf, storage, header_len, codepage=None):
        self.cf, self.storage = cf, storage
        self.fixed = {}
        pe = storage.children.get("__properties_version1.0")
        if pe is not None:
            blob = cf.stream(pe)
            for off in range(header_len, len(blob) - 15, 16):
                ptype, pid = struct.unpack_from("<HH", blob, off)
                self.fixed[pid] = (ptype, blob[off + 8:off + 16])
        self.codepage = codepage or self._codepage()

    def _codepage(self):
        for pid in (0x3FDE, 0x3FFD):  # PR_INTERNET_CPID, PR_MESSAGE_CODEPAGE
            v = self.int(pid)
            if v:
                return {20127: "ascii", 28591: "latin-1", 65001: "utf-8", 936: "gbk", 950: "big5", 932: "cp932", 949: "cp949", 1200: "utf-16-le"}.get(v, f"cp{v}")
        return "cp1252"

    def _stream(self, pid, ptype):
        e = self.storage.children.get(f"__substg1.0_{pid:04X}{ptype:04X}")
        return self.cf.stream(e) if e is not None else None

    def str(self, pid):
        b = self._stream(pid, PT_UNICODE)
        if b is not None:
            return b.decode("utf-16-le", "replace").rstrip("\x00")
        b = self._stream(pid, PT_STRING8)
        if b is not None:
            try:
                return b.decode(self.codepage, "replace").rstrip("\x00")
            except LookupError:
                return b.decode("cp1252", "replace").rstrip("\x00")
        return None

    def bin(self, pid):
        return self._stream(pid, PT_BINARY)

    def int(self, pid):
        f = self.fixed.get(pid)
        if f and f[0] in (PT_LONG, PT_BOOLEAN):
            return struct.unpack_from("<I", f[1])[0] if f[0] == PT_LONG else struct.unpack_from("<H", f[1])[0]
        return None

    def time(self, pid):
        f = self.fixed.get(pid)
        if f and f[0] == PT_SYSTIME:
            ft = struct.unpack_from("<Q", f[1])[0]
            if ft:
                return FILETIME_EPOCH + dt.timedelta(microseconds=ft // 10)
        return None

    def storages(self, prefix):
        return [e for n, e in sorted(self.storage.children.items()) if e.type == 1 and n.startswith(prefix)]


class Attachment:
    def __init__(self, props: _Props):
        self.filename = props.str(0x3707) or props.str(0x3704) or props.str(0x3001) or "(unnamed)"
        self.mime = props.str(0x370E)
        self.method = props.int(0x3705)  # 1 = by value, 5 = embedded message, 6 = OLE
        self.data = props.bin(0x3701) if self.method != 5 else None
        self.size = len(self.data) if self.data is not None else props.int(0x0E20)
        self.content_id = props.str(0x3712)
        self.is_hidden = bool(props.int(0x7FFE) or 0)


class MsgFile:
    def __init__(self, path):
        self.cf = CompoundFile(path)
        p = _Props(self.cf, self.cf.root, 32)
        self.props = p
        self.message_class = p.str(0x001A)
        self.subject = p.str(0x0037) or p.str(0x0E1D) or ""
        self.sender_name = p.str(0x0C1A) or p.str(0x0042) or ""
        self.sender_email = p.str(0x5D01) or p.str(0x5D02) or p.str(0x0C1F) or p.str(0x0065) or ""
        self.to, self.cc, self.bcc = p.str(0x0E04) or "", p.str(0x0E03) or "", p.str(0x0E02) or ""
        self.body = p.str(0x1000) or ""
        html = p.bin(0x1013)
        if html is None:
            html = p.str(0x1013)
        self.html_body = html.decode(p.codepage, "replace") if isinstance(html, (bytes, bytearray)) else html
        self.headers = p.str(0x007D) or ""
        self.message_id = p.str(0x1035)
        self.in_reply_to = p.str(0x1042)
        self.references = p.str(0x1039)
        self.date = p.time(0x0039) or p.time(0x0E06) or p.time(0x3007)
        self.recipients = []
        for st in p.storages("__recip_version1.0_"):
            rp = _Props(self.cf, st, 8, p.codepage)
            self.recipients.append({
                "name": rp.str(0x3001) or "", "address": rp.str(0x39FE) or rp.str(0x3003) or "",
                "type": {1: "to", 2: "cc", 3: "bcc"}.get(rp.int(0x0C15) or 0, "to"),
            })
        self.attachments = [Attachment(_Props(self.cf, st, 8, p.codepage)) for st in p.storages("__attach_version1.0_")]

    def recipients_of(self, rtype):
        return [r for r in self.recipients if r["type"] == rtype]
