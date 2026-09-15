"""Parses real .msg samples (from the extract-msg project) with the standard-library reader.
Samples are fetched on demand; the test is skipped when they cannot be downloaded."""
import os, sys, urllib.request, tempfile, json
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
BASE = "https://raw.githubusercontent.com/TeamMsgExtractor/msg-extractor/master/example-msg-files/"
NAMES = ["strangeDate.msg", "multi-to.msg", "unicode.msg"]
d = tempfile.mkdtemp()
files = []
for n in NAMES:
    try:
        urllib.request.urlretrieve(BASE + n, os.path.join(d, n)); files.append(os.path.join(d, n))
    except Exception as e:
        print("skip (download failed):", n, e)
if not files:
    print("no samples; skipped"); sys.exit(0)
import read_msg
from msgfile import MsgFile
for f in files:
    r = read_msg.parse_msg(f, True, None)
    json.dumps(r, ensure_ascii=False)
    assert r["subject"] and r["to"], r
    print("ok", os.path.basename(f), "|", r["subject"], "|", len(r["attachments"]), "attachments")
if any(f.endswith("unicode.msg") for f in files):
    m = MsgFile([f for f in files if f.endswith("unicode.msg")][0])
    assert [a.filename for a in m.attachments] == ["import OleFileIO.tif", "raised value error.tif"]
    assert m.attachments[0].data[:2] == b"II" and len(m.attachments[0].data) == 969674
    assert m.sender_email == "brizhou@gmail.com" and m.message_id.startswith("<CADtJ4e")
if any(f.endswith("multi-to.msg") for f in files):
    r = read_msg.parse_msg([f for f in files if f.endswith("multi-to.msg")][0], False, None)
    assert [x["address"] for x in r["to"]] == ["alice@example.com", "carol@example.com", "alice@example.com"]
    assert r["cc"][0]["address"] == "dave@example.com" and r["date"].startswith("2021-05-28")
print("msgfile tests passed")
