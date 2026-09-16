"""Parses real .msg samples (from the msg-extractor project) with the standard-library reader.

The samples are downloaded once into tests/.samples/ (git-ignored; CI caches it). Outside CI a
failed download skips the test; on CI (CI=1) it fails, so the suite never passes by accident."""
import json
import os
import sys
import unittest
import urllib.request

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import read_msg  # noqa: E402
from msgfile import MsgFile  # noqa: E402

BASE = "https://raw.githubusercontent.com/TeamMsgExtractor/msg-extractor/master/example-msg-files/"
NAMES = ["strangeDate.msg", "multi-to.msg", "unicode.msg"]
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".samples")


def sample(name):
    path = os.path.join(CACHE, name)
    if not os.path.isfile(path):
        os.makedirs(CACHE, exist_ok=True)
        try:
            urllib.request.urlretrieve(BASE + name, path)
        except Exception as e:
            if os.environ.get("CI"):
                raise
            raise unittest.SkipTest(f"download failed: {name}: {e}")
    return path


class MsgFileTest(unittest.TestCase):
    def test_every_sample_parses(self):
        for n in NAMES:
            r = read_msg.parse_msg(sample(n), True, None)
            json.dumps(r, ensure_ascii=False)
            self.assertTrue(r["subject"] and r["to"], r)

    def test_unicode_attachments(self):
        m = MsgFile(sample("unicode.msg"))
        self.assertEqual([a.filename for a in m.attachments], ["import OleFileIO.tif", "raised value error.tif"])
        self.assertEqual(m.attachments[0].data[:2], b"II")
        self.assertEqual(len(m.attachments[0].data), 969674)
        self.assertEqual(m.sender_email, "brizhou@gmail.com")
        self.assertTrue(m.message_id.startswith("<CADtJ4e"))

    def test_multiple_recipients(self):
        r = read_msg.parse_msg(sample("multi-to.msg"), False, None)
        self.assertEqual([x["address"] for x in r["to"]], ["alice@example.com", "carol@example.com", "alice@example.com"])
        self.assertEqual(r["cc"][0]["address"], "dave@example.com")
        self.assertTrue(r["date"].startswith("2021-05-28"))


if __name__ == "__main__":
    unittest.main()
