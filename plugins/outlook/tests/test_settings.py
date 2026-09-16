"""settings.py and the rerank config that reads it, against a temporary home directory."""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

SCRIPTS = os.path.join(os.path.dirname(__file__), "..", "scripts")
sys.path.insert(0, SCRIPTS)
import settings as ps  # noqa: E402
import rerank  # noqa: E402


class SettingsTest(unittest.TestCase):
    def setUp(self):
        # resolved, because settings.local_dir() resolves paths and Windows temp dirs come back as 8.3 short names
        self.home = str(Path(tempfile.mkdtemp()).resolve())
        self.cwd = str(Path(tempfile.mkdtemp()).resolve())
        self.env = mock.patch.dict(os.environ, {"HOME": self.home, "USERPROFILE": self.home})
        self.env.start()
        self.addCleanup(self.env.stop)
        self.old_cwd = os.getcwd()
        os.chdir(self.cwd)
        self.addCleanup(os.chdir, self.old_cwd)

    def write(self, folder, data):
        os.makedirs(folder, exist_ok=True)
        with open(os.path.join(folder, "settings.json"), "w", encoding="utf-8") as fh:
            json.dump(data, fh)

    def test_layers_merge_key_by_key(self):
        self.write(os.path.join(self.home, ".outlook-skills"), {"store": "PST", "working_hours": {"end": "17:30"}})
        self.write(os.path.join(self.cwd, ".outlook-skills"), {"working_hours": {"start": "10:00"}})
        merged, sources, layers, ld = ps.resolve()
        self.assertEqual(merged["store"], "PST")
        self.assertEqual((merged["working_hours"]["start"], merged["working_hours"]["end"]), ("10:00", "17:30"))
        self.assertEqual(merged["working_hours"]["days"], [1, 2, 3, 4, 5])  # default kept
        self.assertEqual(len(layers), 2)
        self.assertTrue(sources["working_hours.start"].startswith(self.cwd))
        self.assertEqual(str(ld), os.path.join(self.cwd, ".outlook-skills"))

    def test_rerank_section_reaches_rerank_py(self):
        self.write(os.path.join(self.home, ".outlook-skills"), {"rerank": {"gateway": "http://gw:8000/v1", "model": "m", "api_key": "k"}})
        self.assertEqual(rerank._plugin_settings(), {"OUTLOOK_RERANK_URL": "http://gw:8000/v1", "OUTLOOK_RERANK_MODEL": "m", "OUTLOOK_RERANK_API_KEY": "k"})
        with mock.patch.dict(os.environ, {k: "" for k in ("OUTLOOK_RERANK_URL", "ANTHROPIC_BASE_URL", "OPENAI_BASE_URL")}):
            import argparse
            url, model, key, src = rerank.resolve_config(argparse.Namespace(gateway=None, model=None, api_key=None))
        self.assertEqual((url, model, key, src["gateway"]), ("http://gw:8000/v1", "m", "k", ".outlook-skills:OUTLOOK_RERANK_URL"))

    def test_cli_path_and_set(self):
        run = lambda *args: subprocess.run([sys.executable, os.path.join(SCRIPTS, "settings.py"), *args], capture_output=True, text=True, env=os.environ.copy(), cwd=self.cwd)
        r = run("path")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(json.loads(r.stdout)["user_dir"], os.path.join(self.home, ".outlook-skills"))
        r = run("set", "rerank.auto_consent", "true")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIs(ps.resolve()[0]["rerank"]["auto_consent"], True)


if __name__ == "__main__":
    unittest.main()
