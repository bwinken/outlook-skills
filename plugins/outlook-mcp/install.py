#!/usr/bin/env python3
"""Register the Outlook MCP server with hosts that do not use the Claude Code plugin marketplace.

    python install.py                    # print the config snippet (absolute path to server.py) and the claude mcp add command
    python install.py --claude-desktop   # merge the "outlook" entry into Claude Desktop's claude_desktop_config.json (backup kept)
    python install.py --claude-desktop --uninstall
    python install.py --python C:/Python312/python.exe   # use a specific interpreter (the one with pywin32)

Claude Code users normally install the plugin instead: /plugin install outlook-mcp@outlook-skills.
"""
import argparse
import json
import os
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SERVER = HERE / "server.py"
NAME = "outlook"


def desktop_config_path() -> Path:
    if sys.platform == "win32":
        return Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")) / "Claude" / "claude_desktop_config.json"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "Claude" / "claude_desktop_config.json"
    return Path.home() / ".config" / "Claude" / "claude_desktop_config.json"


def entry(python: str) -> dict:
    return {"command": python, "args": [str(SERVER)]}


def merge(path: Path, python: str, remove: bool) -> str:
    data = {}
    if path.is_file():
        data = json.loads(path.read_text(encoding="utf-8-sig") or "{}")
        shutil.copy2(path, path.with_suffix(".json.bak"))
    servers = data.setdefault("mcpServers", {})
    if remove:
        if NAME not in servers:
            return f"no '{NAME}' entry in {path}"
        del servers[NAME]
    else:
        servers[NAME] = entry(python)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return f"{'removed from' if remove else 'written to'} {path} (backup: {path.with_suffix('.json.bak').name}). Restart Claude Desktop."


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--claude-desktop", action="store_true", help="merge into Claude Desktop's config file instead of printing")
    ap.add_argument("--uninstall", action="store_true", help="with --claude-desktop: remove the entry")
    ap.add_argument("--python", default="python", help="interpreter to run the server with (default: python on PATH)")
    a = ap.parse_args(argv)
    if a.claude_desktop:
        print(merge(desktop_config_path(), a.python, a.uninstall))
        return
    snippet = {"mcpServers": {NAME: entry(a.python)}}
    print("Claude Desktop (" + str(desktop_config_path()) + "), or any MCP host that takes a stdio server:\n")
    print(json.dumps(snippet, ensure_ascii=False, indent=2))
    print("\nClaude Code without the marketplace:\n")
    print(f'  claude mcp add --scope user {NAME} -- {a.python} "{SERVER}"')
    print("\nRequirements on the Windows machine: Classic Outlook, Python 3.8+, pip install pywin32.")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
