#!/usr/bin/env python3
"""Print or write the MCP registration for hosts that do not use the Claude Code plugin marketplace.

    python install.py                    # print what to fill in: the Connectors form (Claude Chat / Cowork),
                                         # a mcpServers JSON snippet (Zoo Code, Cursor, ...), the claude mcp add command
    python install.py --claude-desktop   # older Claude Desktop: merge the entry into claude_desktop_config.json (backup kept)
    python install.py --claude-desktop --uninstall
    python install.py --python C:/path/to/python.exe   # default: the interpreter running this script

Run it with the Python that has pywin32 installed; that interpreter's absolute path is what gets printed.
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
READ_ONLY_TOOLS = ["search_mail", "get_thread", "list_calendar", "list_followups", "prepare_meeting", "mailbox_overview", "get_status"]


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


def report(python: str) -> str:
    args_json = json.dumps([str(SERVER)])
    zoo = {"mcpServers": {NAME: {**entry(python), "alwaysAllow": READ_ONLY_TOOLS, "disabled": False}}}
    return "\n".join([
        "Claude Chat / Cowork: Settings > Connectors > add, Transport = Local command (stdio)",
        f"  Name        {NAME}",
        f"  Command     {python}",
        f"  Arguments   {args_json}",
        "  Environment variables: none needed (optional: OUTLOOK_RERANK_URL / OUTLOOK_RERANK_MODEL / OUTLOOK_RERANK_API_KEY)",
        "",
        "Zoo Code (MCP Servers > Edit Global MCP, or .roo/mcp.json):",
        json.dumps(zoo, ensure_ascii=False, indent=2),
        "",
        "Cursor / VS Code / other MCP hosts:",
        json.dumps({"mcpServers": {NAME: entry(python)}}, ensure_ascii=False, indent=2),
        "",
        "Claude Code without the marketplace:",
        f'  claude mcp add --scope user {NAME} -- "{python}" "{SERVER}"',
        "",
        f"Older Claude Desktop ({desktop_config_path()}):  python install.py --claude-desktop",
        "",
        "Requirements on the Windows machine: Classic Outlook, Python 3.8+, pip install pywin32 in the Python above.",
        f'Check:  "{python}" "{SERVER}" --call get_status "{{}}"',
    ])


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--claude-desktop", action="store_true", help="merge into an older Claude Desktop's config file instead of printing")
    ap.add_argument("--uninstall", action="store_true", help="with --claude-desktop: remove the entry")
    ap.add_argument("--python", default=sys.executable, help="interpreter to run the server with (default: this one)")
    a = ap.parse_args(argv)
    if a.claude_desktop:
        print(merge(desktop_config_path(), a.python, a.uninstall))
        return
    print(report(a.python))


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
