# outlook-mcp (plugin)

A **read-only MCP server** for a local Windows Classic Outlook mailbox. Same scripts, same data and the same
read-only policy as the `outlook` skills plugin, exposed as MCP tools instead of skills, so any MCP host can
use them: the Chat tab of Claude Desktop, Claude Code, or other MCP clients on the Windows machine.

Install **either** this plugin **or** the `outlook` skills plugin. Both work at once, but the MCP server then
starts on every session and its tool definitions sit in the context whether or not Outlook is needed, while
the skills load only when a request matches them.

| Tool | What it does |
|---|---|
| `search_mail` | Mail by sender, recipient, subject, body, date range, unread, attachments; one folder, a store, or every store. With `query` and a configured reranker gateway the hits are reranked by relevance |
| `get_thread` | A whole conversation, oldest first, full bodies |
| `list_calendar` | Calendar items in a range, recurrences expanded, conflicts flagged |
| `list_followups` | Replies the user is waiting for (`direction=sent`) or still owes (`direction=received`) |
| `prepare_meeting` | Attendees, recent mail with them, mail about the subject, attachments, for the next or a named meeting |
| `find_attachments` | Attachments by name, extension, size on top of every search filter; `saveto` copies them to a folder |
| `mailbox_overview` | Top senders and recipients, folders, topics, newsletters, recurring meetings; no bodies |
| `get_status` | Outlook version, profiles, accounts, .pst/.ost stores, folder counts |
| `parse_msg_file` | Parse .msg / .eml files on disk without Outlook (any OS) |

Tool inputs are the options of the matching script in `scripts/` (`search_mail` takes `from`, `after`, `anyof`, `store`, `max`, ...); the
server derives each input schema from that script's argparse parser, so `python server.py --list` is the reference.

## Install

**Claude Code** (plugin marketplace):

```
pip install pywin32
/plugin marketplace add bwinken/outlook-skills
/plugin install outlook-mcp@outlook-skills
```

**Claude Desktop, Chat tab** (or any MCP host that runs a stdio server): clone or download this repo, then

```
python <repo>/plugins/outlook-mcp/install.py --claude-desktop
```

writes the `outlook` entry into `%APPDATA%\Claude\claude_desktop_config.json` (backup kept; `--uninstall` removes it). Without the flag it
only prints the snippet, which is:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["C:\\path\\to\\outlook-skills\\plugins\\outlook-mcp\\server.py"]
    }
  }
}
```

`--python C:/path/to/python.exe` picks the interpreter that has pywin32 when several are installed. Restart Claude Desktop afterwards.

**Claude Code without the marketplace**: `claude mcp add --scope user outlook -- python "<repo>/plugins/outlook-mcp/server.py"`.

## Requirements

- Windows with **Classic Outlook** (2016 / 2019 / 2021 / Microsoft 365). "New Outlook" has no COM object model; only `parse_msg_file` and `get_status` with `skipcom=true` work there.
- Python 3.8+ on `PATH` as `python`, and `pip install pywin32`. Everything else, the MCP protocol included, is the standard library.
- Outlook may be open or closed. The first tool call attaches to the running Outlook or starts it in the background. Outlook is not touched at server start-up, so the host starts fast even when Outlook is unavailable; the tool then returns an error instead.

## Settings, store default, reranker

The server reads the same `~/.outlook-skills/settings.json` (and `./.outlook-skills/settings.json`) as the skills plugin; see the repo README for the keys.

- `store`: applied automatically to every tool that takes a store, when the call gives neither `store` nor `allstores`. Set it when the mail lives in a .pst: `python scripts/settings.py set store "20230731"`.
- `rerank.gateway` (or the `OUTLOOK_RERANK_URL` environment variable): when set, `search_mail` with a `query` runs the search with a wide candidate cap (`search.max_candidates`, default 300), sends subject, sender, date and a 500-character preview of each candidate to that gateway in batches of 30, and returns the top `max` (default 10) sorted by `Score`. **No per-call consent question**: configuring the gateway is the consent. Without a gateway the query is ignored and the reply says so in `Rerank.Reason`. The `ANTHROPIC_BASE_URL` fallback that `scripts/rerank.py` uses on the command line is not used here, so previews never go to a gateway the user did not pick for reranking. `rerank.model` and `rerank.api_key` apply as documented for the skills plugin.

Memory notes (`~/.outlook-skills/memory/`) are a skills feature and are not read by the server.

## Read-only policy

Identical to the `outlook` plugin: no `Save`, `Send`, `Delete`, `Move`, `Copy`, `Display`, no property setters, nothing marked read. Reading through COM does not change read/unread state. The only files written are those the user asks for: `find_attachments` with `saveto` and `parse_msg_file` with `extract_to`. Those two tools carry `readOnlyHint: false` in their MCP annotations; every other tool is `readOnlyHint: true`.

## Layout and maintenance

```
server.py          the MCP server (stdio, JSON-RPC 2.0, newline-delimited); --list and --call for debugging
install.py         Claude Desktop config helper
scripts/           verbatim copy of plugins/outlook/scripts, do not edit here
tests/             python tests/test_server.py (fake Outlook in-process, then a real subprocess over stdio)
```

`scripts/` must be a copy because Claude Code installs only a plugin's own directory and git symlinks are not reliable on Windows. Edit
`plugins/outlook/scripts` and run `python tools/sync_scripts.py`; CI runs `--check` and fails when the two differ.

Protocol: `initialize` (versions 2025-06-18, 2025-03-26, 2024-11-05), `ping`, `tools/list`, `tools/call`. Errors from the scripts come back as
tool results with `isError: true`, never as a crash. If a COM call fails because Outlook was closed or restarted, the server drops its
cached connection and retries the call once.
