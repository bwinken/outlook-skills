# outlook-mcp (plugin)

An MCP server over the same scripts as the `outlook` skills plugin, for hosts that want tools instead of skills: Claude Chat / Cowork, Claude Code, Zoo Code, any MCP client on the Windows machine.
Install this **or** `outlook`, not both. Per-client setup: [docs/mcp.md](../../docs/mcp.md).

| Tool | What it does |
|---|---|
| `search_mail` | Mail by sender, recipient, subject, body, dates, unread, attachments; one folder, one store or all. With `query` and a configured reranker, reranked by relevance |
| `get_thread` | A whole conversation, oldest first, full bodies |
| `list_calendar` | Calendar items in a range, recurrences expanded, conflicts flagged |
| `list_followups` | Replies the user is waiting for (`direction=sent`), still owes (`received`), or both from one scan (`both`) |
| `prepare_meeting` | Attendees, recent mail with them, mail about the subject, attachments |
| `find_attachments` | Attachments by name, extension, size on top of every search filter; `saveto` copies them out |
| `mailbox_overview` | Top senders, recipients, folders, topics, newsletters, recurring meetings; no bodies |
| `get_status` | Outlook version, profiles, accounts, .pst/.ost stores, folder counts |
| `parse_msg_file` | Parse .msg / .eml files without Outlook, any OS |
| `draft_mail` | Step 1 of sending: resolve recipients, store the exact outgoing text, return `id` and `confirm`; sends nothing |
| `send_mail` | Step 2: a confirmation window on the user's desktop shows To, Cc, Subject and the full text; Send happens only after they click Send there |

Tool inputs are the options of the matching script; `python server.py --list` prints the schemas, `python server.py --call get_status '{}'` runs one tool.

## Requirements

Windows with Classic Outlook, Python 3.8+ on `PATH` as `python`, `pip install pywin32`. The MCP protocol itself is standard library only. Outlook is first touched on the first tool call, so the host starts fast even when Outlook is unavailable.

## Settings

Same `~/.outlook-skills/settings.json` as the skills ([keys](../../README.md#設定)). Two differences:

- `store` is applied to every call that names neither `store` nor `allstores`.
- `rerank.gateway` (or `OUTLOOK_RERANK_URL`) turns on reranking for `search_mail` with a `query`, without a per-call consent question. The `ANTHROPIC_BASE_URL` fallback of the skills version is not used here.

Memory notes are a skills feature and are not read by the server.

## Read-only, except sending

No `Save`, `Delete`, `Move`, `Display`, no property setters, nothing marked read. `send_mail` is the only tool that changes anything, and only after the user's click in the desktop window (`destructiveHint: true`). The only files written are drafts under `~/.outlook-skills/drafts/` and the ones the user asks for (`find_attachments` `saveto`, `parse_msg_file` `extract_to`); those tools carry `readOnlyHint: false`.

## Layout

```
server.py      stdio JSON-RPC 2.0 server; initialize, ping, tools/list, tools/call; --list / --call for debugging
install.py     prints or writes the MCP registration for non-marketplace hosts
scripts/       verbatim copy of plugins/outlook/scripts (edit there, then python tools/sync_scripts.py; CI checks)
tests/         python -m unittest discover -s tests
```

Script errors come back as tool results with `isError: true`. If Outlook was closed or restarted, the server reconnects and retries the call once. Version comes from `.claude-plugin/plugin.json`.
