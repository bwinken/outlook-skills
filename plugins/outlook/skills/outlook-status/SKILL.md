---
name: outlook-status
description: Show the state of the local Windows Outlook (Classic) installation, READ-ONLY. Reports Outlook version, profiles, accounts, .pst/.ost data files with paths and sizes, top-level folder item and unread counts, and warns if New Outlook is enabled. Use when the user asks "what's in my Outlook", "where are my Outlook files", "how big is my mailbox", "which accounts are configured", or 我的 Outlook 狀況 / 資料檔在哪 / 信箱多大.
---

# outlook-status

Read-only overview of the local Outlook setup. Nothing in Outlook is modified.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, Zoo Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line, and point the user to Claude Code, or to `outlook-open-msg` for .msg/.eml files they export from Outlook.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_status.py"
```

Pure Python (needs `pip install pywin32` once on the Windows machine). Options accept PowerShell-style `-From` or `--from` spellings. Run it from any tool (Bash, cmd, PowerShell); nothing goes through a shell that could mangle quotes, `$` or Chinese. Output is UTF-8 JSON on stdout; `-OutFile <file>` writes it to a file instead (use for large results). If it reports that pywin32 is missing or that Outlook COM cannot start, say so and point to the plugin README requirements.

Options:
- `-OutFile <path>`: write the JSON to a file instead of stdout (use for large mailboxes).
- `-SkipCom`: only inspect the registry and file system; do not connect to Outlook. Use this when Outlook is not installed, is New Outlook, or the user does not want Outlook launched.

## What the JSON contains

- `OfficeVersion`, `OutlookExeVersion`, `DefaultProfile`, `Profiles`
- `NewOutlookEnabled`: if true, COM automation will fail; tell the user this plugin needs Classic Outlook.
- `Accounts`: display name, SMTP address, type (Exchange / IMAP / POP3 / EAS).
- `Stores`: each mailbox or data file with `FilePath`, `SizeMB`, `StoreType`, `IsCachedExchange`, and a `Folders` list with item and unread counts.
- `DataFilesOnDisk`: .pst/.ost found in the default locations, even if not attached to a profile.
- `Warnings`: anything that failed, with the reason.

## Presenting the result

- Summarise in a short table: store name, type, file path, size, Inbox unread count.
- Point out orphaned data files (on disk but not in `Stores`) and stores close to typical size limits (50 GB for modern .pst/.ost).
- If the user wants cleanup, only *recommend*: this plugin never deletes, compacts or moves anything.
- If the default store's Inbox is nearly empty while another store (typically a .pst archive) holds most items, say so and offer to make that store the default for the other skills: `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" set store "<store DisplayName>"`. Only run it after the user agrees.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant).

- `first_run: true` means `~/.outlook-skills` does not exist yet: switch to `outlook-memory`'s onboarding, which asks the user (structured question tool) whether to create a personal memory and scan the mailbox. Respect a "not now" and continue here.
- Apply the merged `settings` (`status.skip_com`, `store`, `language`).
- `memory` is an index (title, category, tags, updated, path), not the notes themselves. When the request names a person, folder, project or routine, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "<word>"` and `show` the matching note, so "Alice" or "供應商的信" resolve to the right address or folder. Do not load every note.
- If the user states something worth keeping, offer to save it through `outlook-memory`; never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-status/reference.md` before presenting results. It documents every JSON field the script returns and the presentation template to use in the reply.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never add commands that call Save, Send, Delete, Move, Compact, or that change any property.
