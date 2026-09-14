---
name: outlook-status
description: Show the state of the local Windows Outlook (Classic) installation, READ-ONLY. Reports Outlook version, profiles, accounts, .pst/.ost data files with paths and sizes, top-level folder item and unread counts, and warns if New Outlook is enabled. Use when the user asks "what's in my Outlook", "where are my Outlook files", "how big is my mailbox", "which accounts are configured", or 我的 Outlook 狀況 / 資料檔在哪 / 信箱多大.
---

# outlook-status

Read-only overview of the local Outlook setup. Nothing in Outlook is modified.

## Run

```
powershell -NoProfile -ExecutionPolicy Bypass -File "${CLAUDE_PLUGIN_ROOT}/scripts/Get-OutlookStatus.ps1"
```

If that fails with "running scripts is disabled on this system" (execution policy enforced by Group Policy, so `-ExecutionPolicy Bypass` is ignored), use the policy-free form, which loads the script text as a script block instead of running the file:

```
powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='${CLAUDE_PLUGIN_ROOT}/scripts'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '${CLAUDE_PLUGIN_ROOT}/scripts/Get-OutlookStatus.ps1')))"
```

Use Windows paths with backslashes inside the single quotes if forward slashes are rejected. Do not try to change the machine's execution policy; that is the user's or IT's decision. See the plugin README section "Execution policy" for the AppLocker / Constrained Language case.

Options:
- `-OutFile <path>`: write JSON to a file instead of stdout (use for large mailboxes).
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

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-status/reference.md` before presenting results. It documents every JSON field the script returns and the presentation template to use in the reply.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never add commands that call Save, Send, Delete, Move, Compact, or that change any property.
