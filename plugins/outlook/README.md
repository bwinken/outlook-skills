# outlook (plugin)

Read-only skills for a **local Windows Classic Outlook** mailbox.

| Skill | What it does |
|---|---|
| `outlook-status` | Outlook version, profiles, accounts, .pst/.ost files, folder counts |
| `outlook-search` | Search mail by sender, subject, body, date, attachments, unread |
| `outlook-thread` | Read a whole conversation with full bodies, ready to summarise |
| `outlook-calendar` | Agenda for a date range with recurrences expanded and conflicts flagged |
| `outlook-open-msg` | Parse a .msg / .eml file without Outlook |

## Requirements

- Windows with **Classic Outlook** (2016 / 2019 / 2021 / Microsoft 365). "New Outlook" has no COM object model and is not supported; `outlook-status -SkipCom` still works there.
- Windows PowerShell 5.1 (built in) or PowerShell 7 for the four COM-based skills.
- Python 3.8+ for `outlook-open-msg`; `pip install extract-msg` for .msg files.
- Outlook may be open or closed. If closed, the COM call starts it in the background under the current user's profile.

## Read-only policy

This plugin **never writes to Outlook**. Concretely, no script or skill may:

- call `Save`, `Send`, `Delete`, `Move`, `Copy`, `Forward`, `Reply`, `ReplyAll`, `Respond`, `Display`;
- set any property (`UnRead`, `Categories`, `FlagStatus`, `Importance`, `BusyStatus`, ...);
- create items, folders, rules, or appointments;
- compact, repair, detach or attach data files;
- write anywhere except a user-specified `-OutFile` / `--extract-to` path.

Reading through COM does not change read/unread state. The shared module `scripts/OutlookReadOnly.psm1` exposes only getters; add new skills on top of it and keep the same rule.

## Layout

```
plugins/outlook/
  .claude-plugin/plugin.json
  scripts/
    OutlookReadOnly.psm1      shared read-only COM helpers
    Get-OutlookStatus.ps1
    Search-OutlookMail.ps1
    Get-OutlookThread.ps1
    Get-OutlookCalendar.ps1
    read_msg.py
  skills/
    outlook-status/SKILL.md
    outlook-search/SKILL.md
    outlook-thread/SKILL.md
    outlook-calendar/SKILL.md
    outlook-open-msg/SKILL.md
```

Skills reference scripts via `${CLAUDE_PLUGIN_ROOT}`, which Claude Code resolves to the installed plugin directory.
