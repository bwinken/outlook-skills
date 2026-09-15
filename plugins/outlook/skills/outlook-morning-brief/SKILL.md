---
name: outlook-morning-brief
description: Morning brief from the local Outlook (Classic) mailbox, READ-ONLY - today's meetings, new and unread mail grouped by urgency, what the user still owes a reply to, and what they are waiting on from others; any part can be asked for on its own at any time of day. Use when the user says "good morning", "brief me", "catch me up", "what's new in my inbox", "anything urgent", "what do I need to do today", or 早安 / 今天怎樣 / 幫我簡報一下 / 有什麼新信 / 有沒有急事 / 未讀摘要 / 今天要處理什麼.
---

# outlook-morning-brief

Read-only. Composes three existing scripts into one briefing; nothing is marked read or changed.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line and point the user to Claude Code.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_calendar.py"                                          # today's agenda
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_search.py" -Unread -After <since> -PreviewLength 400 -Max 200 [-Store X | -AllStores] [-AllFolders]
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_followup.py" -Direction received [-Store X | -AllStores]
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_followup.py" -Direction sent [-Store X | -AllStores]
```

Add `-HighImportance` or `-Flagged` searches when the user asks for "urgent" or "flagged" even if already read.

## Workflow

1. Scope from the wording. "早安 / brief me / 今天怎樣" = all four sections. "有什麼新信 / 有沒有急事 / unread" = the mail section only. "今天要處理什麼" = agenda + needs-action + owed replies. Do not run scripts for sections that will not be shown.
2. `since` for mail: on a weekday, yesterday 18:00; on Monday, Friday 18:00; if the user says "this week", Monday 00:00. Unread only, unless the user says "all".
3. Use the `store` setting or `-AllStores` when mail lives in a PST; `-AllFolders` when rules file mail into subfolders.
4. Load the memory index and open `people` / `projects` notes matching senders, attendees or subjects so key contacts and live projects are recognised.
5. Rank and group mail per reference.md using previews only; open a thread with `outlook-thread` for at most the two or three that decide the day.
6. Present per reference.md. Keep the whole brief under one screen: counts and the few items that matter, not every mail. Offer read-only next steps: open a thread, prep a meeting (`outlook-meeting-prep`), draft a reply in chat.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant). `first_run: true` means `~/.outlook-skills` does not exist yet: hand over to `outlook-memory`'s onboarding first (respect a "not now"). Apply `store` (pass it as `-Store`), `working_hours`, `language`. When a sender, attendee or project is named in memory, use the note's role and keywords. Never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-morning-brief/reference.md` for the ranking rules and the brief template.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never mark read, flag, reply, accept, decline, move or delete. Drafts, if asked for, are written in chat only.
