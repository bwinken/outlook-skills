---
name: outlook-digest
description: Triage the Outlook (Classic) inbox READ-ONLY - today's or this week's unread and recent mail grouped by urgency (needs action, waiting on me, FYI, newsletters), with key people from memory highlighted. Use when the user asks "what's new in my inbox", "summarise my unread mail", "anything urgent", "catch me up on today's email", or 今天有什麼信 / 未讀摘要 / 有沒有急事 / 幫我看一下收件匣.
---

# outlook-digest

Read-only inbox triage built on `outlook_search.py` plus the ranking rules in reference.md. Nothing is marked read.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line and point the user to Claude Code.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_search.py" -Unread -After <today or Monday> -PreviewLength 400 -Max 200 [-Store X | -AllStores] [-AllFolders]
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_search.py" -After <today> -HighImportance   # add when the user wants "urgent" even if read
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_search.py" -Flagged                        # follow-up flags the user set themselves
```

## Workflow

1. Range: "today" = since 00:00 today; "this week" = since Monday; default today. Unread only unless the user says "all".
2. Run the search (add `-AllFolders` when the user has rules that file mail into subfolders; `outlook-status` shows folder counts).
3. Load the memory index and open the `people` and `projects` notes that match senders or subjects, so key contacts and active projects are recognised.
4. Rank and group per reference.md. Read previews only; do not run `-IncludeBody` for every mail. Open a thread with `outlook-thread` only for the two or three that matter most, if needed to judge.
5. Present the digest. Offer follow-ups that stay read-only: open a thread, list who is waiting (`outlook-followup`), draft a reply in chat.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant). `first_run: true` means `~/.outlook-skills` does not exist yet: hand over to `outlook-memory`'s onboarding first (respect a "not now"). Apply `store` (pass it as `-Store`), `language`, and any search defaults. When the request names a person, folder or project, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "<word>"` and `show` the note so aliases resolve. Never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-digest/reference.md` for the ranking rules and the digest template.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never reply, forward, flag, mark read, move or delete. Drafts, if asked for, are written in chat only.

