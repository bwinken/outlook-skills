---
name: outlook-morning-brief
description: Morning brief from the local Outlook (Classic) mailbox, READ-ONLY - today's meetings, new and unread mail grouped by urgency, replies the user still owes, and mails the user is waiting on; any part can be asked for on its own at any time. Use when the user says "good morning", "brief me", "catch me up", "what's new in my inbox", "anything urgent", "what do I need to do today", "who hasn't replied to me", "which emails do I still owe a reply", or 早安 / 今天怎樣 / 幫我簡報一下 / 有什麼新信 / 有沒有急事 / 今天要處理什麼 / 誰還沒回我 / 我還欠誰回信.
---

# outlook-morning-brief

Read-only. Composes three scripts into one briefing (three runs for the full brief, independent of each other: launch them in one step as parallel tool calls); nothing is marked read or changed.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line and point the user to Claude Code.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_calendar.py"                                          # today's agenda
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_search.py" -Unread -After <since> -PreviewLength 400 -Max 200 [-Store X | -AllStores] [-AllFolders]
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_followup.py" -Direction both [-Store X | -AllStores]      # owed and awaited from one scan
```

Add `-HighImportance` or `-Flagged` searches when the user asks for "urgent" or "flagged" even if already read.

## Workflow

1. Scope from the wording. "早安 / brief me / 今天怎樣" = all four sections. "有什麼新信 / 有沒有急事 / unread" = the mail section only. "今天要處理什麼" = agenda + needs-action + owed replies. Do not run scripts for sections that will not be shown.
2. `since` for mail: on a weekday, yesterday 18:00; on Monday, Friday 18:00; if the user says "this week", Monday 00:00. Unread only, unless the user says "all".
3. Use the `store` setting or `-AllStores` when mail lives in a PST; `-AllFolders` when rules file mail into subfolders.
4. Load the memory index and open `people` / `projects` notes matching senders, attendees or subjects so key contacts and live projects are recognised.
5. Rank and group mail per reference.md using previews only; open a thread with `outlook-thread` for at most the two or three that decide the day.
6. Present per reference.md. Keep the whole brief under one screen: counts and the few items that matter, not every mail. Offer next steps: open a thread, reply through `outlook-send`.

## Follow-ups on their own

"誰還沒回我 / who hasn't replied" and "我還欠誰回信 / what do I owe" are the last two sections asked for alone. Run only `outlook_followup.py` for the matching direction:

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_followup.py" -Direction sent      # I wrote, they did not answer (default: older than 3 days)
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_followup.py" -Direction received  # they asked, I did not answer (default: older than 2 days)
```

`-Direction both` returns both lists from one scan, under `Sent` and `Received`. Options: `-Days N` minimum age before a mail counts as waiting; `-Lookback N` days scanned (default 60); `-Store` / `-AllStores`; `-QuestionsOnly` (received) keeps only mails that look like a question; `-Max`, `-PreviewLength`, `-OutFile`.

A mail counts as answered when a later mail in the same conversation comes from the other side, within Inbox and Sent Items. Replies given by phone or chat, or filed elsewhere by a rule, are not seen: say so when the list looks wrong. For **received**, use `LooksLikeQuestion` and `DirectToMe` plus the preview to separate real asks from FYI mails. Present with the stand-alone templates in reference.md §4. Offer next steps: open the thread, or send a nudge through `outlook-send` (it asks before sending).

## Settings and memory

Run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once per conversation, in the same step as the scripts above (all of them as parallel tool calls; the scripts take `store` from the settings themselves). `settings.py profile show` can join that step: it prints a one-line note when there is no profile yet, otherwise the reply habits to use before drafting a reply in the user's voice or judging which mails need a reply. `first_run: true` means `~/.outlook-skills` does not exist yet: give the brief, then offer `outlook-setup` once (respect a "not now"). Apply `working_hours`, `language`. When a sender, attendee or project matches a title or tag in the `memory` index, open that note with `memory.py show "<title>"` and use its role and keywords. Never write memory silently.

## Output format

`${CLAUDE_PLUGIN_ROOT}/skills/outlook-morning-brief/reference.md` holds the ranking rules and the brief template. Read it once per conversation, in the same step as the scripts (parallel tool calls), not as a separate turn.

## Read-only rules

Read-only, per `${CLAUDE_PLUGIN_ROOT}/POLICY.md` (no need to open it): Never mark read, flag, accept, decline, move or delete. Replies go through `outlook-send`, never from here.
