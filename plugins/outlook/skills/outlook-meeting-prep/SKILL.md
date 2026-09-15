---
name: outlook-meeting-prep
description: Prepare for a meeting READ-ONLY from the local Outlook (Classic) mailbox - find the appointment, list attendees, pull the recent mails exchanged with each attendee and mails about the meeting's subject, and collect the attachments seen along the way, then brief the user. Use when the user asks "prep me for my next meeting", "what do I need to know before the 2pm with Cassie", "background for the vendor review", or 幫我準備下一場會議 / 開會前先看一下跟他們的往來 / 這場會議相關的信和附件.
---

# outlook-meeting-prep

Read-only. One script call gathers everything; the skill turns it into a briefing.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line and point the user to Claude Code.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_meeting_prep.py" -Next                    # next upcoming meeting
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_meeting_prep.py" -Subject "供應商" -Days 30
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_meeting_prep.py" -EntryID <appointment id> -AllStores -OutFile "<tmp>/prep.json"
```

Options: `-Horizon N` days ahead to look for the meeting (default 7); `-Days N` how far back to look for mail (default 30); `-Max N` mails per attendee (default 10); `-Store` / `-AllStores`.

## Workflow

1. Identify the meeting: subject words the user gave, or `-Next`. If several match, run `outlook-agenda` for the day and ask which one.
2. Run the script with `-AllStores` when mail lives in a PST. Use `-OutFile` and read the file; the output can be large.
3. Resolve attendees against memory (`memory.py find <surname>`) so roles are known (法務窗口, 供應商 PM ...).
4. Brief per reference.md: purpose and logistics, per-attendee "what is open with them", the thread about the subject, attachments worth opening beforehand, and open questions to raise. Quote decisions and dates from the mails; do not invent.
5. Offer read-only follow-ups: open a specific thread, extract an attachment with `outlook-search` (attachments, `-SaveTo`) (ask first), or draft talking points in chat.

## Delegating heavy reads

When the host offers subagents (Claude Code's Agent tool, including the Code tab in Claude Desktop) and the output has many attendees or more than about 40 mails in total, hand the reading to a subagent so the raw data never enters this conversation. Give it: the path of `prep.json`, the meeting subject, matching memory notes for the attendees, and the briefing template in reference.md §2. Ask it to return the filled briefing, with dates and senders on every claim and nothing else, no raw records. Anything that needs the user's consent (writing memory, sending candidates to the reranker, copying attachments out) stays in this conversation; a subagent never asks the user and never writes. Without subagents, do the same work here but read only what the step needs.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant). `first_run: true` means `~/.outlook-skills` does not exist yet: hand over to `outlook-setup` first (respect a "not now"). Apply `store` (pass it as `-Store`), `language`, and any search defaults. When the request names a person, folder or project, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "<word>"` and `show` the note so aliases resolve. Never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-meeting-prep/reference.md` for the JSON fields and the briefing template.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never reply, forward, flag, mark read, move or delete. Drafts, if asked for, are written in chat only.

