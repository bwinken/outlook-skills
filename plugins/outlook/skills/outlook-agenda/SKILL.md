---
name: outlook-agenda
description: List what is on the user's Outlook (Classic) calendar READ-ONLY for today, a number of days, or a date range, with recurring meetings expanded, overlapping meetings flagged, and unanswered invitations called out. Use when the user asks "what's on my calendar", "meetings this week", "do I have conflicts", "which invites haven't I answered", or 今天有什麼會議 / 這週行程 / 會議有沒有撞期 / 哪場還沒回覆. For "when am I free" questions use outlook-availability instead.
---

# outlook-agenda

Read-only agenda listing. Nothing in Outlook is modified.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, Zoo Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line, and point the user to Claude Code, or to `outlook-open-msg` for .msg/.eml files they export from Outlook.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_calendar.py" [options]
```

Pure Python (needs `pip install pywin32` once on the Windows machine). Options accept PowerShell-style `-From` or `--from` spellings. Run it from any tool (Bash, cmd, PowerShell); nothing goes through a shell that could mangle quotes, `$` or Chinese. Output is UTF-8 JSON on stdout; `-OutFile <file>` writes it to a file instead (use for large results). If it reports that pywin32 is missing or that Outlook COM cannot start, say so and point to the plugin README requirements.

Options:
- default: today (local time).
- `-Days 7`: from today for N days.
- `-Start 2026-09-15 -End 2026-09-20`: explicit range (End exclusive). `-Start` with `-Days` also works.
- `-Store "Mailbox - Name"`: another mailbox's default calendar (needs read permission in Outlook).
- `-IncludeFree`: include items marked Free (reminders, tentative holds). Off by default.
- `-OutFile cal.json`: write to a file.

## Workflow

1. Pick the range from the request. "This week" = Monday to next Monday; "tomorrow" = a one-day range starting tomorrow.
2. Present an agenda grouped by day (see reference.md): time, subject, location, organizer, response status. All-day items first.
3. Call out, after the agenda: overlapping meetings (`Conflicts`), invitations still `NotResponded`, and cancelled meetings still on the calendar.
4. To prepare for a meeting, take attendee names from an item and hand them to `outlook-search -From ...` to gather recent mails.
5. If the user asks when they are free, switch to `outlook-availability`.

## Settings and memory

Run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once per conversation (no Outlook access, instant), **in the same step as the first script run**, as parallel tool calls: the scripts take `store` (search also `search.default_folder` and `search.all_folders`) from the settings themselves, so nothing waits for it. A tool call that runs alone costs a whole model turn; batch the independent ones.

- `first_run: true` means `~/.outlook-skills` does not exist yet: present this result, then offer `outlook-setup` (settings wizard, mailbox scan, reply-habit profile) once per conversation; respect a "not now".
- Apply the merged `settings` (`store`, `language`).
- `memory` in that output is the complete index (title, category, tags, path). Open a note with `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" show "<title>"` only when a title or tag matches a person, folder, project or routine the request names, so "Alice" or "供應商的信" resolve to the right address or folder; no `find` first, never every note.
- If the user states something worth keeping, offer to save it through `outlook-memory`; never write memory silently.

## Output format

`${CLAUDE_PLUGIN_ROOT}/skills/outlook-agenda/reference.md` documents every JSON field the script returns and the presentation template to use in the reply. Read it once per conversation, in the same step as the script run (parallel tool calls), not as a separate turn before answering.

## Read-only rules

Read-only, per `${CLAUDE_PLUGIN_ROOT}/POLICY.md` (no need to open it): Never create, accept, decline, move or delete appointments.
