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
python "${CLAUDE_PLUGIN_ROOT}/scripts/run.py" Get-OutlookCalendar.ps1 [options]
```

Always go through `run.py` (from the Bash tool, or any shell): it finds Windows PowerShell 5.1 or pwsh, passes arguments without shell quoting (spaces, quotes, `$`, Chinese are safe), falls back automatically when Group Policy blocks `-ExecutionPolicy Bypass`, handles the UTF-8 BOM PowerShell 5.1 needs, and prints the script's JSON as UTF-8. Use `--out <file>` instead of `-OutFile` to keep the JSON on disk for large results. Do not run the .ps1 directly and do not change the machine's execution policy. AppLocker / Constrained Language Mode blocks COM entirely; see the plugin README.

Options:
- default: today (local time).
- `-Days 7`: from today for N days.
- `-Start 2026-09-15 -End 2026-09-20`: explicit range (End exclusive). `-Start` with `-Days` also works.
- `-Store "Mailbox - Name"`: another mailbox's default calendar (needs read permission in Outlook).
- `-IncludeFree`: include items marked Free (reminders, tentative holds). Off by default.
- `--out cal.json`: keep the JSON in a file.

## Workflow

1. Pick the range from the request. "This week" = Monday to next Monday; "tomorrow" = a one-day range starting tomorrow.
2. Present an agenda grouped by day (see reference.md): time, subject, location, organizer, response status. All-day items first.
3. Call out, after the agenda: overlapping meetings (`Conflicts`), invitations still `NotResponded`, and cancelled meetings still on the calendar.
4. For meeting prep, take attendee names from an item and hand them to `outlook-search -From ...` to gather recent mails.
5. If the user asks when they are free, switch to `outlook-availability`.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant).

- `first_run: true` means `~/.outlook-skills` does not exist yet: switch to `outlook-memory`'s onboarding, which asks the user (structured question tool) whether to create a personal memory and scan the mailbox. Respect a "not now" and continue here.
- Apply the merged `settings` (`store`, `language`).
- `memory` is an index (title, category, tags, updated, path), not the notes themselves. When the request names a person, folder, project or routine, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "<word>"` and `show` the matching note, so "Alice" or "供應商的信" resolve to the right address or folder. Do not load every note.
- If the user states something worth keeping, offer to save it through `outlook-memory`; never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-agenda/reference.md` before presenting results. It documents every JSON field the script returns and the presentation template to use in the reply.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never create, accept, decline, move or delete appointments.
