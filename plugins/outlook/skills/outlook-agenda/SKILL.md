---
name: outlook-agenda
description: List what is on the user's Outlook (Classic) calendar READ-ONLY for today, a number of days, or a date range, with recurring meetings expanded, overlapping meetings flagged, and unanswered invitations called out. Use when the user asks "what's on my calendar", "meetings this week", "do I have conflicts", "which invites haven't I answered", or 今天有什麼會議 / 這週行程 / 會議有沒有撞期 / 哪場還沒回覆. For "when am I free" questions use outlook-availability instead.
---

# outlook-agenda

Read-only agenda listing. Nothing in Outlook is modified.

## Run

```
powershell -NoProfile -ExecutionPolicy Bypass -File "${CLAUDE_PLUGIN_ROOT}/scripts/Get-OutlookCalendar.ps1" [options]
```

If that fails with "running scripts is disabled on this system" (execution policy enforced by Group Policy, so `-ExecutionPolicy Bypass` is ignored), use the policy-free form, which loads the script text as a script block instead of running the file:

```
powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='${CLAUDE_PLUGIN_ROOT}/scripts'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '${CLAUDE_PLUGIN_ROOT}/scripts/Get-OutlookCalendar.ps1'))) [options]"
```

Use Windows paths with backslashes inside the single quotes if forward slashes are rejected. Do not try to change the machine's execution policy; that is the user's or IT's decision. See the plugin README section "Execution policy" for the AppLocker / Constrained Language case.

Options:
- default: today (local time).
- `-Days 7`: from today for N days.
- `-Start 2026-09-15 -End 2026-09-20`: explicit range (End exclusive). `-Start` with `-Days` also works.
- `-Store "Mailbox - Name"`: another mailbox's default calendar (needs read permission in Outlook).
- `-IncludeFree`: include items marked Free (reminders, tentative holds). Off by default.
- `-OutFile cal.json`

## Workflow

1. Pick the range from the request. "This week" = Monday to next Monday; "tomorrow" = a one-day range starting tomorrow.
2. Present an agenda grouped by day (see reference.md): time, subject, location, organizer, response status. All-day items first.
3. Call out, after the agenda: overlapping meetings (`Conflicts`), invitations still `NotResponded`, and cancelled meetings still on the calendar.
4. For meeting prep, take attendee names from an item and hand them to `outlook-search -From ...` to gather recent mails.
5. If the user asks when they are free, switch to `outlook-availability`.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant). Apply the merged `settings` (`store`, `language`) and read every `memory` file it lists: memory.md holds the user's contact aliases, folder meanings, project keywords and preferences, so "Alice" or "供應商的信" may already be defined there. If the user tells you something worth keeping, offer to save it with `outlook-settings`; do not write memory silently. Details: `${CLAUDE_PLUGIN_ROOT}/skills/outlook-settings/reference.md`.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-agenda/reference.md` before presenting results. It documents every JSON field the script returns and the presentation template to use in the reply.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never create, accept, decline, move or delete appointments.
