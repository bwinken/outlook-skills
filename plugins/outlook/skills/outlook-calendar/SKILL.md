---
name: outlook-calendar
description: Show the user's Outlook (Classic) calendar READ-ONLY for today, a number of days, or a date range, with recurring meetings expanded and overlapping meetings flagged. Use when the user asks "what's on my calendar", "am I free Thursday afternoon", "meetings this week", "do I have conflicts", or 今天有什麼會議 / 這週行程 / 我什麼時候有空 / 會議有沒有撞期.
---

# outlook-calendar

Read-only calendar listing. Nothing in Outlook is modified.

## Run

```
powershell -NoProfile -ExecutionPolicy Bypass -File "${CLAUDE_PLUGIN_ROOT}/scripts/Get-OutlookCalendar.ps1" [options]
```

Options:
- default: today (local time).
- `-Days 7`: from today for N days.
- `-Start 2026-09-15 -End 2026-09-20`: explicit range (End exclusive). `-Start` with `-Days` also works.
- `-Store "Mailbox - Name"`: another mailbox's default calendar (needs read permission in Outlook).
- `-IncludeFree`: include items marked Free (reminders, tentative holds). Off by default.
- `-OutFile cal.json`

## What the JSON contains

- `Items`: sorted by start. Each has `Start`, `End`, `DurationMinutes`, `Subject`, `Location`, `Organizer`, `RequiredAttendees`, `OptionalAttendees`, `BusyStatus`, `MeetingStatus`, `ResponseStatus`, `IsRecurring`, `AllDayEvent`, `BodyPreview`.
- `Conflicts`: pairs of timed items that overlap.
- `Range.Filter`: the Jet filter used (helpful when debugging locale date formats).

## Workflow

1. Pick the range from the request. "This week" = Monday to next Monday; "tomorrow afternoon" = a one-day range, then filter by time in your answer.
2. Present an agenda grouped by day: time, subject, location, organizer, your response status. Mark conflicts and meetings still `NotResponded`.
3. For "when am I free": compute gaps between busy items inside working hours (assume 09:00 to 18:00 local unless the user says otherwise) and list them.
4. For meeting prep, take attendee names from an item and hand them to `outlook-search -From ...` to gather recent mails.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never create, accept, decline, move or delete appointments. Suggested times are given in chat only.
