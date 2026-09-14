---
name: outlook-availability
description: Find when the user is free, READ-ONLY, from their Outlook (Classic) calendar: open slots on a given day or across a date range, the longest gap, or a slot of a required length, within working hours. Use when the user asks "am I free Wednesday", "when do I have time this week", "find me an hour on Thursday afternoon", or 禮拜三有沒有空 / 我什麼時候有空 / 幫我找一小時的空檔 / 這週哪天下午有空. For "what meetings do I have" use outlook-agenda instead.
---

# outlook-availability

Read-only free-slot finder. It reads the calendar and computes gaps; it never books anything.

## Run

```
powershell -NoProfile -ExecutionPolicy Bypass -File "${CLAUDE_PLUGIN_ROOT}/scripts/Get-OutlookCalendar.ps1" [options]
```

If that fails with "running scripts is disabled on this system" (execution policy enforced by Group Policy, so `-ExecutionPolicy Bypass` is ignored), use the policy-free form, which loads the script text as a script block instead of running the file:

```
powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='${CLAUDE_PLUGIN_ROOT}/scripts'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '${CLAUDE_PLUGIN_ROOT}/scripts/Get-OutlookCalendar.ps1'))) [options]"
```

Use Windows paths with backslashes inside the single quotes if forward slashes are rejected. Do not try to change the machine's execution policy; that is the user's or IT's decision. See the plugin README section "Execution policy" for the AppLocker / Constrained Language case.

Options (same script as outlook-agenda):
- default: today. `-Start 2026-09-16 -Days 1` for one day; `-Start ... -End ...` for a range (End exclusive); `-Days 7` for the coming week.
- `-Store "Mailbox - Name"`: another calendar the user can read (a colleague or room mailbox), for "when is X free" questions.
- Do not pass `-IncludeFree`: items marked Free must not block a slot.

## Workflow

1. Turn the request into a date range and constraints:
   - "禮拜三" = the next Wednesday from today (if today is Wednesday, ask whether they mean today or next week only if it matters; default today).
   - "下午" = 13:00–18:00; "早上" = 09:00–12:00. Working hours default 09:00–18:00 local; use the user's if stated.
   - A required length ("找一小時") becomes a minimum slot size.
2. Run the script for that range.
3. Compute the gaps (see reference.md for the exact rules): busy = every returned item except all-day items with BusyStatus other than OutOfOffice; Tentative counts as busy but is noted; merge overlaps; clip to working hours; drop gaps shorter than 30 minutes or the requested length.
4. Present the slots table first, then the day's meetings for context (see reference.md). Recommend one slot when the user asked for one.
5. If the range has no usable slot, say so and show the nearest options just outside working hours or on the next day.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-availability/reference.md` before presenting results. It documents the JSON fields, the gap computation rules and the presentation template.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never create, accept, decline, move or delete appointments. Suggested times are given in chat only; the user books them in Outlook.
