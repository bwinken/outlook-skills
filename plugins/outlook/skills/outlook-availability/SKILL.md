---
name: outlook-availability
description: Find when the user is free, READ-ONLY, from their Outlook (Classic) calendar: open slots on a given day or across a date range, the longest gap, or a slot of a required length, within working hours. Use when the user asks "am I free Wednesday", "when do I have time this week", "find me an hour on Thursday afternoon", or 禮拜三有沒有空 / 我什麼時候有空 / 幫我找一小時的空檔 / 這週哪天下午有空. For "what meetings do I have" use outlook-agenda instead.
---

# outlook-availability

Read-only free-slot finder. It reads the calendar and computes gaps; it never books anything.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, Zoo Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line, and point the user to Claude Code, or to `outlook-open-msg` for .msg/.eml files they export from Outlook.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/run.py" Get-OutlookCalendar.ps1 [options]
```

Always go through `run.py` (from the Bash tool, or any shell): it finds Windows PowerShell 5.1 or pwsh, passes arguments without shell quoting (spaces, quotes, `$`, Chinese are safe), falls back automatically when Group Policy blocks `-ExecutionPolicy Bypass`, handles the UTF-8 BOM PowerShell 5.1 needs, and prints the script's JSON as UTF-8. Use `--out <file>` instead of `-OutFile` to keep the JSON on disk for large results. Do not run the .ps1 directly and do not change the machine's execution policy. AppLocker / Constrained Language Mode blocks COM entirely; see the plugin README.

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

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant).

- `first_run: true` means `~/.outlook-skills` does not exist yet: switch to `outlook-memory`'s onboarding, which asks the user (structured question tool) whether to create a personal memory and scan the mailbox. Respect a "not now" and continue here.
- Apply the merged `settings` (`working_hours.*`, `availability.min_slot_minutes`, `store`, `language`).
- `memory` is an index (title, category, tags, updated, path), not the notes themselves. When the request names a person, folder, project or routine, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "<word>"` and `show` the matching note, so "Alice" or "供應商的信" resolve to the right address or folder. Do not load every note.
- If the user states something worth keeping, offer to save it through `outlook-memory`; never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-availability/reference.md` before presenting results. It documents the JSON fields, the gap computation rules and the presentation template.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never create, accept, decline, move or delete appointments. Suggested times are given in chat only; the user books them in Outlook.
