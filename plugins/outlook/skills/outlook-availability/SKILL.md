---
name: outlook-availability
description: Find when the user is free, READ-ONLY, from their Outlook (Classic) calendar: open slots on a given day or across a date range, the longest gap, or a slot of a required length, within working hours. Use when the user asks "am I free Wednesday", "when do I have time this week", "find me an hour on Thursday afternoon", or 禮拜三有沒有空 / 我什麼時候有空 / 幫我找一小時的空檔 / 這週哪天下午有空. For "what meetings do I have" use outlook-agenda instead.
---

# outlook-availability

Read-only free-slot finder. It reads the calendar and computes gaps; it never books anything. When the user picks a slot and wants it booked, hand over to `outlook-schedule`.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, Zoo Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line, and point the user to Claude Code, or to `outlook-open-msg` for .msg/.eml files they export from Outlook.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_calendar.py" [options]
```

Pure Python (needs `pip install pywin32` once on the Windows machine). Options accept PowerShell-style `-From` or `--from` spellings. Run it from any tool (Bash, cmd, PowerShell); nothing goes through a shell that could mangle quotes, `$` or Chinese. Output is UTF-8 JSON on stdout; `-OutFile <file>` writes it to a file instead (use for large results). If it reports that pywin32 is missing or that Outlook COM cannot start, say so and point to the plugin README requirements.

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

Run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once per conversation (no Outlook access, instant), **in the same step as the first script run**, as parallel tool calls: the scripts take `store` (search also `search.default_folder` and `search.all_folders`) from the settings themselves, so nothing waits for it. A tool call that runs alone costs a whole model turn; batch the independent ones.

- `first_run: true` means `~/.outlook-skills` does not exist yet: present this result, then offer `outlook-setup` (settings wizard, mailbox scan, reply-habit profile) once per conversation; respect a "not now".
- Apply the merged `settings` (`working_hours.*`, `availability.min_slot_minutes`, `store`, `language`).
- `memory` in that output is the complete index (title, category, tags, path). Open a note with `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" show "<title>"` only when a title or tag matches a person, folder, project or routine the request names, so "Alice" or "供應商的信" resolve to the right address or folder; no `find` first, never every note.
- If the user states something worth keeping, offer to save it through `outlook-memory`; never write memory silently.

## Output format

`${CLAUDE_PLUGIN_ROOT}/skills/outlook-availability/reference.md` documents the JSON fields, the gap computation rules and the presentation template. Read it once per conversation, in the same step as the script run (parallel tool calls), not as a separate turn before answering.

## Read-only rules

Read-only, per `${CLAUDE_PLUGIN_ROOT}/POLICY.md` (no need to open it): Never create, accept, decline, move or delete appointments. Suggested times are given in chat only; the user books them in Outlook.
