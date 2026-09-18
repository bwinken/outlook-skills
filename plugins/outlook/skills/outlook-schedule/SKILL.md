---
name: outlook-schedule
description: Create a meeting (invitations to attendees) or an appointment on the user's own calendar in the local Windows Outlook (Classic), one of the two skills that write to Outlook. Checks the slot for overlaps, resolves attendees through Outlook, shows subject, time, location, attendees and text in chat for approval, and creates it only after the user also clicks in a confirmation window on their desktop. Use when the user says "set up a meeting with Alice on Thursday", "send an invite for 2pm", "book an hour for the Q3 review", "block my calendar Friday morning", or 約會議 / 發會議邀請 / 幫我排一個會 / 排會議 / 把週五早上擋起來 / 加一個行程. For "when am I free" use outlook-availability first.
---

# outlook-schedule

Creates calendar items. Two steps, both compulsory: the user approves the draft in chat, then clicks **送出邀請 Send** (meeting) or **建立 Create** (own appointment) in a window that `outlook_meeting.py send` opens on their own screen. Nothing else in this plugin creates or changes calendar items, and nothing skips that window.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, or Claude Code inside Claude Desktop) and offers a structured question tool (AskUserQuestion). The confirmation window appears on that machine's desktop. In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so and stop.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_meeting.py" draft -Subject "..." -Start 2026-09-18T14:00 [-End 2026-09-18T15:00 | -Duration 60] [-Attendees "alice@contoso.com" "PC Liao"] [-Optional "bob"] [-Location "Room 3 / Teams"] [-BodyFile "<tmp>/agenda.txt"] [-Store X]
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_meeting.py" send <id> -Confirm <token>
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_meeting.py" show <id> | discard <id>
```

`draft` resolves every attendee through Outlook (a name it cannot resolve is an error, never a guess), drops the user's own addresses, lists calendar items that overlap the slot in `conflicts`, builds the invitation text (agenda, then the approval footer), stores the draft under `~/.outlook-skills/drafts/` and prints it with `id` and `confirm`. It creates nothing. With attendees the draft is a **meeting** (invitations go out on Send); without, an **appointment** on the user's own calendar (Save). `send` opens the window, waits for the click, creates the item with exactly the stored values, reads them back, aborts on any difference, then sends or saves. Cancel, closing the window or the timeout (default 5 minutes) creates nothing and retires the draft. Recurring meetings, rooms as resources, Teams links and attachments are not supported; say so if asked.

## Workflow

1. **Context.** Run `settings.py show` once per conversation, in the same step as the first script (parallel tool calls; the scripts take `store` from the settings themselves). First run: offer `outlook-setup` afterwards. Names: open the memory note whose title or tag matches (`memory.py show "<title>"`, from the index in that output) for the address. Working hours from settings are the default window for "sometime Thursday".
2. **Pick the slot.** If the user gave a time, take it. If they gave a day or "next week", find free slots the way `outlook-availability` does (`outlook_calendar.py` over the range, gaps within working hours) and propose at most three, or ask. Default length: settings `meeting.default_duration_minutes` (60). Always the user's local time.
3. **Draft.** Write the agenda per reference.md §1 (short, plain, optional) to a UTF-8 file when there is one. Run `draft`. Keep the JSON.
4. **Show and ask.** Present per reference.md §2: subject, date and time with duration, location, required and optional attendees with full addresses, agenda, footer, and every entry of `conflicts` in a visible line (「⚠ 撞期」). One AskUserQuestion (reference.md §3) that repeats the time and lists every attendee, options 「送出邀請」(or 「建立行程」) / 「改時間」/ 「改與會者」/ 「改內容」/ 「取消」. Any change means a new `draft` (discard the old id) and asking again. With conflicts, recommend 「改時間」 unless the user already said to go ahead.
5. **Send.** Run `send <id> -Confirm <token>` and tell the user in one line that a confirmation window has opened on their desktop. The command blocks until they decide. Report per reference.md §4.
6. **After a cancel** do nothing more.

## Hard rules

- Never run `send` before the user chose the send / create option in step 4, and never re-run it after a cancel or timeout without a fresh draft and a fresh yes.
- Never edit a draft file by hand, never change subject, time, location, attendees or text after they were shown. Anything different goes back to step 3.
- Never invite a guessed address. A name `draft` cannot resolve goes back to the user.
- Never silently book over an existing item: conflicts are shown before the question, every time.
- The footer (`Drafted by Claude, reviewed and approved by <name>`) stays in the invitation text.

## Settings

`meeting.default_duration_minutes` (60), `meeting.reminder_minutes` (15), and the shared `send.approver`, `send.footer`, `send.dialog_timeout_seconds`. Change with `outlook-memory` or `settings.py set meeting.default_duration_minutes 30`.

## Output format

`${CLAUDE_PLUGIN_ROOT}/skills/outlook-schedule/reference.md` holds the agenda rules, the draft card, the question wording, the result lines and the draft JSON fields. Read it once per conversation, in the same step as `settings.py show`.

## Read-only rules

Everything else is read-only, per `${CLAUDE_PLUGIN_ROOT}/POLICY.md` (no need to open it). This skill's only write is the Send or Save inside `outlook_meeting.py send`, after the user's click. No accepting, declining, moving or deleting existing items.
