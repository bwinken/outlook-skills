---
name: outlook-thread
description: Read and summarise a whole email conversation (thread) from the local Windows Outlook (Classic) mailbox, READ-ONLY. Returns every message oldest-first with full bodies so you can summarise decisions, action items, open questions and who said what. Use when the user asks "summarise this thread", "what was decided about X", "catch me up on the email about Y", or 幫我摘要這串信 / 這個討論的結論是什麼 / 整理這封信的來龍去脈.
---

# outlook-thread

Read-only conversation reader. Nothing in Outlook is modified.

## Run

```
powershell -NoProfile -ExecutionPolicy Bypass -File "${CLAUDE_PLUGIN_ROOT}/scripts/Get-OutlookThread.ps1" [selector] [options]
```

If that fails with "running scripts is disabled on this system" (execution policy enforced by Group Policy, so `-ExecutionPolicy Bypass` is ignored), use the policy-free form, which loads the script text as a script block instead of running the file:

```
powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='${CLAUDE_PLUGIN_ROOT}/scripts'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '${CLAUDE_PLUGIN_ROOT}/scripts/Get-OutlookThread.ps1'))) [selector]  [options]"
```

Use Windows paths with backslashes inside the single quotes if forward slashes are rejected. Do not try to change the machine's execution policy; that is the user's or IT's decision. See the plugin README section "Execution policy" for the AppLocker / Constrained Language case.

Selectors (use one):
- `-EntryID <id>`: precise, from `outlook-search` output.
- `-ConversationID <id>`: from `outlook-search` output.
- `-Subject "Q3 budget"`: substring; the newest matching mail in `-Folder` (default Inbox) anchors the thread.

Options:
- `-Folder`, `-Store`: where to look for the anchor when using `-Subject`.
- `-MaxBodyChars 20000`: per-message body cap.
- `-OutFile thread.json`: write to file (recommended for long threads, then read the file).

The script first uses Outlook's conversation index; for POP/PST stores without conversation support it falls back to matching `ConversationTopic` across all mail folders (`Method` field tells you which).

## Workflow

1. If the user names a subject, try `-Subject` directly. If several threads could match, run `outlook-search -Subject ... -Max 10` first and let the user pick, then use `-EntryID`.
2. Read the `Messages` array oldest-first. Quoted replies often repeat earlier text; skip repeated quoted blocks when summarising.
3. Produce, in the user's language:
   - one-paragraph summary,
   - decisions made,
   - action items with owner and due date if stated,
   - open questions / who is waiting on whom,
   - participants (`Participants` field).
4. Cite messages by date and sender, not by index.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant). Apply the merged `settings` (`search.default_folder`, `store`, `language`) and read every `memory` file it lists: memory.md holds the user's contact aliases, folder meanings, project keywords and preferences, so "Alice" or "供應商的信" may already be defined there. If the user tells you something worth keeping, offer to save it with `outlook-settings`; do not write memory silently. Details: `${CLAUDE_PLUGIN_ROOT}/skills/outlook-settings/reference.md`.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-thread/reference.md` before presenting results. It documents every JSON field the script returns and the presentation template to use in the reply.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never reply, forward, flag, mark read, move or delete. If the user wants a reply drafted, write the draft text in chat for them to paste into Outlook.
