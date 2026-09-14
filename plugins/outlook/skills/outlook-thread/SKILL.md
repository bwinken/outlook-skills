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

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never reply, forward, flag, mark read, move or delete. If the user wants a reply drafted, write the draft text in chat for them to paste into Outlook.
