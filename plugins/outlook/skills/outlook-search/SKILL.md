---
name: outlook-search
description: Search the local Windows Outlook (Classic) mailbox READ-ONLY by sender, recipient, subject, body text, date range, unread state, or attachments, across one folder or the whole store. Use when the user asks to find emails, "who sent me...", "emails from X last week", "unread mails about Y", "mails with attachments", or 找信 / 搜尋郵件 / 某人寄的信 / 未讀郵件.
---

# outlook-search

Read-only mail search through Outlook COM automation. Results are returned as JSON; nothing is marked read, moved or changed.

## Run

```
powershell -NoProfile -ExecutionPolicy Bypass -File "${CLAUDE_PLUGIN_ROOT}/scripts/Search-OutlookMail.ps1" [options]
```

Options (all optional, combine freely):

| Option | Meaning |
|---|---|
| `-From "alice"` | sender name or address contains |
| `-To "bob"` | To or CC display string contains |
| `-Subject "invoice"` | subject contains |
| `-Body "PO-123"` | plain-text body contains |
| `-Text "budget"` | subject OR body contains |
| `-After 2026-09-01` / `-Before 2026-09-14` | received date range (Before is exclusive) |
| `-HasAttachments` | only mails with attachments |
| `-Unread` | only unread mails |
| `-Folder "Inbox/Projects"` | folder path; default Inbox. Also `"Sent Items"`, `"Deleted Items"`, `"Junk Email"`, or `"\\Store Name\Inbox\Sub"` |
| `-Store "Mailbox - Name"` | pick a specific store (shared mailbox, archive .pst) |
| `-AllFolders` | recurse every mail folder under `-Folder` |
| `-Max 50` | result cap, newest first |
| `-IncludeBody` | include full plain-text body (slower, larger) |
| `-OutFile hits.json` | write JSON to file |

Text matching is case-insensitive substring. Quote values containing spaces.

## Workflow

1. Translate the user's request into options. Prefer narrow filters plus `-Max` over `-IncludeBody` on broad queries.
2. Run the script. If it errors with "Cannot start Outlook COM automation", tell the user Classic Outlook is required and suggest `outlook-status -SkipCom` to check the setup.
3. Present hits as a table: date, from, subject, folder, attachments. Keep `EntryID` handy: `outlook-thread` accepts it to open the full conversation.
4. For "read this mail in full", re-run with `-IncludeBody -Max 1` and the specific filters, or hand the EntryID to `outlook-thread`.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Do not open items with `Display()`, do not change `UnRead`, do not move or delete. If the user asks to act on a mail (reply, delete, flag), explain that this plugin only reads and let them do it in Outlook.
