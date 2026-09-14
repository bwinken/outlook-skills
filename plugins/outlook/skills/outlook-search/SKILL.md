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

If that fails with "running scripts is disabled on this system" (execution policy enforced by Group Policy, so `-ExecutionPolicy Bypass` is ignored), use the policy-free form, which loads the script text as a script block instead of running the file:

```
powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='${CLAUDE_PLUGIN_ROOT}/scripts'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '${CLAUDE_PLUGIN_ROOT}/scripts/Search-OutlookMail.ps1'))) [options]"
```

Use Windows paths with backslashes inside the single quotes if forward slashes are rejected. Do not try to change the machine's execution policy; that is the user's or IT's decision. See the plugin README section "Execution policy" for the AppLocker / Constrained Language case.

Options (all optional, combine freely):

| Option | Meaning |
|---|---|
| `-From "alice"` | sender name or address contains |
| `-To "bob"` | To or CC display string contains |
| `-Subject "invoice"` | subject contains |
| `-Body "PO-123"` | plain-text body contains |
| `-Text "budget"` | subject OR body contains |
| `-AnyOf "報價","quote","pricing"` | any of several terms in subject OR body (keyword expansion) |
| `-After 2026-09-01` / `-Before 2026-09-14` | received date range (Before is exclusive) |
| `-HasAttachments` | only mails with attachments |
| `-Unread` | only unread mails |
| `-Folder "Inbox/Projects"` | folder path; default Inbox. Also `"Sent Items"`, `"Deleted Items"`, `"Junk Email"`, or `"\\Store Name\Inbox\Sub"` |
| `-Store "Mailbox - Name"` | pick a specific store (shared mailbox, archive .pst) |
| `-AllFolders` | recurse every mail folder under `-Folder` |
| `-Max 50` | result cap, newest first |
| `-IncludeBody` | include full plain-text body (slower, larger) |
| `-PreviewLength 500` | length of `BodyPreview` (default 200); use ~500 for reranking |
| `-OutFile hits.json` | write JSON to file |

Text matching is case-insensitive substring. Quote values containing spaces.

## Workflow

1. Translate the user's request into options. Prefer narrow filters plus `-Max` over `-IncludeBody` on broad queries.
2. Run the script. If it errors with "Cannot start Outlook COM automation", tell the user Classic Outlook is required and suggest `outlook-status -SkipCom` to check the setup.
3. Present hits as a table: date, from, subject, folder, attachments. Keep `EntryID` handy: `outlook-thread` accepts it to open the full conversation.
4. For "read this mail in full", re-run with `-IncludeBody -Max 1` and the specific filters, or hand the EntryID to `outlook-thread`.

## Fuzzy / semantic search (reranker)

Substring matching misses typos, synonyms and mixed Chinese/English wording. Escalate in this order and stop as soon as the user has what they need:

1. **Exact**: the filters above.
2. **Keyword expansion** (no network, no consent needed): rewrite the request into several likely terms in both languages and run `-AnyOf`. Example: "上次跟供應商談價格的信" -> `-AnyOf "價格","報價","quote","quotation","pricing" -After <90 days ago> -AllFolders`.
3. **Reranker** (sends data to a gateway, **consent required**):
   - Run `python "${CLAUDE_PLUGIN_ROOT}/scripts/rerank.py" --show-config` to resolve the gateway URL and model.
   - **Ask the user before sending anything.** Tell them, in one short message: which gateway URL and model will be used, that the subject, sender, date and a ~500-character preview of about N candidate mails from the chosen date range will be sent in batches of 30, and ask whether to go ahead. Use AskUserQuestion when available. Proceed only on a clear yes. If they decline, stay with steps 1 and 2.
   - Fetch candidates broadly: `Search-OutlookMail.ps1 -After <date> -AllFolders -Max 300 -PreviewLength 500 -OutFile "<tmp>/candidates.json"` (add `-AnyOf` terms to narrow if the mailbox is large; 100 to 300 candidates is the useful range).
   - Rank: `python "${CLAUDE_PLUGIN_ROOT}/scripts/rerank.py" --query "<the user's request in their own words>" --input "<tmp>/candidates.json" --top 10`.
   - Present the top results with their `Score` (see reference.md). Scores are relative; treat anything far below the best hit as noise.
   - If `--show-config` reports no gateway, say the reranker is not configured and explain the `OUTLOOK_RERANK_URL` / `OUTLOOK_RERANK_MODEL` / `OUTLOOK_RERANK_API_KEY` settings (see plugin README). Do not guess a URL.

Never send full bodies to the gateway (the script only sends previews), and never send candidates the user has not agreed to send.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-search/reference.md` before presenting results. It documents every JSON field the script returns and the presentation template to use in the reply.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Do not open items with `Display()`, do not change `UnRead`, do not move or delete. If the user asks to act on a mail (reply, delete, flag), explain that this plugin only reads and let them do it in Outlook.
