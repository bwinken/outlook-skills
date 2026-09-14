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

Substring matching misses typos, synonyms and mixed Chinese/English wording. Escalate in this order and stop as soon as the user has what they need. **Always narrow first; the reranker is only for when the narrowed set is still too big to read.**

1. **Narrow with what is certain.** Turn every hard fact in the request into filters before any fuzzy step: sender or recipient, date range (default the last 90 days when the user says "recently" or gives no hint), folder, attachments. Run that search with `-Max 300 -PreviewLength 500 -OutFile "<tmp>/candidates.json"`.
2. **Keyword expansion** when the topic is fuzzy (no network, no consent needed): rewrite the topic into several likely terms in both languages and add `-AnyOf`. Example: "上次跟供應商談價格的信" -> `-AnyOf "價格","報價","quote","quotation","pricing"`.
3. **Decide by count** (`Count` in the JSON). The reranker is never the default; it is one branch of this ladder:

   | Count | What to do |
   |---|---|
   | 0 | Relax one filter (wider dates, `-AllFolders`, fewer terms) and retry once. Then report what was tried. |
   | 1 to `search.direct_read_max` (default 20) | Read directly. Show the list, or re-run with `-IncludeBody` on the same filters and answer from the bodies. No reranker. |
   | `direct_read_max`+1 to 100 | Reranker **if** it is usable and the user agrees (step 4, or `rerank.auto_consent` is true). Otherwise **local scan** (step 5). |
   | over 100 | Narrow first: add a sender, tighter dates or more `-AnyOf` terms, or ask the user for one more constraint. If it cannot be narrowed, reranker if usable and agreed; otherwise local scan of the newest 100. |

4. **Reranker** (sends data to a gateway, **consent required**):
   - Run `python "${CLAUDE_PLUGIN_ROOT}/scripts/rerank.py" --show-config`. It resolves the gateway (by default the `ANTHROPIC_BASE_URL` from `~/.claude/settings.json`, overridable with `OUTLOOK_RERANK_URL`), then probes `/v1/rerank` and `/v1/score` with a one-word test request and reports `usable`, `endpoint` and `gateway`. Exit code 1 means no usable reranker: say so in one line, mention the settings names (see plugin README), and go to step 5. Do not guess a URL.
   - **Ask the user before sending anything** unless `rerank.auto_consent` is true in the settings (then state in one line what is being sent and where, and continue). One short message: the gateway URL and model from `--show-config`, that subject, sender, date and a ~500-character preview of the `Count` candidates will be sent in batches of 30, and whether to go ahead. Ask with the host's structured question tool (Claude Code: AskUserQuestion; Roo Code: ask_followup_question), never as plain text. Proceed only on a clear yes; a no goes to step 5.
   - Rank: `python "${CLAUDE_PLUGIN_ROOT}/scripts/rerank.py" --query "<the user's request in their own words>" --input "<tmp>/candidates.json" --top 10`.
   - If the script exits non-zero mid-run (gateway error, timeout after its built-in retries, unexpected response), do not retry by hand: tell the user in one line which step failed and go to step 5.
   - Present the top results with their `Score` (see reference.md). Scores are relative; treat anything far below the best hit as noise. Offer to open the best hits with `-IncludeBody` or `outlook-thread`.

5. **Local scan** (the fallback; no network, nothing leaves the machine): read `<tmp>/candidates.json` yourself and judge relevance from `Subject`, `From` and `BodyPreview`. Pick up to 10 that match the request, newest first among equals. Present them with the 2a table and say plainly that no reranker was used and the pick is your own reading of the previews (see reference.md §4). If the previews are not enough to tell, re-run the search with `-IncludeBody` for the 5 most likely and read those. Then, if the answer is still uncertain, ask the user for one more constraint rather than guessing.

Never send full bodies to the gateway (the script only sends previews), and never send candidates the user has not agreed to send.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant).

- `first_run: true` means `~/.outlook-skills` does not exist yet: switch to `outlook-memory`'s onboarding, which asks the user (structured question tool) whether to create a personal memory and scan the mailbox. Respect a "not now" and continue here.
- Apply the merged `settings` (`search.*` for default folder, lookback window, candidate cap and the direct-read threshold; `store`; `rerank.*` including `auto_consent`, which replaces the per-run consent question when true; `language`).
- `memory` is an index (title, category, tags, updated, path), not the notes themselves. When the request names a person, folder, project or routine, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "<word>"` and `show` the matching note, so "Alice" or "供應商的信" resolve to the right address or folder. Do not load every note.
- If the user states something worth keeping, offer to save it through `outlook-memory`; never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-search/reference.md` before presenting results. It documents every JSON field the script returns and the presentation template to use in the reply.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Do not open items with `Display()`, do not change `UnRead`, do not move or delete. If the user asks to act on a mail (reply, delete, flag), explain that this plugin only reads and let them do it in Outlook.
