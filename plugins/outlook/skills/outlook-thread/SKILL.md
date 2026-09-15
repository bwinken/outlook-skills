---
name: outlook-thread
description: Read and summarise a whole email conversation (thread) from the local Windows Outlook (Classic) mailbox, READ-ONLY. Returns every message oldest-first with full bodies so you can summarise decisions, action items, open questions and who said what. Use when the user asks "summarise this thread", "what was decided about X", "catch me up on the email about Y", or 幫我摘要這串信 / 這個討論的結論是什麼 / 整理這封信的來龍去脈.
---

# outlook-thread

Read-only conversation reader. Nothing in Outlook is modified.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, Zoo Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line, and point the user to Claude Code, or to `outlook-open-msg` for .msg/.eml files they export from Outlook.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/run.py" Get-OutlookThread.ps1 [selector] [options]
```

Always go through `run.py` (from the Bash tool, or any shell): it finds Windows PowerShell 5.1 or pwsh, passes arguments without shell quoting (spaces, quotes, `$`, Chinese are safe), falls back automatically when Group Policy blocks `-ExecutionPolicy Bypass`, handles the UTF-8 BOM PowerShell 5.1 needs, and prints the script's JSON as UTF-8. Use `--out <file>` instead of `-OutFile` to keep the JSON on disk for large results. Do not run the .ps1 directly and do not change the machine's execution policy. AppLocker / Constrained Language Mode blocks COM entirely; see the plugin README.

Selectors (use one):
- `-EntryID <id>`: precise, from `outlook-search` output.
- `-ConversationID <id>`: from `outlook-search` output.
- `-Subject "Q3 budget"`: substring; the newest matching mail in `-Folder` (default Inbox) anchors the thread.

Options:
- `-Folder`, `-Store`: where to look for the anchor when using `-Subject`.
- `-MaxBodyChars 20000`: per-message body cap.
- `--out thread.json`: keep the JSON in a file (recommended for long threads, then read the file).

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

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant).

- `first_run: true` means `~/.outlook-skills` does not exist yet: switch to `outlook-memory`'s onboarding, which asks the user (structured question tool) whether to create a personal memory and scan the mailbox. Respect a "not now" and continue here.
- Apply the merged `settings` (`search.default_folder`, `store`, `language`).
- `memory` is an index (title, category, tags, updated, path), not the notes themselves. When the request names a person, folder, project or routine, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "<word>"` and `show` the matching note, so "Alice" or "供應商的信" resolve to the right address or folder. Do not load every note.
- If the user states something worth keeping, offer to save it through `outlook-memory`; never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-thread/reference.md` before presenting results. It documents every JSON field the script returns and the presentation template to use in the reply.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never reply, forward, flag, mark read, move or delete. If the user wants a reply drafted, write the draft text in chat for them to paste into Outlook.
