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
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_thread.py" [selector] [options]
```

Pure Python (needs `pip install pywin32` once on the Windows machine). Options accept PowerShell-style `-From` or `--from` spellings. Run it from any tool (Bash, cmd, PowerShell); nothing goes through a shell that could mangle quotes, `$` or Chinese. Output is UTF-8 JSON on stdout; `-OutFile <file>` writes it to a file instead (use for large results). If it reports that pywin32 is missing or that Outlook COM cannot start, say so and point to the plugin README requirements.

Selectors (use one):
- `-EntryID <id>`: precise, from `outlook-search` output.
- `-ConversationID <id>`: from `outlook-search` output.
- `-Subject "Q3 budget"`: substring; the newest matching mail in `-Folder` (default Inbox) anchors the thread.

Options:
- `-Folder`, `-Store`: where to look for the anchor when using `-Subject`.
- `-MaxBodyChars 20000`: per-message body cap.
- `-OutFile thread.json`: write to a file (recommended for long threads, then read the file, or delegate the reading; see "Delegating heavy reads").

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

## Delegating heavy reads

When the host offers subagents (Claude Code's Agent tool, including the Code tab in Claude Desktop) and the thread has more than about 15 messages or the JSON exceeds ~100 KB, hand the reading to a subagent so the raw data never enters this conversation. Give it: the path of `thread.json`, the user's actual question (or "full summary"), matching memory notes for the participants, and the template in reference.md §2 with its rules about quoted history and attribution. Ask it to return the filled template (summary, decisions, action items table, open questions, timeline, participants), every claim attributed to sender and date and nothing else, no raw records. Anything that needs the user's consent (writing memory, sending candidates to the reranker, copying attachments out) stays in this conversation; a subagent never asks the user and never writes. Without subagents, do the same work here but read only what the step needs.

## Settings and memory

Run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once per conversation (no Outlook access, instant), **in the same step as the first script run**, as parallel tool calls: the scripts take `store` (search also `search.default_folder` and `search.all_folders`) from the settings themselves, so nothing waits for it. A tool call that runs alone costs a whole model turn; batch the independent ones.

- `first_run: true` means `~/.outlook-skills` does not exist yet: present this result, then offer `outlook-setup` (settings wizard, mailbox scan, reply-habit profile) once per conversation; respect a "not now".
- Apply the merged `settings` (`search.default_folder`, `store`, `language`).
- `memory` in that output is the complete index (title, category, tags, path). Open a note with `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" show "<title>"` only when a title or tag matches a person, folder, project or routine the request names, so "Alice" or "供應商的信" resolve to the right address or folder; no `find` first, never every note.
- If the user states something worth keeping, offer to save it through `outlook-memory`; never write memory silently.

## Output format

`${CLAUDE_PLUGIN_ROOT}/skills/outlook-thread/reference.md` documents every JSON field the script returns and the presentation template to use in the reply. Read it once per conversation, in the same step as the script run (parallel tool calls), not as a separate turn before answering.

## Read-only rules

Read-only, per `${CLAUDE_PLUGIN_ROOT}/POLICY.md` (no need to open it): Never reply, forward, flag, mark read, move or delete. If the user wants to reply, hand over to `outlook-send` with the `EntryID`.
