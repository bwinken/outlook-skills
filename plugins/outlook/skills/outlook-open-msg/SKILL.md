---
name: outlook-open-msg
description: Open and read a saved Outlook message file (.msg) or standard .eml file READ-ONLY, without needing Outlook running. Shows sender, recipients, date, subject, body, attachment list and optionally full transport headers, and can copy attachments out to a folder. Use when the user drops a .msg or .eml file, asks "what's in this email file", "check the headers of this message", "is this phishing", or 打開這個 .msg / 看這封信的內容 / 看郵件標頭 / 把附件抽出來.
---

# outlook-open-msg

Read-only parser for single-message files. The input file is never modified and Outlook is not involved.

## Where this runs

Anywhere Python runs, including Claude Desktop and claude.ai skills (the user attaches the .msg/.eml file) and Cowork. No Outlook needed.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/read_msg.py" <file.msg|file.eml> [more files...] [options]
```

Options:
- `--format json|markdown` (default json). Markdown is easier to show to the user directly.
- `--headers`: include full transport headers (Received, Return-Path, Authentication-Results, etc.).
- `--max-body N`: truncate body to N characters.
- `--extract-to DIR`: copy attachments into DIR. This writes *only* to DIR; the source file stays untouched. Ask before extracting if the user did not request it.

Requirements: Python 3 only. Both `.eml` and `.msg` are parsed with the standard library (`msgfile.py` handles the OLE2 container and MAPI properties); nothing to install.

## Workflow

1. Run with `--format markdown` for a quick read, or json when you need to post-process.
2. Present: from, to/cc, date, subject, attachment names and sizes, then the body (trim long quoted history).
3. **Always** run the quick phishing check before presenting: display name vs address, Reply-To, risky attachment types, look-alike domains, credential or payment asks. If it trips, read POLICY.md ("Phishing warnings") for the levels and the banner; the 🚨 / ⚠️ warning goes first and every link is defanged. For a phishing or delivery question, re-run with `--headers` and add the full table from reference.md §2b (SPF / DKIM / DMARC from `Authentication-Results`, the `Received` chain). Never open attachments; refuse `--extract-to` on a 🚨 mail.
4. Multiple files: pass them all at once; JSON output becomes an array.

## Settings and memory

Run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once per conversation (no Outlook access, instant), **in the same step as the first script run**, as parallel tool calls: the scripts take `store` (search also `search.default_folder` and `search.all_folders`) from the settings themselves, so nothing waits for it. A tool call that runs alone costs a whole model turn; batch the independent ones.

- `first_run: true` means `~/.outlook-skills` does not exist yet: present this result, then offer `outlook-setup` (settings wizard, mailbox scan, reply-habit profile) once per conversation; respect a "not now".
- Apply the merged `settings` (`language`).
- `memory` in that output is the complete index (title, category, tags, path). Open a note with `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" show "<title>"` only when a title or tag matches a person, folder, project or routine the request names, so "Alice" or "供應商的信" resolve to the right address or folder; no `find` first, never every note.
- If the user states something worth keeping, offer to save it through `outlook-memory`; never write memory silently.

## Output format

`${CLAUDE_PLUGIN_ROOT}/skills/outlook-open-msg/reference.md` documents every JSON field the script returns and the presentation template to use in the reply. Read it once per conversation, in the same step as the script run (parallel tool calls), not as a separate turn before answering.

## Read-only rules

Read-only, per `${CLAUDE_PLUGIN_ROOT}/POLICY.md` (no need to open it): Never edit or re-save the .msg/.eml. Never execute or open extracted attachments.
