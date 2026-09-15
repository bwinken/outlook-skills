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

Requirements: `.eml` needs only Python 3. `.msg` needs `pip install extract-msg`; if the script says so, tell the user how to install it.

## Workflow

1. Run with `--format markdown` for a quick read, or json when you need to post-process.
2. Present: from, to/cc, date, subject, attachment names and sizes, then the body (trim long quoted history).
3. **Always** run the quick phishing check (plugin README, "Phishing warnings") before presenting: display name vs address, Reply-To, risky attachment types, look-alike domains, credential or payment asks. If it trips, the 🚨 / ⚠️ warning goes first and every link is defanged. For a phishing or delivery question, re-run with `--headers` and add the full table from reference.md §2b (SPF / DKIM / DMARC from `Authentication-Results`, the `Received` chain). Never open attachments; refuse `--extract-to` on a 🚨 mail.
4. Multiple files: pass them all at once; JSON output becomes an array.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant).

- `first_run: true` means `~/.outlook-skills` does not exist yet: switch to `outlook-memory`'s onboarding, which asks the user (structured question tool) whether to create a personal memory and scan the mailbox. Respect a "not now" and continue here.
- Apply the merged `settings` (`language`).
- `memory` is an index (title, category, tags, updated, path), not the notes themselves. When the request names a person, folder, project or routine, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "<word>"` and `show` the matching note, so "Alice" or "供應商的信" resolve to the right address or folder. Do not load every note.
- If the user states something worth keeping, offer to save it through `outlook-memory`; never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-open-msg/reference.md` before presenting results. It documents every JSON field the script returns and the presentation template to use in the reply.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never edit or re-save the .msg/.eml. Never execute or open extracted attachments.
