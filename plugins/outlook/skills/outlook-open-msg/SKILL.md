---
name: outlook-open-msg
description: Open and read a saved Outlook message file (.msg) or standard .eml file READ-ONLY, without needing Outlook running. Shows sender, recipients, date, subject, body, attachment list and optionally full transport headers, and can copy attachments out to a folder. Use when the user drops a .msg or .eml file, asks "what's in this email file", "check the headers of this message", "is this phishing", or 打開這個 .msg / 看這封信的內容 / 看郵件標頭 / 把附件抽出來.
---

# outlook-open-msg

Read-only parser for single-message files. The input file is never modified and Outlook is not involved.

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
3. For a phishing or delivery question, re-run with `--headers` and check: sender address vs display name, Reply-To differing from From, `Authentication-Results` (SPF / DKIM / DMARC), the chain of `Received` hops, and attachment types. Report findings; do not open attachments.
4. Multiple files: pass them all at once; JSON output becomes an array.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never edit or re-save the .msg/.eml. Never execute or open extracted attachments.
