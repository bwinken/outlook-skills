---
name: outlook-send
description: Send a new mail or reply to one from the local Windows Outlook (Classic) mailbox, the only skill that writes to Outlook. Drafts short, plain mails in the user's own style (profile.md), shows the exact To, Cc, Subject and text in chat for approval, and sends only after the user also clicks Send in a confirmation window on their desktop. Use when the user says "reply to this", "send Alice a mail", "tell them we accept", "answer PC's question", or 回信 / 回覆這封 / 幫我回他 / 寄信給 / 寄一封信 / 跟他說.
---

# outlook-send

The one skill that sends. Two steps, both compulsory: the user approves the draft in chat, then clicks **寄出 Send** in a window that `outlook_send.py send` opens on their own screen. Nothing else in this plugin sends, and nothing can skip that window.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, or Claude Code inside Claude Desktop) and offers a structured question tool (AskUserQuestion). The confirmation window appears on that machine's desktop. In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so and offer to write the text for the user to paste.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_send.py" draft -To "alice@contoso.com" [-Cc "bob"] -Subject "..." -BodyFile "<tmp>/body.txt"
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_send.py" draft -ReplyTo <EntryID> [-ReplyAll] [-Cc "..."] -BodyFile "<tmp>/body.txt"
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_send.py" send <id> -Confirm <token>
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_send.py" show <id> | list | discard <id>
```

`draft` resolves every recipient through Outlook (a name it cannot resolve is an error, never a guess), builds the complete outgoing text (your body, then the approval footer, then the quoted original for replies), stores it under `~/.outlook-skills/drafts/` and prints it with `id` and `confirm`. It sends nothing. `send` opens the window, waits for the click, creates the item with exactly the stored recipients, subject and text, checks the item against the draft, and only then calls Send. Cancel, closing the window or the timeout (default 5 minutes) sends nothing and retires the draft. Attachments are not supported; say so if asked.

## Workflow

1. **Context.** Run `settings.py show` once (first run: hand over to `outlook-setup`). If `profile` is set, run `settings.py profile show`: it holds the user's greetings, closings, signature, typical length and language. For names, `memory.py find "<name>"` gives the address. For a reply, read the mail first: `outlook_search.py -IncludeBody -Max 1` with precise filters, or `outlook_thread.py -EntryID`, and keep the `EntryID`.
2. **Write the body** per reference.md §1: the way the user writes mail, short, the point first, one line per point, greeting and closing from the profile, the user's language. No essays, no bullet walls, no restating the original mail. Write it to a UTF-8 file.
3. **Draft.** Run `draft` with `-BodyFile`. New mail: `-To` and `-Subject` are required. Reply: `-ReplyTo <EntryID>`, `-ReplyAll` only when the user asked to answer everyone or the original clearly went to a group. Keep the JSON it prints.
4. **Show and ask.** Present the draft exactly as reference.md §2: To and Cc with full addresses, Subject, the body verbatim, the footer line, and for a reply one line noting the quoted original is attached below. Then one AskUserQuestion (wording in reference.md §3) whose text lists every To and Cc address, with options 「寄出」(Recommended only if the user already said to send) / 「修改內容」/ 「改收件者」/ 「取消」. Any change means a new `draft` (discard the old id) and asking again. Proceed only on 「寄出」.
5. **Send.** Run `send <id> -Confirm <token>` and tell the user in one line that a confirmation window has opened on their desktop and the mail leaves only when they click 寄出 there. The command blocks until they decide. Report the result per reference.md §4: sent (recipients, subject, time) or cancelled / timed out.
6. **After a cancel** do nothing more. Ask what to change only if the user speaks first.

## Hard rules

- Never run `send` before the user chose 「寄出」 in step 4, and never re-run it after a cancel or timeout without a fresh draft and a fresh yes.
- Never edit a draft file by hand, never pass a body that differs from what was shown, never add or drop a recipient the user did not confirm. If anything changes, go back to step 3.
- Never send to a guessed address. A name `draft` cannot resolve goes back to the user.
- Never send a mail about a message the phishing check marked 🚨 (POLICY.md) without saying so first.
- The footer (`Drafted by Claude, reviewed and approved by <name>`) stays; it is part of the approved text. The name comes from `send.approver` in settings, else the Outlook user's display name.

## Settings

`send.approver`, `send.footer` (`{approver}` placeholder), `send.quote_original` (default true), `send.dialog_timeout_seconds` (default 300). Change with `outlook-memory` or `settings.py set send.approver "Ben Chen"`.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-send/reference.md` for the writing rules, the draft card, the question wording and the result lines.

## Read-only rules

Everything else follows the read-only policy in `${CLAUDE_PLUGIN_ROOT}/POLICY.md`. This skill's only write is the Send inside `outlook_send.py send`, after the user's click. No Save, no Drafts folder, no flags, no marking read.
