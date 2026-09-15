---
name: outlook-followup
description: Find mails that are waiting for a reply, READ-ONLY, from the local Outlook (Classic) mailbox. Two directions - mails the user sent that nobody answered for N days, and mails other people sent the user that ask something and are still unanswered. Use when the user asks "who hasn't replied to me", "what am I still waiting on", "which emails do I still owe a reply", "what did people ask me that I haven't answered", or 誰還沒回我 / 我還在等誰 / 我還欠誰回信 / 有哪些問題我還沒回.
---

# outlook-followup

Read-only. Compares each mail with the later mails in the same conversation to decide whether it was answered.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line and point the user to Claude Code.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_followup.py" -Direction sent      # I wrote, they did not answer (default: older than 3 days)
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_followup.py" -Direction received  # they asked, I did not answer (default: older than 2 days)
```

Options: `-Days N` minimum age before a mail counts as waiting; `-Lookback N` how far back to scan (default 60); `-Store` / `-AllStores`; `-QuestionsOnly` (received) keeps only mails that look like a question or request; `-Max`, `-PreviewLength`, `-OutFile`.

## Workflow

1. Pick the direction from the wording: "who hasn't replied" = sent; "what do I owe" / "what did they ask" = received. If unclear, run both and present two short lists.
2. Use the `store` setting or `-AllStores` when the mail lives in a PST.
3. Present per reference.md, oldest waiting first. For **received**, use `LooksLikeQuestion` and `DirectToMe` plus the preview to separate real asks from FYI mails; say when you are unsure and let the user decide.
4. Offer next steps that stay read-only: open the thread with `outlook-thread`, or draft a nudge in chat. Never send.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant). `first_run: true` means `~/.outlook-skills` does not exist yet: hand over to `outlook-memory`'s onboarding first (respect a "not now"). Apply `store` (pass it as `-Store`), `language`, and any search defaults. When the request names a person, folder or project, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "<word>"` and `show` the note so aliases resolve. Never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-followup/reference.md` for the JSON fields and the presentation template.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never reply, forward, flag, mark read, move or delete. Drafts, if asked for, are written in chat only.

