---
name: outlook-setup
description: Guided setup of the Outlook skills, READ-ONLY toward Outlook - runs automatically the first time (no ~/.outlook-skills yet) and again on request. Four stages, each with clickable questions - confirm the scan and its cost, settings one item at a time (default store, working hours, search window, language, reranker), a read-only mailbox scan that proposes up to 10 memory notes to approve one by one, and a reply-habit and writing-style profile built from Sent Items. Use when the user says "set up the Outlook skills", "run setup again", "re-scan my mailbox", "rebuild my memory", "update my profile", or 初始化 / 重新設定 / 重新掃描信箱 / 重建記憶 / 更新回信習慣.
---

# outlook-setup

The onboarding wizard. Outlook is only read; the only writes are `settings.json`, `memory/<category>/<title>.md` and `profile.md` under `~/.outlook-skills/`, each only after the user said yes in a question dialog.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, or Claude Code inside Claude Desktop) and offers a structured question tool (AskUserQuestion). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so and stop.

## When it runs

- Automatically: any skill sees `first_run: true` from `settings.py show` and hands over here, once per conversation. A "not now" answer ends setup; the other skill continues.
- On request: "重新設定 / run setup again" (all stages, current values shown as defaults), "重新掃描信箱 / rebuild memory" (stage 3 only), "更新回信習慣 / update my profile" (stage 4 only), "改設定" (stage 2 only, or just use `outlook-memory`).

## Commands

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show | init | set <key> <value>
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_status.py"                                  # stores, sizes, folder counts (stage 2 options)
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_overview.py" -Days 180 [-Store X] -OutFile "<tmp>/overview.json"
python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find | new | append
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_style.py" -Days 180 [-Store X] -OutFile "<tmp>/style.json"
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" profile write --file "<tmp>/profile.md"
python "${CLAUDE_PLUGIN_ROOT}/scripts/rerank.py" --show-config
```

## Stage 1: confirm

Run `outlook_status.py` first (fast, read-only) to learn the stores and their sizes. Then ask **one** question with AskUserQuestion, wording per reference.md §1: what the three remaining stages do, that the scan reads the last 180 days of subjects, senders, folders and meetings but no bodies, that nothing is written without a yes, and the expected duration (about a minute per 3,000 mails; say "a few minutes" for a large PST). Options: 「開始設定」(Recommended) / 「只設定，不掃描信箱」/ 「這次先不要」. Not now: create nothing, return.

## Stage 2: settings, one item at a time

`settings.py init` if the folder does not exist. Then ask the items in reference.md §2, **one question per item**, each as an AskUserQuestion with clickable options (batch up to 4 items per dialog when the host allows several questions in one call). Build the options from real data: stores from `outlook_status.py` (mark the one holding most mail as Recommended), the reranker item only if `rerank.py --show-config` reports `usable: true`. On a re-run, the current value is the first option and marked as current. Write each answer with `settings.py set`; skip items the user leaves at default. Finish with the settings table from reference.md §2.

## Stage 3: memory, up to 10 notes, approved one by one

Run `outlook_overview.py` (with `-Store` from stage 2). If the JSON is large, delegate the drafting to a subagent (see "Delegating heavy reads" below). Draft **at most 10** candidate notes, the most useful first: people the user exchanges the most mail with, folders with a clear purpose, the two or three busiest topics, every recurring meeting. Skip newsletters and anyone already in memory (`memory.py find`). Then ask per candidate with AskUserQuestion, up to 4 per dialog, each question showing title, category, tags and the one-line content, options 「加入」/「跳過」/「改一下再加」. For "改一下", ask what to change in the next dialog, then write. Write approved notes with `memory.py new --source bootstrap`; report the count per category and the folder path. On a re-run propose only new or changed items and `append` to existing notes.

## Stage 4: reply habits and writing style

Run `outlook_style.py` (same `-Store`). Turn the JSON into the profile in reference.md §4: who always gets a reply and who never does, how fast, typical length, language, greetings and closings, signature, sending hours. Show it and ask one AskUserQuestion: 「存成 profile.md」(Recommended) / 「修改後再存」/ 「不要存」. Save with `settings.py profile write`. The profile is not memory: it is one file other skills read when they judge what needs a reply or draft one in the user's voice. On a re-run show what changed before saving.

## Delegating heavy reads

When the host offers subagents and the overview JSON is large (more than about 30 senders or topics, or the file exceeds ~50 KB), hand stage 3 drafting to a subagent: give it the path of `overview.json`, reference.md §3 (categories, what each takes, title rules), the current memory index so it does not propose duplicates, and the user's language; ask it to return the 10-row candidate table and nothing else. The approval questions and every write stay in this conversation.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-setup/reference.md` for the exact wording of each question, the settings items and their options, the candidate table, and the profile.md template.

## Read-only rules

Follow the read-only policy in `${CLAUDE_PLUGIN_ROOT}/POLICY.md`. Outlook is only read. Nothing under `~/.outlook-skills/` is written without a yes in a question dialog. Remind the user once to add `.outlook-skills/` to `.gitignore` when the working directory is a git repo.
