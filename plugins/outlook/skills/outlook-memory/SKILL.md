---
name: outlook-memory
description: Manage the plugin's own state, READ-ONLY toward Outlook: personal memory notes (~/.outlook-skills/memory/<category>/<title>.md with YAML front matter; remember / forget / what do you know), settings (working hours, default store and search window, reranker gateway and consent, reply language), and the reply-habit profile (show it, explain it, tweak a line). Use when the user says "remember that Alice is ...", "forget ...", "what do you know about X", "set my working hours to ...", "what settings are you using", "what are my reply habits", "show my profile", or 記住 / 忘掉 / 你記得什麼 / 設定工作時間 / 目前的設定 / 我的回信習慣 / 看一下我的側寫. First-time setup and mailbox scans belong to outlook-setup.
---

# outlook-memory

Memory lives outside Outlook, in `~/.outlook-skills/memory/` (user level) and optionally `./.outlook-skills/memory/` (working directory, preferred for new notes when it exists). One Markdown file per topic, grouped by category, each with YAML front matter (`title`, `category`, `tags`, `created`, `updated`, `source`). Outlook is only ever read.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, Zoo Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line, and point the user to Claude Code, or to `outlook-open-msg` for .msg/.eml files they export from Outlook.

## Commands

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show                 # first_run flag + memory index
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" init [--local]       # create the folders
python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" list [--category X] [--json]
python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "alice" [--json]  # title, tags, body
python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" show "<title or path>"
python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" new --category people --title "Alice Chen" --tags legal,contoso --body "- 法務窗口，alice.chen@contoso.com" [--source bootstrap] [--local]
python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" append "Alice Chen" --body "- 合約由她發起" [--tags contracts]
python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" touch "<title>"        # after editing a body by hand
python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" remove "<title>"
```

`settings.py init` and `memory.py init` are the same command. `settings.py profile show` prints the reply-habit profile that `outlook-setup` builds; other skills read it when judging what needs a reply or drafting in the user's voice.

## 1. First run

Onboarding (settings wizard, mailbox scan with approved notes, reply-habit profile) lives in `outlook-setup`. When `settings.py show` returns `first_run: true`, hand over there once per conversation. "重新掃描信箱 / rebuild memory" also goes to `outlook-setup` (stage 3).

## 2. Remember / forget during normal use

- **"記住 X"** or a fact worth keeping surfaces while another skill runs: first `memory.py find "<key words>"`. If a file matches, `append` to it (and add tags if useful). Otherwise `new` under the fitting category with a clear title (reference.md §3). Show the exact line written. Never write silently: when the user did not say "remember", ask first with the host's question tool (「要記住這件事嗎？」with the proposed line).
- **"忘掉 X"**: `find`, show the match, `remove` (whole file) or edit the body to drop one bullet and `touch`. Confirm what was removed.
- **"你記得什麼？"**: `memory.py list`, present grouped by category (reference.md §4).

## 3. What belongs in memory

Yes: who someone is (role, address, what they usually write about), what a folder is for, a project's keywords and standing decisions, recurring events and their cadence, the user's preferences (language, working hours, formatting).
No: mail bodies or quotes longer than one line, attachments, credentials, anything the user declined to keep.

## Settings

Same folders, one file: `settings.json` holds only the keys the user changed; `settings.example.json` lists every key with its default. The working-directory file overrides the user file key by key.

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show                        # merged settings, source of each key, first_run, memory index
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" set <key> <value> [--local]  # dotted keys: working_hours.end 17:30, store 20230731, rerank.auto_consent true
```

- **"What settings are you using?"**: run `show`, present the table in reference.md §5, and say which file each non-default value comes from.
- **Change a setting**: confirm key and value in one line, run `set`, show the result. `--local` when the user says "for this project" or a `./.outlook-skills/` folder already exists.
- **Default store**: when outlook-status shows the mail lives in a .pst, offer `set store "<name>"` so every skill uses it.
- **Reranker consent**: `rerank.auto_consent true` lets outlook-search skip the per-run confirmation. Only on an explicit ask, and repeat once what will be sent and where.
- Never store credentials in memory notes; the reranker api key belongs in settings.json or an env var.

## Profile (reply habits and writing style)

`profile.md` is one file next to `settings.json`, built by `outlook-setup` stage 4 from `outlook_style.py`. It is not a memory note: it is the user's overall reply behaviour and voice, and other skills read it when they judge what needs a reply or draft one.

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" profile show
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" profile write --file <edited.md>
```

- **"我的回信習慣是什麼 / show my profile"**: run `profile show`, present per reference.md §6 (a short reading, not the raw file). No profile yet: say so and offer `outlook-setup` stage 4 (「更新回信習慣」).
- **"Why is X under never replied?"**: explain the rule (a reply counts only when a later mail from the user exists in the same conversation; phone or chat replies are invisible) and offer to note the exception as a line in the profile.
- **Tweak a line** ("我其實都會回 Cassie", "我的結尾改成 Best regards"): edit the corresponding bullet, keep the front matter, bump `updated`, write with `profile write`, show the diff. Only on the user's ask.
- **Rebuild from data**: hand over to `outlook-setup` (stage 4), which shows what changed before saving.

## Output format

`${CLAUDE_PLUGIN_ROOT}/skills/outlook-memory/reference.md` holds the front-matter schema, per-category guidance, title rules, the list layout, the settings key table, and how to present the profile. Read it once per conversation, in the same step as the first command.

## Read-only rules

Read-only, per `${CLAUDE_PLUGIN_ROOT}/POLICY.md` (no need to open it): Outlook is only read; the only writes are inside `.outlook-skills/` folders, and only after the user agreed.
