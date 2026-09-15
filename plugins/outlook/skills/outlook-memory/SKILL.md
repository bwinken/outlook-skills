---
name: outlook-memory
description: Manage the plugin's own state, READ-ONLY toward Outlook: personal memory (~/.outlook-skills/memory/<category>/<title>.md with YAML front matter; first-run onboarding that scans the mailbox and proposes notes; remember / forget / what do you know) and settings (working hours, default store and search window, reranker gateway and consent, reply language). Use when the user says "remember that Alice is ...", "forget ...", "what do you know about X", "set up my memory", "set my working hours to ...", "what settings are you using", or 記住 / 忘掉 / 你記得什麼 / 建立個人化記憶 / 設定工作時間 / 目前的設定.
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
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_overview.py" -Days 180 -OutFile "<tmp>/overview.json"
```

All scripts are Python (pywin32 for Outlook). `settings.py init` and `memory.py init` are the same command.

## 1. First run (onboarding)

`settings.py show` returns `first_run: true` when `~/.outlook-skills` does not exist. Any skill that sees this hands over here **once per conversation**:

1. Ask with the host's structured question tool (Claude Code: **AskUserQuestion**; Zoo Code / Roo Code: **ask_followup_question**), never as plain text:
   - question: 這是第一次使用 Outlook skills。要建立個人化記憶嗎？我會唯讀掃描最近幾個月的信箱，整理出常聯絡的人、資料夾用途、常見專案主題和固定會議，存成本機的 Markdown 檔（`~/.outlook-skills/memory/`），之後搜尋和摘要時就能認得「Alice」「供應商的信」這類說法。
   - options: 「建立並掃描信箱」(Recommended) / 「只建立空的記憶，不掃描」/ 「這次先不要」.
   - A "not now" answer means: continue the original task, do not ask again this conversation, and do not create anything.
2. On "empty only": run `settings.py init`, say where the folders are, continue the original task.
3. On "build and scan":
   - `settings.py init`, then run `outlook_overview.py -Days 180 -OutFile "<tmp>/overview.json"` (add `-Store` from settings, or the .pst store outlook-status showed as holding the mail). It returns counts only, no bodies.
   - Draft notes from the JSON, following reference.md §2 for what each category takes and how to name titles (delegate this step when the JSON is large; see "Delegating heavy reads"). Aim for quality over quantity: roughly 10 to 25 people, every folder with a clear purpose, 5 to 15 projects/topics, all recurring meetings, and skip newsletters unless the user wants them.
   - Show the draft as a table (title, category, tags, one-line content) and ask with the same question tool: 「全部寫入」/「讓我挑」/「取消」. For "let me pick", ask a multi-select question listing the titles (Zoo Code has no multi-select: list the titles numbered and ask which numbers to keep).
   - Write the approved notes with `memory.py new --source bootstrap`, one call per note. Report how many were written per category and where.
4. Mention once that `.outlook-skills/` should be in `.gitignore` if the working directory is a git repo.

## Delegating heavy reads

When the host offers subagents (Claude Code's Agent tool, including the Code tab in Claude Desktop) and the overview JSON is large (more than about 30 senders or topics, or the file exceeds ~50 KB), hand the reading to a subagent so the raw data never enters this conversation. Give it: the path of `overview.json`, the contents of reference.md §2 and §3 (categories, what each takes, title rules), the current memory index from `settings.py show` so it does not propose duplicates, and the user's language. Ask it to return the draft table from reference.md §4 (title, category, tags, one-line content) as Markdown, with a one-line count per category and nothing else, no raw records. Anything that needs the user's consent (writing memory, sending candidates to the reranker, copying attachments out) stays in this conversation; a subagent never asks the user and never writes. Without subagents, do the same work here but read only what the step needs.

## 2. Remember / forget during normal use

- **"記住 X"** or a fact worth keeping surfaces while another skill runs: first `memory.py find "<key words>"`. If a file matches, `append` to it (and add tags if useful). Otherwise `new` under the fitting category with a clear title (reference.md §3). Show the exact line written. Never write silently: when the user did not say "remember", ask first with the host's question tool (「要記住這件事嗎？」with the proposed line).
- **"忘掉 X"**: `find`, show the match, `remove` (whole file) or edit the body to drop one bullet and `touch`. Confirm what was removed.
- **"你記得什麼？"**: `memory.py list`, present grouped by category (reference.md §4).
- **"重新掃描信箱"**: rerun the overview and propose only *new* or *changed* notes; existing files are appended to, never overwritten, unless the user says so.

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

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-memory/reference.md` for the front-matter schema, per-category guidance, title rules, the onboarding draft table, the list layout, and the settings key table.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Outlook is only read; the only writes are inside `.outlook-skills/` folders, and only after the user agreed.
