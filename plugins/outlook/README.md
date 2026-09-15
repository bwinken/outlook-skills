# outlook (plugin)

Read-only skills for a **local Windows Classic Outlook** mailbox.

| Skill | What it does |
|---|---|
| `outlook-status` | Outlook version, profiles, accounts, .pst/.ost files, folder counts |
| `outlook-search` | Search mail by sender, subject, body, date, attachments, unread |
| `outlook-thread` | Read a whole conversation with full bodies, ready to summarise |
| `outlook-agenda` | What is on the calendar for a date range, recurrences expanded, conflicts and unanswered invites flagged |
| `outlook-availability` | When the user is free: open slots within working hours, or a slot of a required length |
| `outlook-open-msg` | Parse a .msg / .eml file without Outlook (standard library only) |
| `outlook-settings` | The plugin's own settings: working hours, defaults, reranker gateway and consent, reply language |
| `outlook-memory` | Personal memory: first-run onboarding that scans the mailbox and proposes notes; remember / forget / what do you know |

## Hosts

- **Claude Code**: installed as a plugin from the marketplace at the repo root; `${CLAUDE_PLUGIN_ROOT}` in the skill files is expanded by Claude Code.
- **Zoo Code** and other Agent Skills hosts: `python install.py` copies `skills/<name>/*.md` into `~/.roo/skills/` (`--project` for `./.roo/skills/`, `--agents` for `.agents/skills/`) and rewrites `${CLAUDE_PLUGIN_ROOT}` to this folder's absolute path, so the scripts here run from the copies. `--uninstall` removes only what it created.

## Requirements

- Windows with **Classic Outlook** (2016 / 2019 / 2021 / Microsoft 365). "New Outlook" has no COM object model and is not supported; `outlook_status.py -SkipCom` still works there.
- Python 3.8+ and one package: `pip install pywin32` (COM access). Everything else is the standard library.
- Outlook may be open or closed. If closed, the COM call starts it in the background under the current user's profile.

## Read-only policy

This plugin **never writes to Outlook**. Concretely, no script or skill may:

- call `Save`, `Send`, `Delete`, `Move`, `Copy`, `Forward`, `Reply`, `ReplyAll`, `Respond`, `Display`;
- set any property (`UnRead`, `Categories`, `FlagStatus`, `Importance`, `BusyStatus`, ...);
- create items, folders, rules, or appointments;
- compact, repair, detach or attach data files;
- write anywhere except a user-specified `-OutFile` / `--extract-to` path and the plugin's own `.outlook-skills/` folders.

Reading through COM does not change read/unread state. The shared module `scripts/outlook_com.py` exposes only getters; add new skills on top of it and keep the same rule.

## Running the scripts

Every script is plain Python: `python scripts/outlook_search.py -From alice -After 2026-09-01`. Options accept PowerShell-style (`-From`) or long (`--from`) spellings, output is UTF-8 JSON on stdout or `-OutFile`. Nothing runs through PowerShell, so execution policy, code pages and BOMs are not a concern. AppLocker / WDAC policies that block COM automation still block these scripts; `outlook-open-msg` and `outlook_status.py -SkipCom` work regardless.

## Fuzzy search with a reranker (optional)

`outlook-search` narrows first (sender, dates, folder, keyword expansion). Only when the narrowed set is still larger than about 20 mails does it offer a cross-encoder reranker: `scripts/rerank.py` sends `(query, message preview)` pairs to an OpenAI-compatible gateway (vLLM `/v1/rerank`, or `/v1/score`) in batches of 30 and sorts the candidates by relevance. Default model: `bge-reranker-v2-m3`.

**Gateway resolution.** By default the script reuses the `ANTHROPIC_BASE_URL` that Claude Code already has in `~/.claude/settings.json` (`env` block), on the assumption that the same gateway also serves the reranker model. Override with `OUTLOOK_RERANK_URL` when the reranker lives elsewhere. Full order, a more specific name winning wherever it is set (flag, process env, or settings.json):

| Setting | Names, first wins |
|---|---|
| gateway | `--gateway`, `OUTLOOK_RERANK_URL`, `ANTHROPIC_BASE_URL`, `OPENAI_BASE_URL` |
| model | `--model`, `OUTLOOK_RERANK_MODEL`, default `bge-reranker-v2-m3` |
| api key | `--api-key`, `OUTLOOK_RERANK_API_KEY`, `ANTHROPIC_AUTH_TOKEN`, `ANTHROPIC_API_KEY`, `OPENAI_API_KEY` |

Example `~/.claude/settings.json`:

```json
{
  "env": {
    "ANTHROPIC_BASE_URL": "http://your-gateway:8000",
    "ANTHROPIC_AUTH_TOKEN": "token",
    "OUTLOOK_RERANK_MODEL": "bge-reranker-v2-m3"
  }
}
```

**Endpoint probe.** With `--endpoint auto` (default) the script sends a one-word test request to `{base}/v1/rerank`, then `{base}/v1/score`, and uses the first that returns a score. A base URL that is not a reranker gateway (for example the public Anthropic API) fails the probe and is reported as unusable; no mail data is sent. `python rerank.py --show-config` runs the resolution and probe and prints the result (exit 0 usable, 1 not); `--dry-run` prints the documents that would be sent.

**Privacy.** Subject, sender, date and a preview (default 600 characters of combined text) of every candidate leave the machine. The skill therefore asks the user for confirmation, naming the gateway and the number of mails, before every reranker call. Point it at an internal vLLM instance, not a public API.

## Settings

The skills keep their own configuration outside Outlook, in two optional folders; the working-directory layer overrides the user layer key by key:

| Folder | Scope |
|---|---|
| `~/.outlook-skills/` | the user, every project |
| `./.outlook-skills/` (working directory or any parent up to home) | this project |

`settings.json` holds only the keys you changed; `settings.example.json` lists them all with defaults (working hours, minimum free slot, default search window and folder, default store, reranker gateway and consent, reply language). `scripts/settings.py show` merges the layers and reports which file set each key; every skill runs it once per conversation. Change values with the `outlook-settings` skill or `settings.py set key value [--local]`.

## Memory

Memory is what lets the skills understand the user's own vocabulary: that "Alice" is alice.chen@contoso.com, that "the vendor mails" live in `Inbox/Vendors`, that "the budget thread" means the Q3 budget project. It lives next to the settings, never inside Outlook.

**Files.** `memory/<category>/<title>.md`, one Markdown file per topic, in the same two layers as settings. Categories: `people`, `folders`, `projects`, `recurring`, `preferences`. Each file starts with YAML front matter that `scripts/memory.py` maintains:

```
---
title: Alice Chen
category: people
tags: [legal, contoso]
created: 2026-09-14T10:02:11
updated: 2026-09-14T13:40:05
source: bootstrap
---
- 法務窗口，alice.chen@contoso.com
```

The title is the name someone would use to find the note again (a person's name, a folder path, a project's short name, a meeting subject). Related facts go into the existing file (`memory.py append`), which bumps `updated`; `memory.py find` searches title, tags and body so duplicates are avoided.

**First run.** `settings.py show` returns `first_run: true` while `~/.outlook-skills` does not exist. Whichever skill sees it hands over to `outlook-memory`, which asks the user (AskUserQuestion in Claude Code, ask_followup_question in Zoo Code) whether to build a personal memory: build and scan the mailbox, create an empty one, or not now. With a scan, `outlook_overview.py` reads the last 180 days and returns counts only (top senders and recipients, folders, frequent conversation topics, newsletters, recurring meetings; no bodies). Claude drafts notes from that, shows them as a table, and writes only the ones the user approves (`source: bootstrap`).

**Everyday use.** "記住 …" appends or creates a note; "忘掉 …" removes one and says what was removed; "你記得什麼" lists titles by category; "重新掃描信箱" re-runs the overview and proposes only new or changed notes. Other skills read just the index (title, category, tags, updated) and open a note only when a request names a person, folder, project or routine.

**Boundaries.** Nothing is written without the user's say-so. Mail bodies, attachments and credentials never go into memory. The notes contain names and addresses, so add `.outlook-skills/` to `.gitignore` in a repo.

## Phishing warnings

Whenever a skill shows the content of a single mail (`outlook-open-msg`, `outlook-search` with `-IncludeBody`, `outlook-thread`), Claude first runs a quick phishing check and, if it trips, warns **before anything else** in the reply. The warning must be impossible to miss: it comes first, uses the 🚨 banner below, names the concrete signals, and tells the user what not to do. Soft wording ("might be worth checking") is not acceptable for a high-risk verdict.

Levels:

| Level | When | Reply |
|---|---|---|
| 🚨 高風險 | two or more strong signals, or one strong signal plus a request for credentials, payment or urgent action | banner first, then the message card with all links defanged; do not summarise the mail's call to action as if it were legitimate |
| ⚠️ 可疑 | one strong signal, or several weak ones | one bold warning line above the card, links defanged |
| none | nothing tripped | normal card; say "看起來正常" only when asked, never "安全" |

Strong signals: display name and address disagree (「Microsoft Support」 from a non-Microsoft domain); Reply-To differs from From; SPF, DKIM or DMARC `fail` in `Authentication-Results`; a look-alike domain (`micros0ft`, `contoso-secure.ru`); an attachment of a risky type (.html, .htm, .iso, .img, .lnk, .js, .vbs, .zip with such content, macro-enabled Office files); a link whose visible text and target differ; a login or password-reset link to a domain that is not the claimed sender's.
Weak signals: urgency or threats (帳號將被停用、24 小時內), generic greeting, unexpected invoice or delivery notice, external sender writing about internal matters, first-time sender asking for a bank change.

Banner (Traditional Chinese by default; translate for an English user):

```
🚨 **警告：這封信很可能是釣魚信，請勿點擊連結、開啟附件或回覆。** 🚨

判斷依據：
- 顯示名稱「Microsoft Support」，但寄件地址是 support@mail-secure-login.ru
- SPF 與 DKIM 皆為 fail
- 附件 invoice.html 是釣魚常用的檔案類型

建議：在 Outlook 用「回報 > 釣魚」或轉交 IT；如果已經點過連結或輸入過密碼，立刻更改密碼並通知 IT。
```

Rules for both warning levels:
- Defang every URL from the suspicious mail: `hxxps://login[.]contoso-secure[.]ru/reset`. Never leave a clickable link.
- Do not extract or open attachments from a 🚨 mail even if the user asks; explain why once and offer the header analysis instead.
- In lists (search results, thread timelines) mark such messages with 🚨 in the subject cell and say so above the table.
- The verdict is a judgement, not a scan result: say what was checked and what was not (for example, no header data in COM output, so SPF/DKIM unknown).

## Output format

Each skill ships a `reference.md` next to its `SKILL.md` documenting the script's JSON fields and the presentation template Claude should use in its reply (tables, sections, date formats, truncation rules). Change the template there, not in SKILL.md.

## Layout

```
plugins/outlook/
  .claude-plugin/plugin.json
  install.py                installer for Zoo Code (.roo/skills) / .agents/skills hosts
  scripts/
    outlook_com.py            shared read-only COM helpers (pywin32), folder resolution, JSON output
    outlook_status.py
    outlook_search.py
    outlook_thread.py
    outlook_calendar.py       used by outlook-agenda and outlook-availability
    outlook_overview.py       read-only mailbox overview for memory onboarding
    read_msg.py               .msg / .eml parser CLI
    msgfile.py                standard-library OLE2 + MAPI reader for .msg
    rerank.py                 optional reranker client for fuzzy search (asks consent first)
    settings.py               merges ~/.outlook-skills and ./.outlook-skills settings; init / set / show
    memory.py                 memory notes: list / find / new / append / touch / remove
  tests/                      fake Outlook object model + script tests (no Windows needed)
  skills/
    outlook-status/    SKILL.md + reference.md
    outlook-search/    SKILL.md + reference.md
    outlook-thread/    SKILL.md + reference.md
    outlook-agenda/    SKILL.md + reference.md
    outlook-availability/ SKILL.md + reference.md   (same script as agenda)
    outlook-open-msg/  SKILL.md + reference.md
    outlook-settings/  SKILL.md + reference.md
    outlook-memory/    SKILL.md + reference.md
```

Skills reference scripts via `${CLAUDE_PLUGIN_ROOT}`, which Claude Code resolves to the installed plugin directory.
