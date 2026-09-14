# outlook (plugin)

Read-only skills for a **local Windows Classic Outlook** mailbox.

| Skill | What it does |
|---|---|
| `outlook-status` | Outlook version, profiles, accounts, .pst/.ost files, folder counts |
| `outlook-search` | Search mail by sender, subject, body, date, attachments, unread |
| `outlook-thread` | Read a whole conversation with full bodies, ready to summarise |
| `outlook-agenda` | What is on the calendar for a date range, recurrences expanded, conflicts and unanswered invites flagged |
| `outlook-availability` | When the user is free: open slots within working hours, or a slot of a required length |
| `outlook-open-msg` | Parse a .msg / .eml file without Outlook |
| `outlook-settings` | The plugin's own settings: working hours, defaults, reranker gateway and consent, reply language |
| `outlook-memory` | Personal memory: first-run onboarding that scans the mailbox and proposes notes; remember / forget / what do you know |

## Requirements

- Windows with **Classic Outlook** (2016 / 2019 / 2021 / Microsoft 365). "New Outlook" has no COM object model and is not supported; `outlook-status -SkipCom` still works there.
- Windows PowerShell 5.1 (built in) or PowerShell 7 for the COM-based skills (status, search, thread, agenda, availability, memory onboarding).
- Python 3.8+ for `outlook-open-msg`; `pip install extract-msg` for .msg files.
- Outlook may be open or closed. If closed, the COM call starts it in the background under the current user's profile.

## Read-only policy

This plugin **never writes to Outlook**. Concretely, no script or skill may:

- call `Save`, `Send`, `Delete`, `Move`, `Copy`, `Forward`, `Reply`, `ReplyAll`, `Respond`, `Display`;
- set any property (`UnRead`, `Categories`, `FlagStatus`, `Importance`, `BusyStatus`, ...);
- create items, folders, rules, or appointments;
- compact, repair, detach or attach data files;
- write anywhere except a user-specified `-OutFile` / `--extract-to` path and the plugin's own `.outlook-skills/` folders.

Reading through COM does not change read/unread state. The shared library `scripts/OutlookReadOnly.ps1` exposes only getters; add new skills on top of it and keep the same rule.

## Execution policy

Many Windows machines refuse to run `.ps1` files ("running scripts is disabled on this system"). The skills handle this in two layers:

1. **Default policy (Restricted / RemoteSigned)**: every skill runs scripts with `powershell -ExecutionPolicy Bypass -File ...`. That flag applies to the single process only and needs no admin rights or machine changes.
2. **Policy enforced by Group Policy**: the flag is ignored. The skills then fall back to loading the script text as a script block:
   ```
   powershell -NoProfile -Command "$env:OUTLOOK_SKILLS_SCRIPTS='<scripts dir>'; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '<scripts dir>\Search-OutlookMail.ps1'))) -From alice"
   ```
   Execution policy only governs script *files*; text passed to `-Command` or turned into a script block is not checked. For the same reason the shared helper is a plain `.ps1` that scripts dot-source through a script block, not a `.psm1` imported with `Import-Module`. When run this way `$PSScriptRoot` is empty, so the scripts locate the helper through `OUTLOOK_SKILLS_SCRIPTS`.

Neither layer changes any machine or user setting. The plugin never runs `Set-ExecutionPolicy`.

What it cannot get around: AppLocker / WDAC **Constrained Language Mode** blocks `New-Object -ComObject`, so the COM-based skills will not work there regardless of execution policy. `outlook-open-msg` (Python, no COM) still works, and `outlook-status -SkipCom` still reports registry and file-system information.

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

## Settings and memory

The skills keep their own configuration outside Outlook, in two optional folders:

| Folder | Scope |
|---|---|
| `~/.outlook-skills/` | the user, every project |
| `./.outlook-skills/` (working directory or any parent up to home) | this project; overrides the user level key by key |

**Settings**: `settings.json` holds only the keys you changed (`settings.example.json` lists them all). `scripts/settings.py show` merges the layers, reports which file set each key, and every skill runs it once per conversation (working hours, default search window, default store, reranker gateway, reply language). Change with the `outlook-settings` skill or `settings.py set key value`.

**Memory**: `memory/<category>/<title>.md`, one Markdown file per person, folder, project, routine or preference, with YAML front matter (`title`, `category`, `tags`, `created`, `updated`, `source`). Categories: `people`, `folders`, `projects`, `preferences`, `recurring`. `scripts/memory.py` creates, finds, appends to and removes notes and keeps the timestamps; skills only read the index and open the notes a request needs. The first time any skill runs and `~/.outlook-skills` is missing, `outlook-memory` asks whether to build a personal memory: with a yes it runs `Get-OutlookOverview.ps1` (read-only, counts only, no bodies) over the last 180 days, drafts notes about frequent contacts, folders, topics and recurring meetings, shows the draft, and writes only what the user approves. Later, "記住 …" appends to an existing note when one matches or creates a new one. Nothing is written without the user's say-so, and mail bodies never go into memory. Add `.outlook-skills/` to `.gitignore` in a repo; the notes contain names and addresses.

## Output format

Each skill ships a `reference.md` next to its `SKILL.md` documenting the script's JSON fields and the presentation template Claude should use in its reply (tables, sections, date formats, truncation rules). Change the template there, not in SKILL.md.

## Layout

```
plugins/outlook/
  .claude-plugin/plugin.json
  scripts/
    OutlookReadOnly.ps1       shared read-only COM helpers (dot-sourced, not a module)
    Get-OutlookStatus.ps1
    Search-OutlookMail.ps1
    Get-OutlookThread.ps1
    Get-OutlookCalendar.ps1
    read_msg.py
    rerank.py                 optional reranker client for fuzzy search (asks consent first)
    settings.py               merges ~/.outlook-skills and ./.outlook-skills settings; init / set / show
    memory.py                 memory notes: list / find / new / append / touch / remove
    Get-OutlookOverview.ps1   read-only mailbox overview (top senders, folders, topics, recurring meetings) for onboarding
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
