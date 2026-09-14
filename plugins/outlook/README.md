# outlook (plugin)

Read-only skills for a **local Windows Classic Outlook** mailbox.

| Skill | What it does |
|---|---|
| `outlook-status` | Outlook version, profiles, accounts, .pst/.ost files, folder counts |
| `outlook-search` | Search mail by sender, subject, body, date, attachments, unread |
| `outlook-thread` | Read a whole conversation with full bodies, ready to summarise |
| `outlook-calendar` | Agenda for a date range with recurrences expanded and conflicts flagged |
| `outlook-open-msg` | Parse a .msg / .eml file without Outlook |

## Requirements

- Windows with **Classic Outlook** (2016 / 2019 / 2021 / Microsoft 365). "New Outlook" has no COM object model and is not supported; `outlook-status -SkipCom` still works there.
- Windows PowerShell 5.1 (built in) or PowerShell 7 for the four COM-based skills.
- Python 3.8+ for `outlook-open-msg`; `pip install extract-msg` for .msg files.
- Outlook may be open or closed. If closed, the COM call starts it in the background under the current user's profile.

## Read-only policy

This plugin **never writes to Outlook**. Concretely, no script or skill may:

- call `Save`, `Send`, `Delete`, `Move`, `Copy`, `Forward`, `Reply`, `ReplyAll`, `Respond`, `Display`;
- set any property (`UnRead`, `Categories`, `FlagStatus`, `Importance`, `BusyStatus`, ...);
- create items, folders, rules, or appointments;
- compact, repair, detach or attach data files;
- write anywhere except a user-specified `-OutFile` / `--extract-to` path.

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

What it cannot get around: AppLocker / WDAC **Constrained Language Mode** blocks `New-Object -ComObject`, so the four COM-based skills will not work there regardless of execution policy. `outlook-open-msg` (Python, no COM) still works, and `outlook-status -SkipCom` still reports registry and file-system information.

## Fuzzy search with a reranker (optional)

`outlook-search` falls back to a cross-encoder reranker for fuzzy / semantic queries. `scripts/rerank.py` sends `(query, message preview)` pairs to an OpenAI-compatible gateway (vLLM `/v1/rerank`, or `/v1/score` with `--endpoint score`) in batches of 30 and sorts the candidates by relevance. Default model: `bge-reranker-v2-m3`.

Configure the gateway in `~/.claude/settings.json`:

```json
{
  "env": {
    "OUTLOOK_RERANK_URL": "http://your-vllm-host:8000/v1",
    "OUTLOOK_RERANK_MODEL": "bge-reranker-v2-m3",
    "OUTLOOK_RERANK_API_KEY": "optional-bearer-token"
  }
}
```

The same names work as process environment variables or as `--gateway` / `--model` / `--api-key` flags. A specific name always wins over a generic fallback (`OPENAI_BASE_URL`, `ANTHROPIC_BASE_URL`, `OPENAI_API_KEY`, `ANTHROPIC_AUTH_TOKEN`) no matter where each is set. `python rerank.py --show-config` prints what would be used without sending anything; `--dry-run` prints the documents that would be sent.

**Privacy**: subject, sender, date and a preview (default 600 characters of combined text) of every candidate leave the machine. The skill therefore asks the user for confirmation, naming the gateway and the number of mails, before every reranker call. Point it at an internal vLLM instance, not a public API.

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
  skills/
    outlook-status/    SKILL.md + reference.md
    outlook-search/    SKILL.md + reference.md
    outlook-thread/    SKILL.md + reference.md
    outlook-calendar/  SKILL.md + reference.md
    outlook-open-msg/  SKILL.md + reference.md
```

Skills reference scripts via `${CLAUDE_PLUGIN_ROOT}`, which Claude Code resolves to the installed plugin directory.
