---
name: outlook-settings
description: Show, initialise or change the Outlook skills' own settings and memory files (.outlook-skills/ in the working directory or ~/.outlook-skills/): working hours, default search window, default store, reranker gateway, and the memory.md notes Claude keeps (contact aliases, folder meanings, project keywords, preferences). Use when the user says "remember that Alice is ...", "set my working hours to ...", "what settings are you using", "forget ...", or 記住 / 設定工作時間 / 目前的設定 / 初始化設定 / 忘掉.
---

# outlook-settings

Manages the plugin's own configuration. It never touches Outlook; the only files it writes are inside `.outlook-skills/` folders.

## Layout

| Folder | Scope | Wins |
|---|---|---|
| `~/.outlook-skills/` | the user, every project | lower |
| `./.outlook-skills/` (cwd or any parent up to home) | this working directory | higher, key by key |

Each folder holds `settings.json` (only the keys the user changed), `settings.example.json` (every key with its default, reference only) and `memory.md`.

## Commands

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show            # merged settings, which file set each key, memory paths
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" init [--local]  # create the folder and template files (never overwrites)
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" set <key> <value> [--local]
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" memory          # print memory.md contents
```

`set` uses dotted keys (`working_hours.end 17:30`, `search.default_lookback_days 180`, `rerank.auto_consent true`). Values `true/false/null`, numbers and JSON arrays are parsed; anything else is a string. Without `--local` it writes the user-level file.

## Workflow

- **"What settings are you using?"**: run `show`, present the table in reference.md, and say which file each non-default value comes from.
- **Change a setting**: confirm the key and value in one line, run `set`, show the result. Use `--local` when the user says "for this project" or a `./.outlook-skills/` folder already exists here.
- **First use / no folder yet**: run `init` (user level) and tell the user where the files are. Suggest adding `.outlook-skills/` to `.gitignore` when the working directory is a git repo, because memory.md will contain names and addresses.
- **Remember something** ("記住 Alice 是 alice.chen@contoso.com", "供應商的信都在 Inbox/Vendors"): append one bullet under the matching heading of memory.md (prefer the local file when it exists, else the user file; create with `init` if neither exists). Show the exact line you added.
- **Forget something**: remove the bullet, show what was removed.
- **Reranker consent**: `rerank.auto_consent true` lets outlook-search skip the per-run confirmation. Only set it when the user explicitly asks, and repeat once what will be sent and where.

## What may go into memory.md

Yes: contact aliases and addresses, what a folder is for, project names and their keywords, standing decisions ("Q3 budget: plan B"), formatting or language preferences, the user's own working hours or time zone.
No: mail bodies, attachments, anything the user did not ask to keep, credentials (the reranker api key belongs in settings.json or an env var, not memory).

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-settings/reference.md` for the settings key reference, the memory.md structure, and how to present `show`.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Writes are limited to `.outlook-skills/` folders.
