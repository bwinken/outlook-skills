---
name: outlook-settings
description: Show, initialise or change the Outlook skills' own settings (.outlook-skills/settings.json in the working directory or ~/.outlook-skills/): working hours, default search window and folder, default store, reranker gateway and consent, reply language. Use when the user says "set my working hours to ...", "what settings are you using", "use this gateway for reranking", "initialise the settings", or 設定工作時間 / 目前的設定 / 初始化設定. For remembering people, folders or projects use outlook-memory.
---

# outlook-settings

Manages the plugin's own configuration. It never touches Outlook; the only file it writes is `settings.json` inside a `.outlook-skills/` folder. Memory notes are handled by `outlook-memory`.

## Layout

| Folder | Scope | Wins |
|---|---|---|
| `~/.outlook-skills/` | the user, every project | lower |
| `./.outlook-skills/` (cwd or any parent up to home) | this working directory | higher, key by key |

Each folder holds `settings.json` (only the keys the user changed), `settings.example.json` (every key with its default, reference only) and a `memory/` tree managed by `outlook-memory`.

## Commands

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show            # merged settings, which file set each key, memory paths
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" init [--local]  # create the folder and template files (never overwrites)
python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" set <key> <value> [--local]
```

`set` uses dotted keys (`working_hours.end 17:30`, `search.default_lookback_days 180`, `rerank.auto_consent true`). Values `true/false/null`, numbers and JSON arrays are parsed; anything else is a string. Without `--local` it writes the user-level file.

## Workflow

- **"What settings are you using?"**: run `show`, present the table in reference.md, and say which file each non-default value comes from.
- **Change a setting**: confirm the key and value in one line, run `set`, show the result. Use `--local` when the user says "for this project" or a `./.outlook-skills/` folder already exists here.
- **First use / no folder yet** (`first_run: true`): this is `outlook-memory`'s onboarding; let it ask the user. If the user only wants settings, run `init` (user level) and tell them where the files are.
- **Remember / forget / what do you know**: hand over to `outlook-memory`.
- **Reranker consent**: `rerank.auto_consent true` lets outlook-search skip the per-run confirmation. Only set it when the user explicitly asks, and repeat once what will be sent and where.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-settings/reference.md` for the settings key reference and how to present `show`.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Writes are limited to `settings.json` inside `.outlook-skills/` folders.
