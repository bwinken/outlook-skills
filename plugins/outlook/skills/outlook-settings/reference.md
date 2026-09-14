# outlook-settings reference

## 1. settings.json keys

| Key | Default | Used by | Meaning |
|---|---|---|---|
| `language` | `zh-TW` | all | language of Claude's replies and table labels |
| `working_hours.start` / `.end` | `09:00` / `18:00` | availability | working window for free-slot search |
| `working_hours.days` | `[1,2,3,4,5]` | availability | ISO weekdays counted as working days (1 = Monday) |
| `availability.min_slot_minutes` | `30` | availability | gaps shorter than this are dropped |
| `search.default_lookback_days` | `90` | search | `-After` when the user gives no date hint |
| `search.default_folder` | `Inbox` | search, thread | `-Folder` when none is given |
| `search.all_folders` | `false` | search | pass `-AllFolders` by default |
| `search.max_candidates` | `300` | search | `-Max` for reranker candidate fetches |
| `search.direct_read_max` | `20` | search | up to this many hits are read directly instead of reranked |
| `store` | `null` | all COM skills | `-Store` for a shared or archive mailbox |
| `rerank.gateway` / `.model` / `.api_key` | `null` | search (rerank.py) | override the gateway resolution; beats `~/.claude/settings.json`, loses to env vars and flags |
| `rerank.auto_consent` | `false` | search | skip the per-run confirmation before sending candidates to the reranker |
| `status.skip_com` | `false` | status | always run `outlook-status` with `-SkipCom` (New Outlook machines) |

`settings.py show` returns `settings` (merged), `sources` (key → file for every non-default key), `layers` (files read), `memory` (memory.md and memory/*.md files found), `memory_stats` (entries and bytes per file), `memory_hint` (null, or a suggestion to prune or split), `local_dir`, `user_dir`.

## 2. memory.md structure

```
# Outlook memory
## 人物與別名
- Alice = Alice Chen <alice.chen@contoso.com>，法務窗口
## 資料夾
- Inbox/Vendors：供應商往來
## 專案關鍵字
- Q3 預算：方案 B，1.5M，VP review 9/19
## 偏好
- 表格日期用 MM/dd；回覆用繁體中文
```

Splitting when it grows: a folder may hold `memory/*.md` topic files instead of, or in addition to, `memory.md` (`people.md`, `folders.md`, `projects.md`, `preferences.md`). `settings.py show` lists all of them with entry counts and sets `memory_hint` once the total passes 300 entries; when you see the hint, propose pruning stale bullets or splitting, and do it only on the user's yes. Move bullets verbatim; never rewrite them while moving.

Rules:
- One bullet per fact, under the heading it belongs to; add a heading only if none fits.
- Keep bullets short (one line). Put the date in parentheses when the fact can go stale, e.g. `（2026/09）`.
- When a new fact contradicts an old bullet, replace the old one and say so.

## 3. Presenting `show`

```
## Outlook skills 設定

| 項目 | 值 | 來源 |
|---|---|---|
| 工作時間 | 09:00–17:30，週一到週五 | ~/.outlook-skills |
| 找空檔最短長度 | 45 分鐘 | ./.outlook-skills |
| 搜尋預設回溯 | 90 天 | 預設 |
| 預設資料夾 | Inbox | 預設 |
| 預設信箱 | （主信箱） | 預設 |
| Reranker | gateway http://vllm.internal:8000，模型 bge-reranker-v2-m3，每次詢問 | ./.outlook-skills |

記憶檔：~/.outlook-skills/memory.md（4 條）、./.outlook-skills/memory.md（2 條）
```

Show the api key as `已設定` / `未設定`, never its value. List only the rows in the table above; skip internal keys.
