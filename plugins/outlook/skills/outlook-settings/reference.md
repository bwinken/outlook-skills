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

`settings.py show` returns `first_run` (true when `~/.outlook-skills` does not exist yet), `settings` (merged), `sources` (key → file for every non-default key), `layers` (files read), `memory` (index: count, per_category, entries with title/category/tags/updated/path), `memory_hint`, `local_dir`, `user_dir`.

## 2. Presenting `show`

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

記憶：~/.outlook-skills/memory/ 共 29 條（people 12、folders 6、projects 8、recurring 3）；用 outlook-memory 查看或修改
```

Show the api key as `已設定` / `未設定`, never its value. List only the rows in the table above; skip internal keys.
