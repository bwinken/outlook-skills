# outlook-skills

<p align="center"><b>📚 outlook skills</b>&ensp;|&ensp;<a href="docs/mcp.md">🔌 outlook-mcp</a></p>

讓 Claude **唯讀**你本機的 Windows Outlook（Classic）：找信、讀討論串、看行程、找空檔、早安簡報、會前準備、解析 .msg / .eml。
不寄信、不刪信、不標已讀、不動任何東西。

兩種包裝，擇一安裝：

| | 給誰 | 安裝 |
|---|---|---|
| **outlook**（skills） | Claude Code、Claude Desktop 的 Code 分頁 | `pip install pywin32`<br>`/plugin marketplace add bwinken/outlook-skills`<br>`/plugin install outlook@outlook-skills` |
| **outlook-mcp**（MCP server） | Claude Chat / Cowork、Zoo Code、其他 MCP host | 見 [docs/mcp.md](docs/mcp.md) |

更新：`/plugin marketplace update outlook-skills` 再 `/plugin update outlook@outlook-skills`。
需要 Windows + Classic Outlook + Python 3.8+（New Outlook 沒有 COM，不支援）。裝不起來看 [疑難排解](docs/install-troubleshooting.md)。

## 能問什麼

| Skill | 例子 |
|---|---|
| `outlook-setup` | 第一次自動執行：設定、掃描信箱建立記憶、統計回信習慣 |
| `outlook-status` | 我的 Outlook 資料檔在哪、信箱多大、有哪些帳號 |
| `outlook-search` | 找 Alice 上週寄的信、「上次跟供應商談價格的信」、最大的附件、把附件存出來 |
| `outlook-thread` | 幫我摘要「Q3 預算」這串討論，結論和待辦是什麼 |
| `outlook-agenda` | 今天有什麼會議、這週行程、有沒有撞期、哪場還沒回覆 |
| `outlook-availability` | 禮拜三有沒有空、幫我找一小時的空檔 |
| `outlook-morning-brief` | 早安今天怎樣、有沒有急事、誰還沒回我、我還欠誰回信 |
| `outlook-meeting-prep` | 幫我準備下一場會議 |
| `outlook-open-msg` | 打開這個 .msg / .eml、看標頭判斷是不是釣魚信 |
| `outlook-memory` | 記住 Alice 是誰、忘掉、你記得什麼、設定工作時間 |

Chat 分頁與 Cowork 碰不到本機 Outlook，只有 `outlook-open-msg` 能用；其他請用 Code 分頁、Claude Code，或改裝 MCP 版。

## 設定

`~/.outlook-skills/settings.json`，只寫要改的 key（工作目錄下的 `./.outlook-skills/settings.json` 逐 key 覆蓋）：

```json
{
  "store": "20230731",
  "working_hours": { "start": "09:00", "end": "17:30" },
  "search": { "default_lookback_days": 180 },
  "rerank": { "gateway": "http://gw:8000/v1", "auto_consent": true }
}
```

| key | 預設 | 用途 |
|---|---|---|
| `store` | null | 預設信箱；郵件在 PST 時填它的名稱 |
| `working_hours.start` / `.end` / `.days` | 09:00 / 18:00 / 1-5 | 找空檔的範圍 |
| `availability.min_slot_minutes` | 30 | 短於此的空檔不列 |
| `search.default_lookback_days` / `.default_folder` / `.all_folders` | 90 / Inbox / false | 沒說日期、資料夾時的搜尋範圍 |
| `search.direct_read_max` | 20 | 結果不超過此數就直接讀，不用 reranker |
| `rerank.gateway` / `.model` / `.api_key` / `.auto_consent` | null / bge-reranker-v2-m3 / null / false | 模糊搜尋用的 reranker；`auto_consent` 為 true 就不每次問 |
| `language` | zh-TW | 回覆語言 |

跟 Claude 說「把工作時間改成 9 點到 5 點半」即可，或直接改檔案。

## 記憶

`~/.outlook-skills/memory/<分類>/<標題>.md`，讓 skill 認得「Alice」是 alice.chen@contoso.com、「供應商的信」在 `Inbox/Vendors`。

- 第一次使用由 `outlook-setup` 引導，掃描信箱後逐條問你要不要記。
- 平常說「記住 …」「忘掉 …」「你記得什麼」。
- 不存內文、附件、金鑰；沒同意的不寫。檔案含人名與 email，在 git repo 裡請把 `.outlook-skills/` 加進 `.gitignore`。

## 開發

```
ruff check .                                            # lint
python -m unittest discover -s plugins/outlook/tests     # scripts（假 Outlook，任何 OS 都能跑）
python -m unittest discover -s plugins/outlook-mcp/tests # MCP server
python tools/sync_scripts.py                            # 改了 plugins/outlook/scripts 之後同步到 outlook-mcp
```

細節：[plugins/outlook/README.md](plugins/outlook/README.md)、[plugins/outlook-mcp/README.md](plugins/outlook-mcp/README.md)。

## Roadmap

- 規則稽核：列出 Outlook rules，標出自動轉寄到外部、自動刪除等可疑設定
- 聯絡人查詢
- 收件匣裡的信直接做標頭釣魚分析（目前只有 .msg / .eml 檔）
- 不開 Outlook 直接讀 .pst
