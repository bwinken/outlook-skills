# outlook-skills

<p align="center"><b>📚 outlook skills</b>&ensp;|&ensp;<a href="docs/mcp.md">🔌 outlook-mcp</a></p>

讓 Claude **唯讀**你本機的 Windows Outlook（Classic）：找信、讀討論串、看行程、找空檔、早安簡報、會前準備、解析 .msg / .eml。不寄信、不刪信、不標已讀。

## 安裝

前置：Windows、Classic Outlook（New Outlook 沒有 COM，不支援）、Python 3.8+。

**Claude Code、Claude Desktop 的 Code 分頁**

```
pip install pywin32
/plugin marketplace add bwinken/outlook-skills
/plugin install outlook@outlook-skills
```

裝好後直接問，例如「今天有什麼會議」。第一次會自動跑 `outlook-setup` 引導設定。

**Claude Chat、Cowork、Zoo Code、其他 MCP host**

改裝 MCP 版 `outlook-mcp`，步驟在 [docs/mcp.md](docs/mcp.md)。和 skills 版擇一。

**Claude Desktop Chat 分頁 / claude.ai 上傳 skill**

```
python install.py --zip
```

到 Customize → Skills → + 上傳 `dist/` 裡的 zip。那裡碰不到本機 Outlook，只有 `outlook-open-msg`（解析 .msg / .eml）能用。

裝不起來（settings 衝突、proxy、離線）看 [docs/install-troubleshooting.md](docs/install-troubleshooting.md)。

## 更新

```
/plugin marketplace update outlook-skills
/plugin update outlook@outlook-skills
```

MCP 版：`/plugin update outlook-mcp@outlook-skills`，非 Claude Code 的 host 則 `git pull`。zip 上傳的重新打包再上傳。

## 功能

| Skill | 問法 |
|---|---|
| `outlook-search` | 找 Alice 上週寄的信、「上次跟供應商談價格的信」、最大的附件、把附件存出來 |
| `outlook-thread` | 幫我摘要「Q3 預算」這串討論，結論和待辦是什麼 |
| `outlook-agenda` | 今天有什麼會議、這週行程、有沒有撞期、哪場還沒回覆 |
| `outlook-availability` | 禮拜三有沒有空、幫我找一小時的空檔 |
| `outlook-morning-brief` | 早安今天怎樣、有沒有急事、誰還沒回我、我還欠誰回信 |
| `outlook-meeting-prep` | 幫我準備下一場會議 |
| `outlook-open-msg` | 打開這個 .msg / .eml、看標頭判斷是不是釣魚信 |
| `outlook-status` | 我的 Outlook 資料檔在哪、信箱多大、有哪些帳號 |
| `outlook-memory` | 記住 Alice 是誰、忘掉、你記得什麼、設定工作時間 |
| `outlook-setup` | 重新設定、重新掃描信箱 |

## 設定與記憶

設定在 `~/.outlook-skills/settings.json`，只寫要改的 key（`./.outlook-skills/settings.json` 逐 key 覆蓋）：

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

跟 Claude 說「把工作時間改成 9 點到 5 點半」即可。

記憶在 `~/.outlook-skills/memory/<分類>/<標題>.md`，讓 skill 認得「Alice」是誰、「供應商的信」在哪個資料夾。說「記住 …」「忘掉 …」「你記得什麼」；沒同意的不寫，不存內文、附件、金鑰。檔案含人名與 email，在 git repo 裡請把 `.outlook-skills/` 加進 `.gitignore`。

## 開發

```
ruff check .
python -m unittest discover -s plugins/outlook/tests       # 假 Outlook，任何 OS
python -m unittest discover -s plugins/outlook-mcp/tests
python tools/sync_scripts.py                               # 改了 plugins/outlook/scripts 後同步到 outlook-mcp
```

細節：[plugins/outlook/README.md](plugins/outlook/README.md)、[plugins/outlook-mcp/README.md](plugins/outlook-mcp/README.md)。

## Roadmap

目前只讀。下一步加寫入功能，每個動作都先給你看內容、確認後才執行：

- 寄信
- 回信
- 建立會議
