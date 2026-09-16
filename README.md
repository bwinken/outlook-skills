# outlook-skills

**Read-only** skills for a local Windows Outlook (Classic) mailbox, packaged as a Claude Code plugin marketplace.

## Install

| Host | 安裝 | 更新 |
|---|---|---|
| Claude Code | `pip install pywin32`<br>`/plugin marketplace add bwinken/outlook-skills`<br>`/plugin install outlook@outlook-skills` | `/plugin marketplace update outlook-skills`<br>`/plugin update outlook@outlook-skills` |
| Claude Desktop（Code 分頁） | 同 Claude Code | 同 Claude Code |
| Claude Desktop（Chat 分頁）/ claude.ai | `python outlook-skills/install.py --zip`，到 Customize → Skills → + 上傳 `dist/` 裡的 zip | 重新打包上傳 |
| MCP 版（Chat 分頁可讀本機信箱、其他 MCP host） | 見下方「也有 MCP 版」 | |

Chat 分頁與 Cowork 碰不到本機 Outlook，那裡只有 `outlook-open-msg`（解析附上的 .msg / .eml）能用；其他 skill 請在 Code 分頁或 Claude Code 使用（或改用下面的 MCP 版）。

<details>
<summary><b>🔌 也有 MCP 版</b> — 同一套功能包成 MCP server，Claude Desktop Chat 分頁和其他 MCP host 也能讀本機 Outlook。點開看安裝方式與功能對應表</summary>

<br>

`outlook-mcp` 是這個 marketplace 的第二個 plugin，跑的是同一套 Python script、同一份設定、同一條唯讀政策，差別是以 MCP tool 而不是 skill 接到 Claude。**和 `outlook` 擇一安裝**：兩個都裝也能動，但 MCP 的 tool 會一直佔 context。

### 安裝 MCP

前置：Windows、Classic Outlook、Python 3.8+ 在 PATH 上，然後 `pip install pywin32`。

**Claude Code / Claude Desktop Code 分頁**（走 plugin marketplace）：

```
/plugin marketplace add bwinken/outlook-skills
/plugin install outlook-mcp@outlook-skills
```

更新：`/plugin marketplace update outlook-skills`，再 `/plugin update outlook-mcp@outlook-skills`。

**Claude Desktop Chat 分頁**（Chat 分頁不吃 plugin，要登記到它的 MCP 設定檔）：

1. 把 repo 放到本機，之後不要移動：`git clone https://github.com/bwinken/outlook-skills C:\tools\outlook-skills`（或 GitHub 頁面 Code → Download ZIP 解壓）。
2. 寫入設定：`python C:\tools\outlook-skills\plugins\outlook-mcp\install.py --claude-desktop`。它會在 `%APPDATA%\Claude\claude_desktop_config.json` 加一個 `outlook` 項目（原檔留 `.bak` 備份）。裝了多個 Python 時加 `--python C:\path\to\python.exe` 指定有 pywin32 的那個。
3. 完全關閉 Claude Desktop 再開，Chat 分頁的工具列會出現 outlook 的 tool。

更新：到 repo 目錄 `git pull`，重開 Claude Desktop。移除：同一個指令加 `--uninstall`。

**其他 MCP host**（Cursor、VS Code 等）：`python plugins/outlook-mcp/install.py` 不加參數會印出這段，貼進該 host 的 MCP 設定即可：

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["C:\\tools\\outlook-skills\\plugins\\outlook-mcp\\server.py"]
    }
  }
}
```

沒有 marketplace 的 Claude Code 也可以：`claude mcp add --scope user outlook -- python "C:\tools\outlook-skills\plugins\outlook-mcp\server.py"`。

檢查有沒有通：`python plugins/outlook-mcp/server.py --call get_status "{}"`，看得到帳號和 store 就是 COM 正常；出現 pywin32 或 Classic Outlook 的錯誤訊息照 [docs/install-troubleshooting.md](docs/install-troubleshooting.md) 處理。

### 功能對應

| Skill | 對應 MCP tool | 差異 |
|---|---|---|
| `outlook-status` | `get_status` | 相同 |
| `outlook-search` | `search_mail`、`find_attachments` | 模糊搜尋：設定有填 `rerank.gateway` 就自動 rerank，不逐次問；沒填就只有子字串搜尋 |
| `outlook-thread` | `get_thread` | 相同 |
| `outlook-agenda` | `list_calendar` | 相同 |
| `outlook-availability` | `list_calendar` | 空檔由模型從行事曆算，沒有 skill 裡的計算規則 |
| `outlook-morning-brief` | `list_calendar` + `search_mail` + `list_followups` | 沒有簡報流程，模型自己組合 |
| `outlook-meeting-prep` | `prepare_meeting` | 相同 |
| `outlook-open-msg` | `parse_msg_file` | 讀磁碟上的檔案；沒有釣魚信判讀指引 |
| `outlook-setup` | `mailbox_overview` | 只有掃描統計，沒有設定精靈、不寫記憶 |
| `outlook-memory` | 無 | 記憶與設定精靈是 skill 專屬；`store` 設定 MCP 會自動套用 |

tool 參數與細節見 [plugins/outlook-mcp/README.md](plugins/outlook-mcp/README.md)。

</details>

裝不起來看 [docs/install-troubleshooting.md](docs/install-troubleshooting.md)（settings 衝突、proxy、離線手動安裝）。

## Skills

| Skill | Ask things like |
|---|---|
| `outlook-setup` | 第一次自動執行：逐項設定、掃描信箱建立記憶、統計回信習慣；之後可「重新設定」「重新掃描信箱」 |
| `outlook-status` | 我的 Outlook 資料檔在哪、信箱多大、有哪些帳號 |
| `outlook-search` | 找 Alice 上週寄給我的信、模糊搜尋「上次跟供應商談價格的信」、找附件、最大的附件、把附件存出來 |
| `outlook-thread` | 幫我摘要「Q3 預算」這串討論、結論和待辦是什麼 |
| `outlook-agenda` | 今天有什麼會議、這週行程、有沒有撞期、哪場還沒回覆 |
| `outlook-availability` | 禮拜三有沒有空、幫我找一小時的空檔、這週哪天下午有空 |
| `outlook-morning-brief` | 早安今天怎樣、有什麼新信、有沒有急事、誰還沒回我、我還欠誰回信 |
| `outlook-meeting-prep` | 幫我準備下一場會議、跟他們最近的往來和附件 |
| `outlook-open-msg` | 打開這個 .msg / .eml、看標頭判斷是不是釣魚信 |
| `outlook-memory` | 記住 Alice 是誰、忘掉、你記得什麼、設定工作時間、目前的設定 |

All skills only read. They never send, save, move, delete, flag or mark anything in Outlook. See [plugins/outlook/README.md](plugins/outlook/README.md) for requirements and the full read-only policy.

## 設定

`~/.outlook-skills/settings.json`（使用者層級）與 `./.outlook-skills/settings.json`（工作目錄層級，逐 key 覆蓋）。只寫你要改的 key，其餘用預設值；`settings.example.json` 列出全部。例如：

```json
{
  "store": "20230731",
  "working_hours": { "start": "09:00", "end": "17:30" },
  "search": { "default_lookback_days": 180 },
  "rerank": { "auto_consent": true }
}
```

| key | 預設 | 用途 |
|---|---|---|
| `store` | null | 預設信箱，郵件在 PST 時設成它的名稱 |
| `working_hours.start` / `.end` / `.days` | 09:00 / 18:00 / 週一到五 | 找空檔的範圍 |
| `availability.min_slot_minutes` | 30 | 短於此的空檔不列 |
| `search.default_lookback_days` | 90 | 沒說日期時的搜尋範圍 |
| `search.default_folder` / `.all_folders` | Inbox / false | 搜尋預設資料夾、是否含子資料夾 |
| `search.direct_read_max` | 20 | 結果不超過此數直接讀，不用 reranker |
| `rerank.gateway` / `.model` / `.api_key` / `.auto_consent` | null / bge-reranker-v2-m3 / null / false | reranker 設定；auto_consent 為 true 就不每次問 |
| `language` | zh-TW | 回覆語言 |

用 `outlook-memory` skill 改（「把工作時間改成 9 點到 5 點半」），或直接編輯檔案。

`outlook-mcp` 讀同一份設定：`store` 會自動套到每個 tool；`rerank.gateway` 有填時 `search_mail` 帶 `query` 就自動 rerank（不逐次問，填了就算同意），沒填就回一般子字串搜尋結果。

## 記憶

讓 skill 認得「Alice」是 alice.chen@contoso.com、「供應商的信」是 `Inbox/Vendors`。存在 `~/.outlook-skills/memory/<分類>/<標題>.md`，分類有 people、folders、projects、recurring、preferences，一個主題一檔，開頭是 YAML frontmatter：

```
---
title: Alice Chen
category: people
tags: [legal, contoso]
created: 2026-09-14T10:02:11
updated: 2026-09-14T13:40:05
---
- 法務窗口，alice.chen@contoso.com
```

- 第一次使用由 `outlook-setup` 引導：先問要不要開始，設定逐項用選單選，唯讀掃描最近 180 天的信箱（只看寄件者、主旨、資料夾、會議，不讀內文）後最多提出 10 條記憶，一條一條問你要不要加入。
- 平常說「記住 …」「忘掉 …」「你記得什麼」即可；相關的事會加進同一個檔案。
- 另有 `profile.md`：從寄件備份統計出來的回信習慣與寫作風格（回誰、多快、語氣、簽名），擬回信和判斷該不該回時會參考。
- 不存郵件內文、附件、金鑰；沒同意的不寫。記憶檔含人名與 email，工作目錄是 git repo 時把 `.outlook-skills/` 加進 `.gitignore`。

## Roadmap

- 規則稽核：列出 Outlook rules，標出自動轉寄到外部、自動刪除等可疑設定
- 聯絡人查詢
- 收件匣裡的信直接做標頭釣魚分析（目前只有 .msg / .eml 檔）
- 不開 Outlook 直接讀 .pst
