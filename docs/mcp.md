# outlook-mcp

<p align="center"><a href="../README.md">📚 outlook skills</a>&ensp;|&ensp;<b>🔌 outlook-mcp</b></p>

`outlook-mcp` 是這個 marketplace 的第二個 plugin。它把 `outlook` skills 用的同一套 Python script 包成一個 **唯讀的 MCP server**，給不吃 skills、或想直接用 tool 的 client 用：Claude Chat、Cowork、Zoo Code、其他 MCP host。設定檔（`~/.outlook-skills/settings.json`）、唯讀政策都和 skills 相同。

**和 `outlook` skills 擇一安裝。** 兩個都裝也能動，但 MCP server 每個 session 都會啟動、九個 tool 一直佔 context；平常用 Claude Code 的人裝 skills 就夠了。

## 前置

- Windows、Classic Outlook（New Outlook 沒有 COM，不支援）。
- Python 3.8+，並在 **同一個 Python** 裡 `pip install pywin32`。
- 除了 Claude Code 的 plugin 安裝以外，其他 client 都要把 repo 放到本機固定位置，例如 `git clone https://github.com/bwinken/outlook-skills C:\tools\outlook-skills`（或 GitHub 頁面 Code → Download ZIP 解壓）。之後不要移動這個資料夾。

放好以後先跑一次，它會印出每種 client 要填的值（Python 絕對路徑、server.py 絕對路徑）：

```
python C:\tools\outlook-skills\plugins\outlook-mcp\install.py
```

用哪個 `python` 跑這行，印出來的就是哪個 Python 的路徑，所以請用裝了 pywin32 的那個。

## 各 client 怎麼裝

### Claude Code、Claude Desktop 的 Code 分頁

走 plugin marketplace，不用 clone：

```
/plugin marketplace add bwinken/outlook-skills
/plugin install outlook-mcp@outlook-skills
```

更新：`/plugin marketplace update outlook-skills`，再 `/plugin update outlook-mcp@outlook-skills`。plugin 用 PATH 上的 `python` 啟動 server，所以 pywin32 要裝在那個 Python。

不想用 marketplace（或 GitHub 連不上）：clone 之後

```
claude mcp add --scope user outlook -- C:\path\to\python.exe "C:\tools\outlook-skills\plugins\outlook-mcp\server.py"
```

### Claude Chat（claude.ai 桌面版 / Claude Desktop 的 Chat 分頁）、Cowork

在 Settings → Connectors 新增一個 connector，Transport 選 **Local command (stdio)**，欄位這樣填（`install.py` 會印出你機器上的實際路徑）：

| 欄位 | 填什麼 |
|---|---|
| Name | `outlook` |
| Transport | Local command (stdio) |
| Command | Python 可執行檔的**絕對路徑**，例如 `C:\Users\me\AppData\Local\Programs\Python\Python312\python.exe`。這個欄位不接受 `python` 這種相對名稱；查法：`python -c "import sys; print(sys.executable)"` |
| Arguments | `["C:\\tools\\outlook-skills\\plugins\\outlook-mcp\\server.py"]`（JSON 陣列，反斜線要寫兩個） |
| Environment variables | 通常留空。要用 reranker 又不想寫進 settings.json，可以在這裡設 `OUTLOOK_RERANK_URL`、`OUTLOOK_RERANK_MODEL`、`OUTLOOK_RERANK_API_KEY` |
| Environment helper script | 留空 |
| Startup timeout | 預設即可。server 啟動時不碰 Outlook，第一次呼叫 tool 才連，所以啟動很快 |
| Tool policy | 全部 tool 都是唯讀的，可以放心設成不必逐次核准；`find_attachments`（`saveto` 會把附件複製到磁碟）和 `parse_msg_file`（`extract_to` 同理）想保留核准就留給使用者控制 |

存檔後在 Chat 或 Cowork 的工具列會出現 outlook 的 tool。這是 Chat 分頁唯一能讀本機 Outlook 的方式；以 skill zip 上傳到 Chat 只有 `outlook-open-msg` 能用。

舊版 Claude Desktop 沒有這個表單，改編輯 `%APPDATA%\Claude\claude_desktop_config.json`：

```
python C:\tools\outlook-skills\plugins\outlook-mcp\install.py --claude-desktop
```

它會加一個 `outlook` 項目（原檔留 `.bak`），完全關閉 Claude Desktop 再開。移除：同一個指令加 `--uninstall`。

### Zoo Code（Roo Code 系）

MCP Servers 面板 → Edit Global MCP（或專案內的 `.roo/mcp.json`），加：

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\path\\to\\python.exe",
      "args": ["C:\\tools\\outlook-skills\\plugins\\outlook-mcp\\server.py"],
      "alwaysAllow": ["search_mail", "get_thread", "list_calendar", "list_followups", "prepare_meeting", "mailbox_overview", "get_status"],
      "disabled": false
    }
  }
}
```

`alwaysAllow` 列的是純讀取的 tool；`find_attachments` 和 `parse_msg_file` 可能寫檔，留給逐次核准。Zoo Code 也可以繼續用 skills（`python install.py` 裝到 `~/.roo/skills/`），兩者擇一。

### 其他 MCP host（Cursor、VS Code、Windsurf 等）

大多吃同一種 JSON：

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\path\\to\\python.exe",
      "args": ["C:\\tools\\outlook-skills\\plugins\\outlook-mcp\\server.py"]
    }
  }
}
```

`install.py` 不加參數就印這段，貼進該 host 的 MCP 設定即可。

## 檢查有沒有通

不經任何 client，直接呼叫一個 tool：

```
python C:\tools\outlook-skills\plugins\outlook-mcp\server.py --call get_status "{}"
```

看到帳號、store 清單就是 COM 正常。出現「pywin32 is not installed」就在同一個 Python 裝它；出現「Cannot start Outlook COM automation」表示不是 Classic Outlook 或被 AppLocker / WDAC 擋了，見 [install-troubleshooting.md](install-troubleshooting.md)。`--list` 可以印出全部 tool 和參數。

## 設定

讀的是和 skills 一樣的 `~/.outlook-skills/settings.json`（key 見主 README 的「設定」）。MCP 特別的兩點：

- `store`：有設就自動套到每個 tool，呼叫時沒指定 `store` 也沒 `allstores` 就用它。郵件在 .pst 的人請設。
- `rerank.gateway`（或環境變數 `OUTLOOK_RERANK_URL`）：有填時，`search_mail` 帶自然語言 `query` 就自動把候選（上限 `search.max_candidates`，預設 300）送去 gateway 排序，回傳前 `max` 筆（預設 10）並附 `Score`。**不會逐次問**，填了 gateway 就算同意。沒填時 `query` 被忽略，回一般子字串搜尋結果，`Rerank.Reason` 會說明。skills 用的 `ANTHROPIC_BASE_URL` 退路在 MCP 不適用，避免預覽被送到沒明確指定的 gateway。

記憶檔（人名、資料夾別名）和設定精靈是 skills 的功能，MCP 不讀。

## 功能對應

| Skill | 對應 MCP tool | 差異 |
|---|---|---|
| `outlook-status` | `get_status` | 相同 |
| `outlook-search` | `search_mail`、`find_attachments` | 模糊搜尋見上面的 rerank 說明 |
| `outlook-thread` | `get_thread` | 相同 |
| `outlook-agenda` | `list_calendar` | 相同 |
| `outlook-availability` | `list_calendar` | 空檔由模型從行事曆算，沒有 skill 裡的計算規則 |
| `outlook-morning-brief` | `list_calendar` + `search_mail` + `list_followups` | 沒有簡報流程，模型自己組合 |
| `outlook-meeting-prep` | `prepare_meeting` | 相同 |
| `outlook-open-msg` | `parse_msg_file` | 讀磁碟上的檔案；沒有釣魚信判讀指引 |
| `outlook-setup` | `mailbox_overview` | 只有掃描統計，沒有設定精靈、不寫記憶 |
| `outlook-memory` | 無 | skill 專屬 |

tool 參數、協定細節、維護方式見 [plugins/outlook-mcp/README.md](../plugins/outlook-mcp/README.md)。

## 更新、移除

- Claude Code plugin：`/plugin update outlook-mcp@outlook-skills`；移除 `/plugin uninstall outlook-mcp@outlook-skills`。
- 其他：到 repo 目錄 `git pull`（或重新下載 ZIP 覆蓋），重開 client。移除就把 connector / JSON 項目刪掉；舊版 Desktop 用 `install.py --claude-desktop --uninstall`。
