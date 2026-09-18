# outlook-skills

<p align="center"><b>📚 outlook skills</b>&ensp;|&ensp;<a href="docs/mcp.md">🔌 outlook-mcp</a></p>

讓 Claude 用你本機的 Windows Outlook（Classic）：找信、讀討論串、看行程、找空檔、早安簡報、解析 .msg / .eml，以及回信、寄信、約會議。

讀的部分唯讀；寄信和建立會議要你點兩次：先在對話裡看過草稿說好，再在桌面跳出的確認視窗按一次，才會真的送出。哪個 skill 會寫，見下面功能表。

## 安裝

前置：Windows、Classic Outlook（New Outlook 沒有 COM，不支援）、Python 3.8+。

**Claude Code、Claude Desktop 的 Code 分頁**

```
pip install pywin32
/plugin marketplace add bwinken/outlook-skills
/plugin install outlook@outlook-skills
```

Troubleshooting: [docs/install-troubleshooting.md](docs/install-troubleshooting.md)

## 更新

```
/plugin marketplace update outlook-skills
/plugin update outlook@outlook-skills
```

## 功能

| Skill | 問法 | 對 Outlook |
|---|---|---|
| `outlook-search` | 找 Alice 上週寄的信、「上次跟供應商談價格的信」、最大的附件、把附件存出來 | 唯讀 |
| `outlook-thread` | 幫我摘要「Q3 預算」這串討論，結論和待辦是什麼 | 唯讀 |
| `outlook-agenda` | 今天有什麼會議、這週行程、有沒有撞期、哪場還沒回覆 | 唯讀 |
| `outlook-availability` | 禮拜三有沒有空、幫我找一小時的空檔 | 唯讀 |
| `outlook-morning-brief` | 早安今天怎樣、有沒有急事、誰還沒回我、我還欠誰回信 | 唯讀 |
| `outlook-open-msg` | 打開這個 .msg / .eml、看標頭判斷是不是釣魚信 | 唯讀（不碰 Outlook） |
| `outlook-status` | 我的 Outlook 資料檔在哪、信箱多大、有哪些帳號 | 唯讀 |
| `outlook-memory` | 記住 Alice 是誰、忘掉、你記得什麼、設定工作時間 | 唯讀（只寫自己的設定檔） |
| `outlook-setup` | 重新設定、重新掃描信箱 | 唯讀（只寫自己的設定檔） |
| `outlook-send` | 回他說好、幫我回這封、把桌面上的報價單寄給 Alice | **會寄信**：對話確認＋桌面視窗按「寄出」 |
| `outlook-schedule` | 幫我約 Cassie 週四下午開會、發個邀請、把週五早上擋起來 | **會建立會議、送邀請**：對話確認＋桌面視窗按「送出」 |

讀信箱走 Outlook 的 Table 物件，一次拿回幾百封的摘要，只有要看內文或附件的那幾封才會真的打開；Outlook 沒開時腳本會在背景啟動它，第一次會多等幾秒。

唯讀的 skill 不刪信、不搬信、不標已讀。會寫的兩個，寄出的內容和你在草稿看到的完全一樣，頁尾註明 Drafted by Claude, approved by 你；搜尋的「把附件存出來」只寫到你指定的資料夾。

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
| `search.default_lookback_days` / `.default_folder` / `.all_folders` | 90 / Inbox / false | 沒說日期、資料夾時的搜尋範圍；`store`、`default_folder`、`all_folders` 腳本會自己套用 |
| `search.direct_read_max` | 20 | 結果不超過此數就直接讀，不用 reranker |
| `rerank.gateway` / `.model` / `.api_key` / `.auto_consent` | null / bge-reranker-v2-m3 / null / false | 模糊搜尋用的 reranker；`auto_consent` 為 true 就不每次問 |
| `send.approver` / `.footer` / `.quote_original` / `.dialog_timeout_seconds` | Outlook 使用者名稱 / Drafted by Claude… / true / 300 | 寄信與會議頁尾的核准者與文字、回信是否引用原信、確認視窗等多久 |
| `meeting.default_duration_minutes` / `.reminder_minutes` | 60 / 15 | 沒說多久時的會議長度、提醒 |
| `language` | zh-TW | 回覆語言 |

改設定不用開檔案，直接說：

> 把工作時間改成 9 點到 5 點半

記憶在 `~/.outlook-skills/memory/<分類>/<標題>.md`，讓 skill 認得「Alice」是誰、「供應商的信」在哪個資料夾。也是用說的：

> 記住 Alice 是法務窗口，alice.chen@contoso.com
>
> 忘掉 Alice
>
> 你記得什麼

沒同意的不寫；不存內文、附件、金鑰。檔案含人名與 email，工作目錄是 git repo 時把 `.outlook-skills/` 加進 `.gitignore`。
