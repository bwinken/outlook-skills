# outlook-skills

**Read-only** skills for a local Windows Outlook (Classic) mailbox, packaged as a Claude Code plugin marketplace and installable into Zoo Code (Roo Code's successor) or any Agent Skills host.

## Install

| Host | 安裝 | 更新 |
|---|---|---|
| Claude Code | `/plugin marketplace add bwinken/outlook-skills`<br>`/plugin install outlook@outlook-skills` | `/plugin marketplace update outlook-skills`<br>`/plugin update outlook@outlook-skills` |
| Zoo Code | `git clone https://github.com/bwinken/outlook-skills`<br>`python outlook-skills/install.py` | `git -C outlook-skills pull`<br>`python outlook-skills/install.py` |
| Claude Desktop / claude.ai | `python outlook-skills/install.py --zip`，到 Customize → Skills → + 上傳 `dist/` 裡的 zip | 重新打包上傳 |

Claude Desktop 的聊天 skill 與 Cowork 跑在沙箱裡，碰不到本機 Outlook，所以那裡只有 `outlook-open-msg`（解析附上的 .msg / .eml）能完整使用；其他 skill 需要 Claude Code 或 Zoo Code 在同一台 Windows 上執行。

裝不起來看 [docs/install-troubleshooting.md](docs/install-troubleshooting.md)（settings 衝突、proxy、離線手動安裝）。

## Skills

| Skill | Ask things like |
|---|---|
| `outlook-status` | 我的 Outlook 資料檔在哪、信箱多大、有哪些帳號 |
| `outlook-search` | 找 Alice 上週寄給我的信、有附件的未讀郵件、模糊搜尋「上次跟供應商談價格的信」（可選 reranker） |
| `outlook-thread` | 幫我摘要「Q3 預算」這串討論、結論和待辦是什麼 |
| `outlook-agenda` | 今天有什麼會議、這週行程、有沒有撞期、哪場還沒回覆 |
| `outlook-availability` | 禮拜三有沒有空、幫我找一小時的空檔、這週哪天下午有空 |
| `outlook-open-msg` | 打開這個 .msg / .eml、看標頭判斷是不是釣魚信 |
| `outlook-settings` | 設定工作時間、目前用什麼設定、reranker gateway |
| `outlook-memory` | 第一次使用時建立個人化記憶、記住 Alice 是誰、忘掉、你記得什麼 |

All skills only read. They never send, save, move, delete, flag or mark anything in Outlook. See [plugins/outlook/README.md](plugins/outlook/README.md) for requirements, the full read-only policy, and how the skills cope with a locked-down PowerShell execution policy.

## 設定

設定放在 `settings.json`，兩層，工作目錄那層逐 key 覆蓋使用者那層：

| 位置 | 範圍 |
|---|---|
| `~/.outlook-skills/settings.json` | 使用者，所有專案共用 |
| `./.outlook-skills/settings.json`（工作目錄或往上任一層） | 這個專案 |

可以設定的有工作時間、找空檔的最短長度、搜尋預設回溯天數與資料夾、預設信箱、reranker gateway 與是否免每次詢問、回覆語言。用 `outlook-settings` skill 改，或直接編輯檔案。

## 記憶

記憶讓 skill 認得你的說法。你說「找 Alice 的信」時，Claude 需要知道 Alice 是 alice.chen@contoso.com；你說「供應商的信」時，需要知道那是 `Inbox/Vendors`。這些對應關係存在本機的 Markdown 檔，不在 Outlook 裡。

**位置與格式**：`~/.outlook-skills/memory/<分類>/<標題>.md`，一個主題一個檔，分類有 people（人物）、folders（資料夾）、projects（專案與主題）、recurring（定期事務）、preferences（偏好）。每個檔案開頭有 YAML frontmatter：

```
---
title: Alice Chen
category: people
tags: [legal, contoso]
created: 2026-09-14T10:02:11
updated: 2026-09-14T13:40:05
source: bootstrap
---
- 法務窗口，alice.chen@contoso.com
- 合約相關的信都由她發起（2026/09）
```

標題就是以後用來找它的名字（人名、資料夾路徑、專案短名、會議主旨）。之後有相關的事要記，會直接加進同一個檔案，並更新 `updated`。

**第一次使用**：任何 skill 發現 `~/.outlook-skills` 不存在時，會先問你要不要建立個人化記憶，三個選項：建立並掃描信箱、只建立空的、這次先不要。選擇掃描時，會唯讀彙整最近 180 天的常聯絡人、資料夾用途、常見主題、定期會議（只看計數與主旨，不讀內文），整理成草稿表格給你看，再問要全部寫入、讓你挑、還是取消。沒有你的同意不會寫任何檔案。

**平常使用**：
- 「記住 Alice 是法務窗口」：先找有沒有已存在的 Alice 檔，有就追加一行，沒有就新建。
- 「忘掉 X」：刪掉那個檔或那一行，會告訴你刪了什麼。
- 「你記得什麼？」：依分類列出所有標題。
- 「重新掃描信箱」：再跑一次彙整，只提議新的或有變的內容。
- 其他 skill 執行時只讀索引（標題、分類、tags），遇到需要的名字才打開對應的檔案，不會每次載入全部。

**不會存進記憶的**：郵件內文、附件、任何金鑰，以及你沒同意要記的東西。記憶檔會有人名和 email，若工作目錄是 git repo，記得把 `.outlook-skills/` 加進 `.gitignore`。

## Layout

```
.claude-plugin/marketplace.json   marketplace manifest
plugins/outlook/                  the plugin (skills + scripts)
install.py                        Zoo Code installer (shortcut to plugins/outlook/install.py)
```

## Roadmap ideas

Not implemented yet, kept here as candidates: follow-up tracker (sent mails with no reply), inbox digest, meeting prep, attachment finder, mail-header phishing triage, rules audit (detect suspicious auto-forward rules), contacts lookup, .pst archive explorer.
