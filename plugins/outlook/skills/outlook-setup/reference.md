# outlook-setup reference

All questions go through the host's structured question tool (AskUserQuestion). Labels below are Traditional Chinese; translate for an English user.

## 1. Stage 1 question

```
question: 這是第一次使用 Outlook skills，要現在完成設定嗎？
接下來會做三件事：
1. 逐項設定（預設信箱、工作時間、搜尋範圍、語言、reranker），每項都用選單選。
2. 唯讀掃描最近 180 天的信箱（只看寄件者、主旨、資料夾、會議，不讀內文），整理出最多 10 條記憶，逐條問你要不要加入。
3. 從寄件備份統計你的回信習慣（回誰、多快、語氣、簽名），存成側寫檔。
預估 {N} 分鐘（信箱 {store} 約 {items} 封）。沒有你點「加入」「存」之前不會寫任何檔案，Outlook 本身完全不會被更動。
options:
  開始設定 (Recommended)   — 三個階段都做
  只設定，不掃描信箱       — 只做第 1 步
  這次先不要               — 什麼都不建立
```

## 2. Stage 2 items

One question each; batch up to 4 per dialog. Options are built from real data where noted. First option is the recommendation, or the current value on a re-run (label it 「目前：…」).

| # | key | question | options |
|---|---|---|---|
| 1 | `store` | 主要的信箱是哪一個？ | each store from `outlook_status.py` as 「{DisplayName}（{SizeMB} MB，收件匣 {n} 封）」, the largest by items marked (Recommended); plus 「Exchange 預設信箱」 |
| 2 | `working_hours.start` / `.end` | 工作時間？（找空檔用） | 09:00–18:00 (Recommended) / 09:00–17:30 / 08:30–17:30 / 其他（用 Other 輸入） |
| 3 | `working_hours.days` | 工作日？ | 週一到週五 (Recommended) / 週一到週六 / 其他 |
| 4 | `availability.min_slot_minutes` | 找空檔時，多短的空檔可以忽略？ | 30 分鐘 (Recommended) / 15 分鐘 / 60 分鐘 |
| 5 | `search.default_lookback_days` | 沒說日期時，搜尋預設回溯多久？ | 90 天 (Recommended) / 30 天 / 180 天 / 365 天 |
| 6 | `search.all_folders` | 搜尋預設要不要包含收件匣的子資料夾？ | 包含 (Recommended if outlook_status shows subfolders with mail) / 只搜收件匣 |
| 7 | `language` | 回覆語言？ | 繁體中文 (Recommended) / English / 跟著我的訊息 |
| 8 | `rerank.auto_consent` | 模糊搜尋時可以直接送候選郵件的主旨與預覽到 {gateway}（{model}）重新排序嗎？ — only when `--show-config` is usable | 每次先問我 (Recommended) / 可以，不用每次問 / 不要用 reranker（set `rerank.gateway` to `"none"` is not needed; just leave auto_consent false and note the preference in profile.md） |

Write with `settings.py set <key> <value>`; `working_hours.days` as a JSON array, e.g. `[1,2,3,4,5]`. After the last item show:

```
設定完成（~/.outlook-skills/settings.json）：
| 項目 | 值 |
|---|---|
| 主要信箱 | 20230731 |
| 工作時間 | 09:00–17:30，週一到週五 |
| 搜尋回溯 | 90 天，含子資料夾 |
| 語言 | 繁體中文 |
| Reranker | 每次先問 |
```

## 3. Stage 3 candidates

Pick at most 10 from the overview, in this priority: top correspondents who are both sender and recipient (people), folders with many items and a clear purpose (folders), the busiest 2 or 3 conversation topics that are not automated (projects), every recurring meeting (recurring). Never newsletters. Title rules and per-category content: see `../outlook-memory/reference.md` §2 and §3.

Per-candidate question (up to 4 per dialog):

```
question: 加入記憶？[people] Cassie Tsai
tags: legal, contoso
內容：法務窗口，cassie.tsai@contoso.com，近半年 42 封往來，多為合約與法務意見
options: 加入 (Recommended) / 跳過 / 改一下再加
```

After the dialogs: `已寫入 {N} 條到 ~/.outlook-skills/memory/（people {a}、folders {b}、projects {c}、recurring {d}），跳過 {M} 條。`

## 4. Stage 4 profile.md

### `outlook_style.py` output (JSON)

| Field | Meaning |
|---|---|
| `Window` | `Since`, `Days`, `SentAnalysed`, `ReceivedAnalysed` (mails scanned), `BodySamples` (newest sent mails whose text was read for length, language, greetings, closings and signature; `-BodySamples`, default 300) |
| `Me` | the user's addresses (accounts + current user) |
| `ReplyRate.OverallHuman` | replied ÷ received, newsletters excluded (0 to 1) |
| `ReplyRate.BySender[]` | `Name`, `Address`, `Received`, `Replied`, `ToMe` (times the user was in To), `Rate`, `Newsletter` |
| `ReplyRate.AlwaysReplied[]` | senders with ≥2 mails and rate ≥ 0.8 |
| `ReplyRate.NeverReplied[]` | senders with ≥3 mails and 0 replies |
| `ReplyRate.Newsletters[]` | addresses with a List-Unsubscribe header |
| `ReplyLatencyHours` | `Median`, `P75`, `Within1h` (share), `Samples` |
| `Length` | `MedianChars` of the user's own text (quoted history removed), `Buckets` Short<200 / Medium200-800 / Long>800 |
| `Language` | `MedianCjkRatio`, `MostlyChinese` |
| `Greetings[]` / `Closings[]` | recurring first / last lines (`Line`, `Count`), 40 characters max |
| `Signature[]` | recurring trailing block (`Block` with lines joined by " / ", `Count`) |
| `SendHours` | `Histogram` by hour, `Typical` (the 4 busiest hours) |

"Replied" means a later mail from the user exists in the same conversation (Inbox + Sent Items). Replies by phone or chat are invisible; say so when a well-known contact shows up under "never replied".

### Template

Built from the JSON above. Show it in full, then ask 「存成 profile.md」/「修改後再存」/「不要存」.

```
---
title: 回信習慣與寫作風格
source: outlook_style
window_days: 180
created: 2026-09-16T09:12:00
updated: 2026-09-16T09:12:00
---
## 回誰、不回誰
- 整體回覆率（排除電子報）：{OverallHuman × 100}%
- 幾乎都回：{AlwaysReplied: name（rate）...}
- 從不回：{NeverReplied: name（received 封）...}
- 電子報／通知：{Newsletters 數量} 個來源，一律不回

## 回得多快
- 中位數 {Median} 小時，75% 在 {P75} 小時內回；{Within1h × 100}% 在一小時內

## 長度與語言
- 中位數 {MedianChars} 字；短信（<200 字）{n} 封、中等 {n}、長信 {n}
- 主要語言：{繁體中文 | 英文}（中文字比例 {ratio}）

## 開頭與結尾
- 常用開頭：{Greetings top 3}
- 常用結尾：{Closings top 3}
- 簽名：{Signature top 1, lines joined by " / "}

## 寄信時段
- 多在 {Typical hours} 點寄出

## 給其他 skill 的提示
- 擬回信草稿時用上面的開頭、結尾、長度與語言。
- 判斷「該不該回」時：從不回的寄件者與電子報不列入待回覆；幾乎都回的人優先。
```

Rules:
- Percentages rounded to whole numbers; hours to one decimal.
- Names in the profile come from `outlook_style.py` (`Name` field) with the address in parentheses only when ambiguous.
- Do not include any mail text beyond the recurring greeting, closing and signature lines the script already shortened.
- On a re-run, keep `created`, refresh `updated`, and show a three-line "what changed" note before asking to save.
