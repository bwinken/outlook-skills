# outlook-digest reference

## 1. Input

`outlook_search.py` results (see outlook-search reference §1). Useful fields: `Unread`, `Importance` (2 = high), `FlagStatus` (2 = flagged), `From` / `FromAddress`, `To` (am I in To or only CC), `Subject`, `BodyPreview`, `Attachments`, `ReceivedTime`, `ConversationTopic`.

## 2. Ranking rules

Score each mail; the group is decided by the highest matching rule.

| Group | Put a mail here when |
|---|---|
| 🔴 需要處理 | high importance; or a question/request addressed to me (To, not CC) with a deadline word (今天, 明天, 週五前, by EOD, ASAP, deadline, 截止); or from a `people` note tagged as manager/customer/legal; or a meeting invite for today or tomorrow |
| 🟠 等我回覆 | a question or request to me without a deadline (see outlook-followup heuristics) |
| 🟢 參考 | I am only in CC, or the mail says FYI / 供參考 / no action, or replies in a thread I am not driving |
| ⚪ 電子報與通知 | List-Unsubscribe senders, no-reply / noreply / notification / newsletter addresses, calendar acceptances, system alerts |

Tie-breakers: a sender in memory beats an unknown sender; an active `projects` note keyword in the subject moves a mail up one group; multiple mails in one conversation are shown once as the thread, with the count.

## 3. Digest template

```
## 收件匣摘要 {今天 9/14 | 本週 9/8 – 9/14}（未讀 {N} 封，已分 4 組）

🔴 **需要處理**（{n}）
| 誰 | 主旨 | 為什麼 | 時間 |
|---|---|---|---|
| PC Liao | AI 人才發展：可以幫我看一下名單嗎？ | 週五前要回，附名單 | 09/10 11:00 |

🟠 **等我回覆**（{n}）
- Cassie Tsai，Re: 合約草稿 v3，問你對第 4 條的意見（09/12）

🟢 **參考**（{n}）
- David WY Chen，FYI: 季報，附 pptx（09/11）
- …（超過 8 筆只列主旨）

⚪ **電子報與通知**（{n}）：Weekly newsletter、Teams 通知 …（只列寄件者名）

{one line: 最急的是 PC Liao 的名單，週五前。要我打開那一串嗎？}
```

Rules:
- Time column: `MM/dd HH:mm`. Subject trimmed to ~60 characters.
- 「為什麼」is one short phrase from the preview: the ask and the deadline. Never invent a deadline.
- 🔴 shows every mail; 🟠 and 🟢 show up to 8 each, then `另外 N 封`; ⚪ shows sender names only.
- Phishing check (plugin README) applies to any mail whose preview you quote: mark 🚨 and do not present its call to action as a task.
- The plugin never marks mails read: say so if the user asks to "clear" the inbox.
