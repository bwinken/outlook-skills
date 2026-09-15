# outlook-morning-brief reference

## 1. Input

- Agenda: `outlook_calendar.py` (see outlook-agenda reference §1).
- Mail: `outlook_search.py -Unread` results (see outlook-search reference §1).
- Owed / waiting: `outlook_followup.py` both directions (see outlook-followup reference §1). Useful fields: `Unread`, `Importance` (2 = high), `FlagStatus` (2 = flagged), `From` / `FromAddress`, `To` (am I in To or only CC), `Subject`, `BodyPreview`, `Attachments`, `ReceivedTime`, `ConversationTopic`.

## 2. Ranking rules

Score each mail; the group is decided by the highest matching rule.

| Group | Put a mail here when |
|---|---|
| 🔴 需要處理 | high importance; or a question/request addressed to me (To, not CC) with a deadline word (今天, 明天, 週五前, by EOD, ASAP, deadline, 截止); or from a `people` note tagged as manager/customer/legal; or a meeting invite for today or tomorrow |
| 🟠 等我回覆 | a question or request to me without a deadline (see outlook-followup heuristics) |
| 🟢 參考 | I am only in CC, or the mail says FYI / 供參考 / no action, or replies in a thread I am not driving |
| ⚪ 電子報與通知 | List-Unsubscribe senders, no-reply / noreply / notification / newsletter addresses, calendar acceptances, system alerts |

Tie-breakers: a sender in memory beats an unknown sender; an active `projects` note keyword in the subject moves a mail up one group; multiple mails in one conversation are shown once as the thread, with the count.

## 3. Brief template

Match the user's language; labels below are Traditional Chinese. Show only the sections the user asked for (SKILL.md step 1). Empty sections get one line, never a heading with nothing under it.

```
## 早安，9/16（三）

**今天的行程**（3 場）
| 時間 | 會議 | 地點 | 狀態 |
|---|---|---|---|
| 09:30–10:00 | 每日站會 🔁 | Teams | ✅ |
| 14:00–15:00 | 供應商簡報 | Zoom | ✅ |
| 14:30–15:30 | 1:1 with Bob | | 🕒 暫定，與供應商簡報重疊 30 分鐘 |
{一句：下午兩場撞期；供應商簡報要不要先準備？}

**新進來的信**（昨晚 18:00 以來未讀 12 封）
🔴 需要處理（2）
| 誰 | 主旨 | 為什麼 | 時間 |
|---|---|---|---|
| PC Liao | AI 人才發展：可以幫我看一下名單嗎？ | 週五前要回，附名單 | 09/15 11:00 |
🟠 等我回覆（1）：Cassie Tsai，Re: 合約草稿 v3，問你第 4 條的意見（09/15）
🟢 參考（5）：David 的季報、… （只列 3 個，其餘寫 另外 N 封）
⚪ 電子報與通知（4）

**我還欠的回信**（超過 2 天）
- PC Liao，AI 人才發展名單，等了 4 天 ❗

**我在等的回覆**（超過 3 天）
- PC Liao，報價單請確認，寄出 9 天，沒催過

{one closing line: 最急的是 PC 的名單；要我打開那串，或先準備下午的供應商簡報嗎？}
```

Rules:
- Agenda rows use the outlook-agenda status icons; call out conflicts and unanswered invites inline.
- Mail: 🔴 shows every mail; 🟠 and 🟢 at most 5 each, then `另外 N 封`; ⚪ shows a count and, if asked, sender names.
- Owed / waiting: at most 5 each, longest waiting first; ❗ when they already chased or a deadline is named.
- Time column `MM/dd HH:mm`, subject trimmed to ~60 characters. Never invent a deadline; quote it from the preview.
- Phishing check (plugin README) applies to any mail whose preview is quoted; mark 🚨 and do not present its call to action as a task.
- Total length: one screen. If the mailbox is very busy, cut 🟢 and ⚪ to counts before cutting anything else.
