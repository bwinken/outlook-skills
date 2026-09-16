# outlook-morning-brief reference

## 1. Input

- Agenda: `outlook_calendar.py` (see outlook-agenda reference §1).
- Mail: `outlook_search.py -Unread` results (see outlook-search reference §1).
- Owed / waiting: `outlook_followup.py` both directions (see outlook-morning-brief reference §1). Useful fields: `Unread`, `Importance` (2 = high), `FlagStatus` (2 = flagged), `From` / `FromAddress`, `To` (am I in To or only CC), `Subject`, `BodyPreview`, `Attachments`, `ReceivedTime`, `ConversationTopic`.

## 2. Ranking rules

Score each mail; the group is decided by the highest matching rule.

| Group | Put a mail here when |
|---|---|
| 🔴 需要處理 | high importance; or a question/request addressed to me (To, not CC) with a deadline word (今天, 明天, 週五前, by EOD, ASAP, deadline, 截止); or from a `people` note tagged as manager/customer/legal; or a meeting invite for today or tomorrow |
| 🟠 等我回覆 | a question or request to me without a deadline (see outlook-morning-brief heuristics) |
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
- Phishing check (POLICY.md) applies to any mail whose preview is quoted; mark 🚨 and do not present its call to action as a task.
- Total length: one screen. If the mailbox is very busy, cut 🟢 and ⚪ to counts before cutting anything else.

## 4. Follow-ups on their own (`outlook_followup.py`)

```
{
  "Direction": "sent" | "received", "Days": <min age>, "Lookback": <days scanned>, "Me": [<my addresses>],
  "Scanned": { "Inbox": <n>, "Sent": <n> },
  "Count": <int>, "Results": [ <message summary> + extra fields, ... ]   // longest waiting first
}
```

With `-Direction both` the top level has `Direction: "both"`, `Lookback`, `Me`, `Scanned`, and the two lists as `Sent` and `Received`, each `{ Days, Count, Results }` shaped as above.

Extra fields on each result (message summary fields are the same as outlook-search):

| Direction | Field | Meaning |
|---|---|---|
| both | `WaitingDays` | days since the mail, whole days |
| sent | `Counterparts` | `[{Name, Address}]` the To recipients who owe the reply |
| sent | `MyLaterNudges` | how many later mails I sent in the same conversation (already chased) |
| received | `LooksLikeQuestion` | heuristic: `?`, 請問/麻煩/煩請/可否, please/could you/let me know ... |
| received | `DirectToMe` | I am in To (not only CC) |
| received | `TheirLaterNudges` | later mails from the same sender in the conversation (they chased) |

A conversation is matched by `ConversationID`, falling back to `ConversationTopic`. "Answered" means a later mail from the other side exists in Inbox or Sent Items within the lookback window; replies stored elsewhere are not seen.


Match the user's language; labels below are Traditional Chinese.

### sent: 誰還沒回我

```
最近 60 天寄出、超過 3 天沒收到回覆的信有 {Count} 封：

| 等了 | 對象 | 主旨 | 寄出 | 已催 |
|---|---|---|---|---|
| 9 天 | PC Liao | 報價單請確認 | 09/05 | |
| 5 天 | David WY Chen, PC Liao | Re: Q3 預算討論 | 09/09 | 1 次 |

要我打開哪一串，或幫你擬一封催促的草稿？（草稿只會出現在這裡，不會寄出）
```

### received: 我還欠誰回信

```
別人寄給你、超過 2 天還沒回的信有 {N} 封，其中 {M} 封看起來是在等你回答：

**看起來在等你回答**
| 等了 | 誰 | 主旨 | 他們要什麼 |
|---|---|---|---|
| 4 天 | PC Liao | AI 人才發展：可以幫我看一下名單嗎？ | 週五前看名單 |

**可能只是 FYI**（只列主旨）
- David WY Chen，FYI: 季報（09/11）
```

Rules:
- Sort by `WaitingDays` descending. Show at most 15 rows per table; say how many more.
- "他們要什麼" is one short phrase distilled from `BodyPreview`; write `（不確定）` if the preview does not say.
- A mail with `TheirLaterNudges` > 0 gets ❗ before the subject: they already chased.
- Newsletters and automated senders (no-reply, noreply, notification) go to the FYI list without comment.
- Never mark anything as answered or flag it; the plugin only reads.
