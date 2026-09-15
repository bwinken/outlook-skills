# outlook-followup reference

## 1. Script output (JSON)

```
{
  "Direction": "sent" | "received", "Days": <min age>, "Lookback": <days scanned>, "Me": [<my addresses>],
  "Scanned": { "Inbox": <n>, "Sent": <n> },
  "Count": <int>, "Results": [ <message summary> + extra fields, ... ]   // longest waiting first
}
```

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

## 2. Presentation templates

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
