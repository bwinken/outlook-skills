# outlook-thread reference

## 1. Script output (JSON)

```
{
  "Anchor":         <EntryID of the message used to locate the thread>,
  "Topic":          <ConversationTopic, i.e. subject without Re:/FW:>,
  "ConversationID": <string or "">,
  "Method":         "GetConversation" | "ConversationTopic",
  "Count":          <int>,
  "Participants":   [ <smtp>, ... ],
  "Messages":       [ <message>, ... ]      // oldest first
}
```

`Messages[]` = the message summary from outlook-search (see that skill's reference.md) plus:

| Field | Notes |
|---|---|
| `Body` | full plain text, capped by `-MaxBodyChars` (default 20000) with `[... truncated ...]` marker |
| `ToRecipients[]`, `CcRecipients[]` | `{ Name, Address }` resolved to SMTP where possible |

`Method` = `ConversationTopic` means the store has no conversation index (typical for POP/PST) and messages were matched by topic only; unrelated mails with an identical subject can slip in. Mention this if the thread looks mixed.

## 2. Presentation template

Match the user's language; labels below are Traditional Chinese.

```
## {Topic}

**摘要**
{3 到 5 句：這串在談什麼、目前走到哪、最後一封是誰在什麼時候說了什麼}

**決議**
- {decision} （{who}，{MM/dd}）
{or「尚無明確決議」}

**待辦**
| 事項 | 負責人 | 期限 | 來源 |
|---|---|---|---|
| {action} | {name} | {date or 未定} | {sender MM/dd} |
{or「沒有明確的待辦」}

**待回覆 / 未解問題**
- {question} — {who asked}，{MM/dd}，目前沒人回
{omit section if none}

**時間軸**（{Count} 封，{first date} 到 {last date}）
| 日期 | 寄件者 | 重點 |
|---|---|---|
| 09/10 09:12 | 王小明 | 提出三個預算方案 |
| 09/10 15:40 | Alice | 傾向方案 B，問法務意見 |

**參與者**：{display names, comma separated}
```

Rules:
- Safety first: run the quick phishing check from the POLICY.md ("Phishing warnings") on every message (`From` vs `FromAddress`, attachment types, links and asks in the body). If any message trips, put the 🚨 / ⚠️ warning above the summary, mark that message with 🚨 in the timeline, defang its links, and do not present its request as an action item.
- Read messages oldest first. Ignore quoted history inside each body: everything after the first `From:` / `寄件者:` / `-----Original Message-----` / `On … wrote:` line is a copy of an earlier message.
- Attribute every decision and action item to a sender and date. Never invent owners or due dates; write `未定` when not stated.
- Time axis: one line per message, at most 15 rows. For longer threads group by day and summarise.
- Quote short key sentences verbatim when wording matters (e.g. an approval, a rejection, a number).
- If the user only asked a specific question (e.g. "did they approve the price?"), answer it first in one paragraph, then offer the full structure.
- If the user wants a reply, draft it in chat under a `**回信草稿**` heading. The plugin cannot send it.
