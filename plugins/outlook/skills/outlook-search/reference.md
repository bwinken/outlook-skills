# outlook-search reference

## 1. Script output (JSON)

```
{
  "Query":   { From, To, Subject, Body, Text, AnyOf[], After, Before, HasAttachments, Unread, Folders[], Dasl, Max },
  "Count":   <int>,
  "Results": [ <message summary>, ... ]     // newest first
}
```

Message summary (shared with outlook-thread):

| Field | Type | Notes |
|---|---|---|
| `EntryID` | string | pass to `outlook-thread -EntryID` |
| `Folder` | string | e.g. `\\Mailbox - Alice\Inbox\Projects` |
| `ReceivedTime` | ISO datetime | local time |
| `SentOn` | ISO datetime / null | |
| `From` | string | display name |
| `FromAddress` | string | SMTP where resolvable |
| `To`, `CC` | string | display strings as Outlook shows them |
| `Subject` | string | |
| `Unread` | bool | |
| `HasAttachments` | bool | |
| `Attachments[]` | `FileName`, `Size` (bytes), `Type` (1 file, 5 embedded, 6 OLE) | |
| `Size` | int | bytes, whole item |
| `Importance` | 0 low / 1 normal / 2 high | |
| `FlagStatus` | 0 none / 1 completed / 2 flagged | |
| `Categories` | string | comma separated |
| `ConversationID`, `ConversationTopic` | string | |
| `BodyPreview` | string | first `-PreviewLength` chars (default 200), whitespace collapsed |
| `Body` | string | only with `-IncludeBody` |

## 2. Presentation templates

Match the user's language; labels below are Traditional Chinese.

### 2a. Result list (default)

```
在 {Folders 縮寫，例如「收件匣」或「收件匣及其子資料夾」} 找到 {Count} 封{ ，顯示最新 N 封 — if Count == Max}：

| # | 日期 | 寄件者 | 主旨 | 附件 | 狀態 |
|---|---|---|---|---|---|
| 1 | 09/12 14:03 | 王小明 | Q3 預算討論 | 📎 2 | 未讀 |
| 2 | 09/11 09:20 | Alice Chen | Re: 合約草稿 | | |

{一句話補充，例如「最新一封是 Alice 昨天寄的合約草稿」— optional}
```

Rules:
- Date column: `MM/dd HH:mm` when within the current year, else `yyyy/MM/dd`.
- Sender column: display name; add the address in parentheses only if the name is ambiguous or the user asked for addresses.
- Subject: trim to about 60 characters with `…`.
- Attachment column: `📎 N` or blank. Status column: `未讀`, `🚩` for flagged, `‼` for high importance; blank otherwise.
- Show at most 20 rows. If there are more, say `還有 {Count-20} 封，需要的話可以縮小條件或告訴我要看哪幾封`.
- Do not print EntryIDs in the table. Keep them in context; offer `要看哪一封的完整內容或整串對話？` when relevant.
- When `Count` is 0: state the filters used in plain words and suggest one relaxation (wider dates, `-AllFolders`, fewer keywords).

### 2b. Single message card (when the user asks to read one mail, run with `-IncludeBody -Max 1`)

```
**{Subject}**
寄件者：{From} <{FromAddress}>
收件者：{To}
副本：{CC}                                 — omit if empty
時間：{yyyy/MM/dd HH:mm}
附件：{FileName} ({size, human readable}), …  — omit if none
資料夾：{Folder}

---
{Body, trimmed: drop quoted history below the first "From:"/"寄件者:" separator unless the user asks for it; cap at ~3000 chars with a note}
```

## 3. Reranker output (`rerank.py`)

```
{
  "Query":      <query text>,
  "Gateway":    <endpoint URL used>,
  "Endpoint":   "rerank" | "score",
  "Model":      <model name>,
  "Candidates": <int, mails sent>,
  "Batches":    <int, requests made>,
  "Count":      <int, results returned>,
  "Results":    [ <message summary> + "Score": <float>, ... ]   // highest score first
}
```

Presentation: same table as 2a with an extra leading `分數` column (two decimals), sorted by score. Add one line above the table:

```
已用 {Model} 對 {Candidates} 封候選郵件重新排序（{Batches} 批），最相關的前 {Count} 封：
```

Rules:
- Scores from bge-reranker-v2-m3 are not probabilities; show them for relative comparison only. If the best score is low (e.g. below 0.2) say the match is weak.
- Mention when the candidate set was cut by `-Max`: a relevant mail older than the window will not appear.
