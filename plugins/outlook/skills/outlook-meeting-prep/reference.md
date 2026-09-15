# outlook-meeting-prep reference

## 1. Script output (JSON)

```
{
  "Meeting":      <appointment summary: Subject, Start, End, Location, Organizer, RequiredAttendees, ResponseStatus ...>,
  "Attendees":    [ { Name, Address }, ... ],          // me excluded
  "Since":        "<date>", "Keywords": [ ... ],       // words from the subject used for the topic search
  "ByAttendee":   [ { Attendee, Count, Mails: [ <message summary>, ... ] }, ... ],   // newest first, from and to them
  "AboutSubject": [ <message summary>, ... ],          // mails matching the keywords, not already listed above
  "Attachments":  [ { FileName, SizeKB, From, ReceivedTime, Subject, EntryID }, ... ]
}
```

Message summary fields are the outlook-search ones (`BodyPreview` is 300 characters here).

## 2. Briefing template

Match the user's language; labels below are Traditional Chinese.

```
## 會前簡報：{Subject}
{9/16（三）14:00–15:00 · Zoom · 主辦 Cassie Tsai · 你 ✅ 已接受}

**與會者**
| 誰 | 角色（來自記憶） | 最近往來 |
|---|---|---|
| Cassie Tsai | 法務窗口 | 3 封，最新 09/12 合約草稿 v3 |
| PC Liao | 人資 | 2 封，最新 09/10 名單，你還沒回 |

**跟每個人還開著的事**
- **Cassie**：09/12 送來法務版合約 v3，改了第 4 條付款條件，等你的意見。
- **PC**：09/10 請你看人才名單，週五前，你尚未回覆 ❗

**這個主題的討論**（{n} 封）
- 09/08 David 提出三個方案；09/11 定案方案 B（1.5M）。

**建議先開的附件**
| 檔案 | 來自 | 日期 | 為什麼 |
|---|---|---|---|
| 合約草稿_v3_legal.docx | Cassie | 09/12 | 會上要談第 4 條 |
| 人才名單.xlsx | PC | 09/10 | 你被問到的名單 |

**可以在會上提的問題**
- 第 4 條的付款條件財務同意了嗎？
- 名單的截止是週五，會上要不要先對一下？
```

Rules:
- Roles come from memory notes; write `（不在記憶裡）` rather than guessing.
- "還開著的事" is built from `BodyPreview` and, when needed, one `outlook-thread` call on the newest mail per attendee; cite date and sender for every claim.
- Mark with ❗ anything the user still owes (use outlook-followup logic on the previews: a question to me with no later mail from me).
- List at most 5 attachments; prefer ones from the last two weeks and those named in the mails.
- Questions to raise must come from the mails (open items, unanswered asks, contradictions); no generic advice.
- Never offer to accept, decline or reschedule; never send anything.
