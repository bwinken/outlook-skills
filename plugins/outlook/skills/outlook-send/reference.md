# outlook-send reference

Labels are Traditional Chinese; match the user's language and the profile's language.

## 1. Writing the body

Write like a person answering mail between meetings, not like a report.

- **Length**: the profile's `Length.MedianChars` is the target; without a profile, 2 to 6 short lines. Never more than about 120 Chinese characters or 80 English words unless the user dictated more.
- **Order**: the answer or decision first, then at most one reason, then the one thing the reader should do (with a date if there is one). Nothing else.
- **Greeting and closing**: the first of `Greetings` and `Closings` in the profile, then the profile `Signature` block if the user has one. Without a profile: 「{name}，」 / 「謝謝」, or "Hi {name}," / "Thanks," in English.
- **Language**: the profile's `Language` (MostlyChinese), else settings `language`, else the original mail's language.
- **Tone**: the profile's; otherwise plain and polite. No exclamation marks, no emoji, no "I hope this email finds you well".
- **Replies**: do not repeat what the original said; answer it. Every question in the original gets one line. Never invent facts, dates or commitments the user did not give; put a `[?]` placeholder and ask instead.
- **Plain text only**: no Markdown, no headings, no tables. Line breaks between points.

Example (profile says short, Chinese, closes with 「謝謝」):

```
PC 你好，

名單看過了，沒問題，兩位候選人都可以進下一輪。
週五前我會把面試時間回給你。

謝謝
Ben
```

## 2. Draft card (in chat, before asking)

Show what `draft` returned, verbatim. The body block is copied from `body`, not re-typed.

```
**草稿** `{id}`
| | |
|---|---|
| 收件者 | Cassie Tsai <cassie.tsai@contoso.com>; PC Liao <pc.liao@contoso.com> |
| 副本 | （無） |
| 主旨 | RE: 合約草稿 v3 - 法務意見 |

```
{body}
```
頁尾：{footer}
附件：報價.pdf（1.2 MB，C:\Users\me\Desktop\報價.pdf）
（回信：原信會引用在頁尾下方，{quote 行數} 行）
```

Omit the 附件 line when there are none. Say in one line where the greeting, closing and length came from (profile or default) so the user can correct the style once.

## 3. The question (AskUserQuestion)

```
question: 要寄出這封信嗎？
收件者：cassie.tsai@contoso.com、pc.liao@contoso.com
副本：（無）
主旨：RE: 合約草稿 v3 - 法務意見
附件：報價.pdf（1.2 MB）
內容如上；寄出的信會和草稿完全一致，並附上「Drafted by Claude, reviewed and approved by Ben」的頁尾。
選「寄出」後桌面會跳出確認視窗，再按一次「寄出」才會真的送出。
options:
  寄出                — 打開確認視窗
  修改內容            — 說一下要改什麼
  改收件者 / 副本     — 說一下要加誰、拿掉誰
  取消                — 不寄，草稿作廢
```

「寄出」 is marked (Recommended) only when the user's request already said to send ("回他說好", "寄給 Alice"); for "幫我擬一封" leave the recommendation off.

## 4. Result lines

`send` printed `{"Sent": true, ...}`:

```
✅ 已寄出（2026-09-16 14:05）給 cassie.tsai@contoso.com、pc.liao@contoso.com，主旨「RE: 合約草稿 v3 - 法務意見」，附件 報價.pdf。副本在寄件備份。
```

`InheritedAttachmentsRemoved` lists the original's inline pictures that Outlook had copied into the reply item and that `send` removed before the check; they were never part of the draft, so mention them only if the user asks why a picture from the original is not in the reply.

Exit 1 with "Cancelled in the confirmation window":

```
已取消，沒有寄出。草稿 {id} 作廢；要改再說。
```

Any other error (unresolved recipient, item mismatch, Outlook not reachable): quote the script's message in one line, state that nothing was sent, and stop.

## 5. Draft JSON

| Field | Meaning |
|---|---|
| `id`, `status`, `created` | draft id; `draft`, `sent`, `cancelled` or `discarded` |
| `mode` | `new`, `reply`, `reply_all` |
| `reply_to` | `{EntryID, From, Subject, ReceivedTime, ReplyAll}` for replies |
| `to[]`, `cc[]` | `{Name, Address}` as Outlook resolved them; exactly what will be sent |
| `subject` | as sent; replies get `RE:` unless the original already had it |
| `body` | your text as given |
| `footer` | approval line from settings `send.footer` with the approver's name |
| `quote` | quoted original (replies, when `send.quote_original` is true) |
| `full_body` | `body` + `footer` + `quote`: the complete outgoing text, shown in the window and verified before Send |
| `attachments[]` | `{Path, Name, Size, Sha256}`; re-hashed by `send` before attaching, names verified on the item (a reply item's inherited inline pictures are removed before that check) |
| `approver` | name in the footer |
| `confirm` | token for `send`; recomputed from the content, so an edited draft file is refused |
