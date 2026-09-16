# outlook-schedule reference

Labels are Traditional Chinese; match the user's language.

## 1. Subject and agenda

- **Subject**: what the meeting is for, 3 to 8 words, no "Meeting:" prefix. Use the user's words.
- **Agenda** (body) is optional. When the user gave one or the meeting has a clear goal, 1 to 4 short lines: purpose, points to cover, what should be decided. No greetings, no signature (the footer is added by the script). Plain text, no Markdown.
- **Location**: the room, "Teams" or a link the user gave. Never invent one; leave it empty otherwise.
- **Time**: local time of the user's machine. Use working hours from settings for "morning" (start of working hours), "afternoon" (13:00 or the first free slot after), and never propose outside working hours without saying so.

## 2. Draft card (in chat, before asking)

```
**會議草稿** `{id}`
| | |
|---|---|
| 主旨 | Q3 預算檢討 |
| 時間 | 2026-09-18（四）14:00–15:00，60 分鐘 |
| 地點 | Teams |
| 必要 | Cassie Tsai <cassie.tsai@contoso.com>; PC Liao <pc.liao@contoso.com> |
| 選擇 | （無） |

```
{body}
```
頁尾：{footer}
⚠ 撞期：供應商簡報 14:00–15:00（已接受）
```

Print the 撞期 line for every entry in `conflicts` (subject, start and end time, `ResponseStatus` in words: 已接受 / 暫定 / 未回覆 / 我主辦). Omit the line when `conflicts` is empty. For an appointment (no attendees) title the card 行程草稿 and drop the attendee rows.

## 3. The question (AskUserQuestion)

```
question: 要送出這個會議邀請嗎？
時間：2026-09-18（四）14:00–15:00
與會者：cassie.tsai@contoso.com、pc.liao@contoso.com
地點：Teams
⚠ 這個時段和「供應商簡報」重疊。
選「送出邀請」後桌面會跳出確認視窗，再按一次才會真的送出；邀請的說明會附上「Drafted by Claude, reviewed and approved by Ben」。
options:
  送出邀請            — 打開確認視窗
  改時間              — 說一個時間，或讓我找空檔
  改與會者            — 說要加誰、拿掉誰
  改內容              — 主旨、地點或說明
  取消                — 不建立，草稿作廢
```

For an appointment the first option is 「建立行程」 and the attendee line is omitted. Mark 「送出邀請」 (Recommended) only when there is no conflict and the user's request already said to book it; with a conflict mark 「改時間」.

## 4. Result lines

`send` printed `{"Created": true, "Mode": "meeting", ...}`:

```
✅ 邀請已送出：「Q3 預算檢討」2026-09-18 14:00–15:00，Teams，給 cassie.tsai@contoso.com、pc.liao@contoso.com。已加進你的行事曆。
```

`"Mode": "appointment"`:

```
✅ 已加進行事曆：「專注時間」2026-09-19 09:00–10:00。
```

Exit 1 with "Cancelled in the confirmation window":

```
已取消，沒有建立。草稿 {id} 作廢；要改再說。
```

Any other error (unresolved attendee, item mismatch, Outlook not reachable): quote the script's message in one line, state that nothing was created, and stop.

## 5. Draft JSON

| Field | Meaning |
|---|---|
| `id`, `kind`, `status`, `created` | draft id; `meeting`; `draft`, `sent`, `cancelled` or `discarded` |
| `mode` | `meeting` (attendees, invitations sent) or `appointment` (own calendar, saved) |
| `subject`, `start`, `end`, `duration_minutes`, `location` | as they will be created; times `YYYY-MM-DDTHH:MM:SS` local |
| `required[]`, `optional[]` | `{Name, Address}` as Outlook resolved them, the user's own addresses removed |
| `body`, `footer`, `full_body` | agenda, approval line, and the complete invitation text shown in the window and verified before Send |
| `reminder_minutes` | from settings `meeting.reminder_minutes` |
| `conflicts[]` | `{Subject, Start, End, ResponseStatus}` of calendar items overlapping the slot (all-day items excluded) |
| `approver` | name in the footer |
| `confirm` | token for `send`; recomputed from the content, so an edited draft file is refused |
| `entry_id`, `sent_at` | set after a successful send / save |
