# outlook-availability reference

## 1. Script output (JSON)

Same as outlook-agenda (`outlook_calendar.py`): `Range`, `Calendar`, `Count`, `Conflicts`, `Items[]` with `Start`, `End`, `AllDayEvent`, `Subject`, `Location`, `Organizer`, `BusyStatus`, `MeetingStatus`, `ResponseStatus`, `IsRecurring`. See `../outlook-agenda/reference.md` for the full field table. Items with `BusyStatus = Free` are already excluded by the script.

## 2. Gap computation

Apply in this order, per day:

1. Working window: 09:00–18:00 local unless the user stated other hours; narrow it further for "早上" (09:00–12:00) or "下午" (13:00–18:00).
2. Busy intervals: every item's `[Start, End)` except
   - all-day items, unless `BusyStatus` is `OutOfOffice` (then the whole day is unavailable, say so and skip the day);
   - items with `MeetingStatus` Canceled / ReceivedAndCanceled.
   `Tentative` items count as busy; remember them for the note.
3. Merge overlapping or touching busy intervals.
4. Gaps = working window minus merged busy intervals.
5. Drop gaps shorter than the minimum: the length the user asked for, else 30 minutes.
6. Sort by day, then start time.

## 3. Presentation template

Match the user's language; labels below are Traditional Chinese.

```
## 空檔 {range label: 9/16（三） | 本週 9/14 – 9/20}（工作時間 09:00–18:00）

| 日期 | 空檔 | 長度 |
|---|---|---|
| 9/16（三） | 10:00–11:00 | 1 小時 |
| 9/16（三） | 12:00–14:00 | 2 小時 |
| 9/16（三） | 15:30–18:00 | 2.5 小時 |

{one line: 最長的整段是 12:00–14:00。 / 符合「一小時」的有 3 段，建議 15:30–16:30，前後都沒有會議。}
{if any tentative: 14:30–15:30 含 1 場暫定會議（1:1 with Bob），已當作忙碌計算。}

**當天行程**（only for one- or two-day ranges; for a week, show only the day the user asked about or skip）
| 時間 | 會議 | 狀態 |
|---|---|---|
| 09:30–10:00 | 每日站會 🔁 | ✅ |
| 11:00–12:00 | Q3 預算 VP review 預備 | ❓ 未回覆 |
```

Rules:
- Length: minutes below 60 (`45 分鐘`), otherwise hours with one decimal when not whole (`1.5 小時`).
- Multi-day range: one table for the whole range, rows grouped by day; days with no usable slot get one row `（沒有符合的空檔）`.
- When the user asked for a specific length, only list slots at least that long and recommend one; prefer slots not adjacent to another meeting.
- Nothing free in the range: say so, then show the closest alternatives (next day, or 08:30 / 18:00 edges) clearly marked as outside working hours.
- A `❓ 未回覆` meeting still blocks the slot; mention that declining it would free the time, but do not offer to decline it.
- Never offer to book, send an invite or hold the slot; the plugin is read-only.
