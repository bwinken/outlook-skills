# outlook-agenda reference

## 1. Script output (JSON)

```
{
  "Range":     { Start, End, Filter },      // End is exclusive
  "Calendar":  <folder path>,
  "Count":     <int>,
  "Conflicts": [ { A, AStart, AEnd, B, BStart, BEnd }, ... ],
  "Items":     [ <appointment>, ... ]       // sorted by Start
}
```

Appointment:

| Field | Values / notes |
|---|---|
| `Start`, `End` | ISO local datetime; all-day items start at 00:00 |
| `DurationMinutes` | int |
| `AllDayEvent` | bool |
| `Subject`, `Location` | string |
| `Organizer` | display name |
| `RequiredAttendees`, `OptionalAttendees` | `;` separated display names |
| `BusyStatus` | Free / Tentative / Busy / OutOfOffice / WorkingElsewhere |
| `MeetingStatus` | NonMeeting (personal appointment) / Meeting (I organise) / Received / Canceled / ReceivedAndCanceled |
| `ResponseStatus` | None / Organized / Tentative / Accepted / Declined / NotResponded |
| `IsRecurring` | bool; the item is already the expanded occurrence |
| `Categories` | string |
| `BodyPreview` | first 300 chars |

Items with `BusyStatus = Free` are excluded unless `-IncludeFree` was passed.

## 2. Presentation template

Match the user's language; labels below are Traditional Chinese.

```
## {range label: 今天 9/14（一） | 本週 9/14 – 9/20 | 9/15 – 9/19}

**9/14（一）**
| 時間 | 會議 | 地點 | 主辦 | 狀態 |
|---|---|---|---|---|
| 全天 | 出差台中 | | | |
| 09:30–10:00 | 每日站會 🔁 | Teams | Alice | ✅ |
| 14:00–15:30 | Q3 預算審查 | 3F 會議室 | 王小明 | ❓ 未回覆 |

**9/15（二）**
（沒有行程）

**衝突**（only if Conflicts non-empty）
- 9/16 14:00–15:00「供應商簡報」與 14:30–15:30「1:1 with Bob」重疊 30 分鐘

**需要處理**（only if any）
- {N} 場會議尚未回覆：{subjects}
- {N} 場已取消但仍在行事曆上：{subjects}
```

Rules:
- Time column `HH:mm–HH:mm`; all-day items show `全天` and go first in the day.
- Status icons: ✅ Accepted / Organized, ❓ NotResponded, 🕒 Tentative, ❌ Canceled, blank for personal appointments. Add 🔁 after the subject when `IsRecurring`.
- Day header uses `M/d（週）` with the Chinese weekday character (一二三四五六日).
- Skip empty days inside a one-day range; list them as `（沒有行程）` inside a multi-day range so the user sees gaps.
- Location: shorten Teams / Zoom / Meet URLs to the platform name.
- More than 40 items: group per day with counts and list only meetings that need action, then offer the full list.
- The plugin is read-only: never offer to accept, decline or reschedule.
