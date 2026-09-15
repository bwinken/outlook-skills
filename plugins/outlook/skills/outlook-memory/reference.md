# outlook-memory reference

## 1. File format

Path: `<root>/memory/<category>/<slug>.md`, slug derived from the title (lower-cased ASCII, CJK kept, punctuation to `-`). `memory.py` handles the front matter; do not hand-write it.

```
---
title: Alice Chen
category: people
tags: [legal, contoso, contracts]
created: 2026-09-14T10:02:11
updated: 2026-09-14T13:40:05
source: bootstrap
---
- 法務窗口，alice.chen@contoso.com
- 合約相關的信都由她發起（2026/09）
```

`source` is `bootstrap` for notes proposed from the mailbox scan, `user` for notes the user asked for. Body is bullets, one fact per line; add a date in parentheses when the fact can go stale.

## 2. Categories and what the overview JSON gives you

| Category | Title | Body bullets | From overview |
|---|---|---|---|
| `people` | the person's display name as Outlook shows it (`Alice Chen`), or `王小明` | address; role or team if inferable from signature-free signals (domain, folder, topics they appear in); what they usually write about; how often (`近半年 42 封`) | `TopSenders`, `TopRecipients` (merge the same address; a person who is both is a close correspondent) |
| `folders` | the folder path without the store (`Inbox/Vendors`) | what goes there, judged from top senders and topics inside it; item count and newest date | `Folders` plus `TopSenders[].Folder` |
| `projects` | the conversation topic cleaned of Re:/FW: (`Q3 預算`) or a name the user uses | keywords to search with (`-AnyOf` terms), who drives it, last activity date | `TopTopics` (skip one-off or automated subjects) |
| `recurring` | the meeting subject (`每日站會`) or a routine (`週報`) | cadence in words (`每週一到五 09:30，30 分鐘`), organizer, attendees | `RecurringMeetings` (`Pattern`, `Interval`, `DayOfWeekMask`, `StartTime`) |
| `preferences` | one file per preference area (`回覆格式`, `工作時間`) | the preference itself | never from the scan; only from the user |

Newsletters (`Newsletters` list, senders with a List-Unsubscribe header) are not people. Offer one `folders`-style note titled `電子報與通知` listing them only if the user wants it.

`DayOfWeekMask` bits: 1 Sun, 2 Mon, 4 Tue, 8 Wed, 16 Thu, 32 Fri, 64 Sat. `Pattern` Weekly with mask 62 and Interval 1 = 每週一到五.

## 3. Title rules

- A title is the name someone would type to find the note again: the person's name, the folder path, the project's short name, the meeting subject. Not a sentence, not a date.
- One topic per file. `Alice Chen` and `Bob Lin` are two files, not `法務團隊`. A project gets one file even if it spans many threads.
- Before creating, run `memory.py find` with the person's surname, the address, or the project keyword; a match means `append`, not `new`.
- Keep titles stable. If a project is renamed, add the new name as a tag and a bullet rather than renaming the file.
- Tags: lower-case, short, reusable across files (`contoso`, `legal`, `vendor`, `budget`, `2026`). Three to five per file.

## 4. Presenting

Onboarding draft (before writing anything):

```
從最近 180 天的 1,842 封郵件與行事曆整理出以下記憶草稿（只存摘要，不存內文）：

| 分類 | 標題 | tags | 內容 |
|---|---|---|---|
| people | Alice Chen | legal, contoso | 法務窗口，alice.chen@contoso.com，近半年 42 封 |
| folders | Inbox/Vendors | vendor | 供應商往來，312 封，最新 09/12 |
| projects | Q3 預算 | budget, 2026 | 關鍵字：預算、方案 B、VP review；王小明主導 |
| recurring | 每日站會 | standup | 每週一到五 09:30，30 分鐘，Alice 主辦 |

共 {N} 條：people {a}、folders {b}、projects {c}、recurring {d}。
```

Then the question. After writing: `已寫入 {N} 條到 ~/.outlook-skills/memory/（people 12、folders 6、projects 8、recurring 3）`.

"What do you remember" (`memory.py list`):

```
## 我記得的（{N} 條）

**人物**（12）：Alice Chen、王小明、Bob Lin、…
**資料夾**（6）：Inbox/Vendors、Inbox/Projects/Alpha、…
**專案**（8）：Q3 預算、合約 v3、…
**定期事務**（3）：每日站會、週報、月結
**偏好**（2）：回覆格式、工作時間

要看哪一條的內容？
```

Show the body of a single note with `memory.py show` when asked.

## 5. Settings

| Key | Default | Used by | Meaning |
|---|---|---|---|
| `language` | `zh-TW` | all | language of Claude's replies and table labels |
| `working_hours.start` / `.end` | `09:00` / `18:00` | availability | working window for free-slot search |
| `working_hours.days` | `[1,2,3,4,5]` | availability | ISO weekdays counted as working days (1 = Monday) |
| `availability.min_slot_minutes` | `30` | availability | gaps shorter than this are dropped |
| `search.default_lookback_days` | `90` | search | `-After` when the user gives no date hint |
| `search.default_folder` | `Inbox` | search, thread | `-Folder` when none is given |
| `search.all_folders` | `false` | search | pass `-AllFolders` by default |
| `search.max_candidates` | `300` | search | `-Max` for reranker candidate fetches |
| `search.direct_read_max` | `20` | search | up to this many hits are read directly instead of reranked |
| `store` | `null` | all COM skills | `-Store` for a shared or archive mailbox |
| `rerank.gateway` / `.model` / `.api_key` | `null` | search (rerank.py) | override the gateway resolution; beats `~/.claude/settings.json`, loses to env vars and flags |
| `rerank.auto_consent` | `false` | search | skip the per-run confirmation before sending candidates to the reranker |
| `status.skip_com` | `false` | status | always run `outlook-status` with `-SkipCom` (New Outlook machines) |

`settings.py show` returns `first_run` (true when `~/.outlook-skills` does not exist yet), `settings` (merged), `sources` (key → file for every non-default key), `layers` (files read), `memory` (index: count, per_category, entries with title/category/tags/updated/path), `memory_hint`, `local_dir`, `user_dir`.

Presenting `show`:

```
