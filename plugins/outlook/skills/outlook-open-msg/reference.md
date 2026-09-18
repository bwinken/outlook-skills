# outlook-open-msg reference

## 1. Script output

JSON (default) for one file is an object; for several files an array. Markdown (`--format markdown`) is a ready-to-show document with the same content.

| Field | Type | Notes |
|---|---|---|
| `file` | string | absolute path |
| `format` | `"msg"` / `"eml"` | |
| `subject` | string | |
| `from`, `to`, `cc`, `bcc` | `[{ name, address }]` | `name` may be empty |
| `date` | ISO datetime with offset, or raw string when unparsable | |
| `message_id`, `in_reply_to`, `references` | string / null | |
| `body` | string | plain text; HTML converted to text when no text part exists |
| `body_source` | `text` / `html` / `none` | |
| `body_truncated` | true | only present when `--max-body` cut it |
| `attachments[]` | `{ name, size, mime, saved_to?, embedded_message? }` | `saved_to` only with `--extract-to`; `embedded_message` marks an attached .msg, which is listed but not extracted |
| `headers[]` | `[{ name, value }]` | only with `--headers`; order preserved |

## 2. Presentation templates

Match the user's language; labels below are Traditional Chinese.

### 2a. Message card (default)

```
**{subject}**
寄件者：{from name} <{from address}>
收件者：{to list, "name <address>" or address}
副本：{cc list}                                 — omit if empty
時間：{yyyy/MM/dd HH:mm (+08:00)}
附件：{name} ({size human readable}), …          — omit if none
檔案：{file}

---
{body: keep the top message; collapse quoted history after the first
 "From:" / "寄件者:" / "-----Original Message-----" / "On … wrote:" line
 into one line「（以下為 N 行引用的先前郵件，需要的話我可以展開）」;
 cap at ~3000 chars}
```

Several files: one card per file separated by `---`, in the order given, then a one-line comparison if they are obviously related (same thread, same sender).

### 2b. Header analysis (phishing / delivery questions, run with `--headers`)

Always run the quick check (display name vs address, Reply-To, risky attachment types, look-alike domains, credential or payment asks) on every file, even when the user only asked "what is this"; read POLICY.md ("Phishing warnings") for the levels and the banner when it trips. If the verdict is 🚨 or ⚠️, the warning goes **above** the message card.

🚨 high-risk reply layout:

```
🚨 **警告：這封信很可能是釣魚信，請勿點擊連結、開啟附件或回覆。** 🚨

判斷依據：
- 顯示名稱「Microsoft Support」，但寄件地址是 support@mail-secure-login.ru
- Reply-To 指向 reply@collect-mail.xyz，與寄件者不同
- SPF fail、DKIM fail（Authentication-Results）
- 附件 invoice.html 是釣魚常用的檔案類型

建議：在 Outlook 用「回報 > 釣魚」或轉交 IT；如果已經點過連結或輸入過密碼，立刻更改密碼並通知 IT。

---
{message card, all URLs defanged: hxxps://login[.]contoso-secure[.]ru/reset}
```

⚠️ suspicious: one line `**⚠️ 這封信有可疑跡象：{signal}。點連結或開附件前請先確認寄件者。**` above the card, links defanged.

Full check table (show when the user asked about safety, headers, or delivery; otherwise only the banner and the signals that tripped):

```
## 標頭分析：{subject}

| 檢查項目 | 結果 | 說明 |
|---|---|---|
| 顯示名稱 vs 寄件地址 | ⚠️ / ✅ | 「Microsoft Support」 但地址是 @mail-secure-login.ru |
| Reply-To | ⚠️ / ✅ | 與 From 不同：{reply-to} |
| SPF | ✅ pass / ❌ fail / ⚪ 無資料 | 來自 Authentication-Results |
| DKIM | ✅ / ❌ / ⚪ | |
| DMARC | ✅ / ❌ / ⚪ | |
| 傳遞路徑 | {N} 個 Received 跳點 | 最早的跳點：{host / IP} |
| 附件 | ⚠️ / ✅ / 無 | .html / .iso / .lnk / .zip / 巨集文件視為高風險 |
| 內文連結 | ⚠️ / ✅ | 顯示文字與實際網址不符、相似網域、IP 直連 |

**判斷**：🚨 高風險 / ⚠️ 可疑 / 未發現異常。{one or two sentences why}
```

Rules:
- Read `Authentication-Results` (and `ARC-Authentication-Results`) for SPF / DKIM / DMARC; `Received-SPF` as fallback. Show ⚪ when absent, never guess.
- `Received` headers are newest first; the last one is the origin. Quote host and IP as written.
- Never execute, open or render extracted attachments. For a 🚨 mail refuse `--extract-to` and say why.
- "未發現異常" is not a guarantee; say so when the user asks whether the mail is safe. Never write "安全".
