# Policy for every outlook skill

Read by the skills at run time. Two rules: never write to Outlook except through `outlook-send`, whose Send happens only after the user clicked Send in a window on their desktop; and warn before showing a mail that looks like phishing.

## Read-only policy

Apart from `scripts/outlook_send.py send` (see below), no script or skill may:

- call `Save`, `Send`, `Delete`, `Move`, `Copy`, `Forward`, `Reply`, `ReplyAll`, `Respond`, `Display`;
- set any property (`UnRead`, `Categories`, `FlagStatus`, `Importance`, `BusyStatus`, ...);
- create items, folders, rules, or appointments;
- compact, repair, detach or attach data files;
- write anywhere except a user-specified `-OutFile` / `--extract-to` path and the plugin's own `.outlook-skills/` folders.

Reading through COM does not change read/unread state. The shared module `scripts/outlook_com.py` exposes only getters; add new skills on top of it and keep the same rule.

**The one exception, sending mail.** `outlook_send.py` creates and sends a mail in two steps: `draft` stores the exact recipients, subject and text and sends nothing; `send` opens a confirmation window on the user's own desktop showing all of it, and calls Send only after the user clicks Send there. The item is compared with the stored draft first and not sent on any difference. No flag, setting or environment variable skips the window. The `outlook-send` skill additionally gets the user's yes in chat, on the recipients and the text, before it runs `send`. No other skill sends, saves to Drafts, replies, forwards or accepts anything.

## Phishing warnings

Whenever a skill shows the content of a single mail (`outlook-open-msg`, `outlook-search` with `-IncludeBody`, `outlook-thread`), Claude first runs a quick phishing check and, if it trips, warns **before anything else** in the reply. The warning must be impossible to miss: it comes first, uses the 🚨 banner below, names the concrete signals, and tells the user what not to do. Soft wording ("might be worth checking") is not acceptable for a high-risk verdict.

Levels:

| Level | When | Reply |
|---|---|---|
| 🚨 高風險 | two or more strong signals, or one strong signal plus a request for credentials, payment or urgent action | banner first, then the message card with all links defanged; do not summarise the mail's call to action as if it were legitimate |
| ⚠️ 可疑 | one strong signal, or several weak ones | one bold warning line above the card, links defanged |
| none | nothing tripped | normal card; say "看起來正常" only when asked, never "安全" |

Strong signals: display name and address disagree (「Microsoft Support」 from a non-Microsoft domain); Reply-To differs from From; SPF, DKIM or DMARC `fail` in `Authentication-Results`; a look-alike domain (`micros0ft`, `contoso-secure.ru`); an attachment of a risky type (.html, .htm, .iso, .img, .lnk, .js, .vbs, .zip with such content, macro-enabled Office files); a link whose visible text and target differ; a login or password-reset link to a domain that is not the claimed sender's.
Weak signals: urgency or threats (帳號將被停用、24 小時內), generic greeting, unexpected invoice or delivery notice, external sender writing about internal matters, first-time sender asking for a bank change.

Banner (Traditional Chinese by default; translate for an English user):

```
🚨 **警告：這封信很可能是釣魚信，請勿點擊連結、開啟附件或回覆。** 🚨

判斷依據：
- 顯示名稱「Microsoft Support」，但寄件地址是 support@mail-secure-login.ru
- SPF 與 DKIM 皆為 fail
- 附件 invoice.html 是釣魚常用的檔案類型

建議：在 Outlook 用「回報 > 釣魚」或轉交 IT；如果已經點過連結或輸入過密碼，立刻更改密碼並通知 IT。
```

Rules for both warning levels:
- Defang every URL from the suspicious mail: `hxxps://login[.]contoso-secure[.]ru/reset`. Never leave a clickable link.
- Do not extract or open attachments from a 🚨 mail even if the user asks; explain why once and offer the header analysis instead.
- In lists (search results, thread timelines) mark such messages with 🚨 in the subject cell and say so above the table.
- The verdict is a judgement, not a scan result: say what was checked and what was not (for example, no header data in COM output, so SPF/DKIM unknown).
