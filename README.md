# outlook-skills

A Claude Code plugin marketplace with **read-only** skills for a local Windows Outlook (Classic) mailbox.

## Install

Inside Claude Code:

```
/plugin marketplace add bwinken/outlook-skills
/plugin install outlook@outlook-skills
```

## Skills

| Skill | Ask things like |
|---|---|
| `outlook-status` | 我的 Outlook 資料檔在哪、信箱多大、有哪些帳號 |
| `outlook-search` | 找 Alice 上週寄給我的信、有附件的未讀郵件 |
| `outlook-thread` | 幫我摘要「Q3 預算」這串討論、結論和待辦是什麼 |
| `outlook-calendar` | 今天有什麼會議、這週哪裡有空、有沒有撞期 |
| `outlook-open-msg` | 打開這個 .msg / .eml、看標頭判斷是不是釣魚信 |

All skills only read. They never send, save, move, delete, flag or mark anything in Outlook. See [plugins/outlook/README.md](plugins/outlook/README.md) for requirements, the full read-only policy, and how the skills cope with a locked-down PowerShell execution policy.

## Layout

```
.claude-plugin/marketplace.json   marketplace manifest
plugins/outlook/                  the plugin (skills + scripts)
```

## Roadmap ideas

Not implemented yet, kept here as candidates: follow-up tracker (sent mails with no reply), inbox digest, meeting prep, attachment finder, mail-header phishing triage, rules audit (detect suspicious auto-forward rules), contacts lookup, .pst archive explorer.
