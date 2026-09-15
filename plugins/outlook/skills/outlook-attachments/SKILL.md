---
name: outlook-attachments
description: Find attachments in the local Outlook (Classic) mailbox READ-ONLY - by file name, extension, size, sender, date or folder; list the biggest ones; and, with the user's consent, copy selected files out to a folder. Use when the user asks "find the contract PDF Alice sent", "biggest attachments in my mailbox", "all xlsx files from last month", "save that attachment to my desktop", or 找附件 / 最大的附件 / 上個月的 Excel 附件 / 把附件存出來.
---

# outlook-attachments

Read-only. Copying a file out uses Outlook's own SaveAsFile, which reads the attachment and writes only to the folder the user chose; the mailbox is untouched.

## Where this runs

Needs Windows with Classic Outlook and a host that executes commands on that same machine (Claude Code, or Claude Code inside Claude Desktop). In a Claude Desktop chat skill or Cowork the sandbox cannot reach Outlook: say so in one line and point the user to Claude Code.

## Run

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_attachments.py" -Name 合約 -After 2026-06-01
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_attachments.py" -Ext pptx,xlsx -MinSizeKB 500 -Sort size -Top 20 -AllStores -AllFolders
python "${CLAUDE_PLUGIN_ROOT}/scripts/outlook_attachments.py" -From alice -Ext pdf -SaveTo "C:/Users/<me>/Desktop/from-alice"
```

Takes every `outlook_search.py` option (`-From`, `-Subject`, `-After`, `-Folder`, `-Store`, `-AllStores`, `-AllFolders`, ...) plus: `-Name` file name contains; `-Ext` extensions; `-MinSizeKB`; `-Sort date|size|name`; `-Top N`; `-IncludeEmbedded` (inline images, OLE); `-SaveTo <folder>`.

## Workflow

1. Turn the request into filters. "Biggest" = `-Sort size -AllFolders` (add `-AllStores` for PST mail); "from X" = `-From`; "last month" = `-After`/`-Before`.
2. Run without `-SaveTo` first and show the list (reference.md). Inline images are excluded unless the user wants them.
3. **Copying out needs consent**: only pass `-SaveTo` after the user named or agreed to a folder in this conversation. Ask with the host's question tool when the request was vague ("save it somewhere"). Report the exact paths written.
4. Never open, run or render extracted files. For a mail flagged 🚨 by the phishing check, refuse to extract and say why.

## Settings and memory

Before the first Outlook call in a conversation, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/settings.py" show` once (no Outlook access, instant). `first_run: true` means `~/.outlook-skills` does not exist yet: hand over to `outlook-memory`'s onboarding first (respect a "not now"). Apply `store` (pass it as `-Store`), `language`, and any search defaults. When the request names a person, folder or project, run `python "${CLAUDE_PLUGIN_ROOT}/scripts/memory.py" find "<word>"` and `show` the note so aliases resolve. Never write memory silently.

## Output format

Read `${CLAUDE_PLUGIN_ROOT}/skills/outlook-attachments/reference.md` for the JSON fields and the presentation template.

## Read-only rules

Follow the plugin's read-only policy in `${CLAUDE_PLUGIN_ROOT}/README.md`. Never reply, forward, flag, mark read, move or delete. Drafts, if asked for, are written in chat only.

