# outlook (plugin)

Skills for a local Windows Classic Outlook mailbox. Skill list and settings: [repo README](../../README.md).
Rules the skills follow at run time (read-only except `outlook-send` and `outlook-schedule`, which send only after the user's click in a desktop window; phishing warnings): [POLICY.md](POLICY.md).

## Requirements

- Windows with Classic Outlook (2016 to Microsoft 365). New Outlook has no COM object model.
- Python 3.8+ and `pip install pywin32`. Everything else is the standard library.
- Outlook may be open or closed; a closed Outlook is started in the background.

## Scripts

Every skill runs plain Python from `scripts/`, output is UTF-8 JSON on stdout (`-OutFile x.json` to a file). Options take `-From` or `--from`.

```
python scripts/outlook_search.py -From alice -After 2026-09-01 -Max 20
python scripts/outlook_thread.py -Subject "Q3 budget"
python scripts/outlook_calendar.py -Days 7
python scripts/outlook_followup.py -Direction both
python scripts/outlook_meeting_prep.py -Next
python scripts/outlook_send.py draft -ReplyTo <EntryID> -BodyFile body.txt   # then: send <id> -Confirm <token>
python scripts/outlook_meeting.py draft -Subject "Q3 review" -Start 2026-09-18T14:00 -Attendees alice@contoso.com   # then: send <id> -Confirm <token>
python scripts/outlook_attachments.py -Ext pdf -Sort size -Top 10
python scripts/outlook_overview.py -Days 180
python scripts/outlook_status.py -SkipCom
python scripts/read_msg.py message.msg --headers
python scripts/settings.py show
python scripts/memory.py find alice
```

| Script | Used by |
|---|---|
| `outlook_com.py` | shared read-only COM helpers: connect, folders, dates, item summaries |
| `outlook_search.py`, `outlook_attachments.py`, `rerank.py` | outlook-search |
| `outlook_thread.py` | outlook-thread |
| `outlook_calendar.py` | outlook-agenda, outlook-availability |
| `outlook_followup.py` | outlook-morning-brief |
| `outlook_meeting_prep.py` | outlook-meeting-prep |
| `outlook_send.py` | outlook-send: `draft` stores the exact outgoing text, `send` opens the confirmation window and sends after the click |
| `outlook_meeting.py` | outlook-schedule: same two steps for a meeting or appointment, with an overlap check in the draft |
| `outlook_overview.py`, `outlook_style.py` | outlook-setup (mailbox scan, reply-habit profile) |
| `read_msg.py`, `msgfile.py` | outlook-open-msg (no Outlook needed, any OS) |
| `settings.py`, `memory.py` | every skill, outlook-memory |

## Other hosts

- **Zoo Code / Agent Skills hosts**: `python install.py` copies the skill files into `~/.roo/skills/` (`--agents` for `~/.agents/skills/`, `--project` for the current directory) and points them at this folder. `--uninstall` removes them.
- **Claude Desktop / claude.ai upload**: `python install.py --zip` builds one zip per skill in `dist/`.
- **MCP** instead of skills: [plugins/outlook-mcp](../outlook-mcp/README.md).

## Reranker (optional)

Fuzzy search sends subject, sender, date and a preview of each candidate to an OpenAI-compatible reranker (vLLM `/v1/rerank` or `/v1/score`). Use an internal gateway; the skill asks before sending unless `rerank.auto_consent` is true.

| Setting | First match wins |
|---|---|
| gateway | `--gateway`, `OUTLOOK_RERANK_URL`, `rerank.gateway` in settings.json, `ANTHROPIC_BASE_URL`, `OPENAI_BASE_URL` |
| model | `--model`, `OUTLOOK_RERANK_MODEL`, `rerank.model`, default `bge-reranker-v2-m3` |
| api key | `--api-key`, `OUTLOOK_RERANK_API_KEY`, `rerank.api_key`, `ANTHROPIC_AUTH_TOKEN`, `ANTHROPIC_API_KEY`, `OPENAI_API_KEY` |

Environment names are read from the process, then from the `env` block of `~/.claude/settings.json`. Note the fallback: with nothing configured, the skills version reuses Claude Code's `ANTHROPIC_BASE_URL` gateway. `python scripts/rerank.py --show-config` shows what would be used and whether the endpoint answers.

## Layout

```
.claude-plugin/plugin.json
POLICY.md            read-only policy, the send exception and phishing rules, read by the skills
install.py           copy skills to other hosts, or build zips
scripts/             the Python scripts above
skills/<name>/       SKILL.md (when and how) + reference.md (JSON fields, reply template)
tests/               fake Outlook object model + unittest suites
```

Tests: `python -m unittest discover -s tests`. Skills reference scripts as `${CLAUDE_PLUGIN_ROOT}/scripts/...`, which Claude Code resolves to the installed plugin directory.
