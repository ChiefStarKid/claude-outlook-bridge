# outlook-bridge

A Python script that gives Claude Code direct access to your Outlook Desktop App via Windows COM automation — no OAuth, no API keys, no cloud intermediary.

## Why this exists

Most Claude integrations for email require OAuth setup, API credentials, or a cloud intermediary. This bridge takes a different approach: if Outlook is already open on your machine and you're signed in, you already have access — no extra authentication needed.

The bridge started as a simple read-only tool so Claude could look up emails during a conversation without context-switching. Over time it grew to cover the full email lifecycle: composing and sending, reply and reply-all, forwarding with an optional note, moving emails between folders, and querying the calendar. The pattern that emerged — a thin Python script that Claude calls via Bash, returning structured JSON — turned out to be more reliable and simpler to maintain than any webhook or API approach.

If you use Outlook on Windows and work with Claude Code, this lets Claude become a genuine email assistant: reading threads for context, drafting replies, chasing actions, and keeping your calendar in view — all without leaving the conversation.

---

## Prerequisites

- **Windows** (COM automation is Windows-only)
- **Python 3.9+**
- **Outlook Desktop App** — the classic Windows thick client, signed in and running. This does _not_ work with the new Outlook web wrapper, OWA, or Mac Outlook.
- **pywin32**: `pip install pywin32`

---

## Installation

```bash
git clone https://github.com/your-org/outlook-bridge.git
cd outlook-bridge
pip install -r requirements.txt
```

Place `outlook_bridge.py` wherever you'd like to call it from — typically `~/.claude/` if you're integrating with Claude Code.

---

## Claude Code permissions

Claude invokes the bridge via a single Bash call: `python /path/to/outlook_bridge.py [command]`. You need to grant Claude Code permission to run that command.

**Scoped (recommended)** — add to `.claude/settings.json` in your project:

```json
{
  "permissions": {
    "allow": ["Bash(python *outlook_bridge.py*:*)"]
  }
}
```

This allows Claude to run the bridge script only — no other Bash commands are auto-approved.

**Time-bound** — Claude Code doesn't have a native session-expiry permission model. Use project-level settings (`.claude/settings.json` in a specific project folder) rather than global settings. The permission only applies in that project's context; remove or comment out the entry when you no longer want Claude to have email access.

**Broad** — allow all Bash in your Claude Code settings. Not recommended unless you already trust Claude broadly.

> ⚠️ **Auto-send warning**: `send`, `reply`, and `forward` all send immediately by default. Always instruct Claude to use `--draft` unless you explicitly want auto-send. See [CLAUDE_INTEGRATION.md](CLAUDE_INTEGRATION.md) for recommended CLAUDE.md setup.

---

## Commands

All output is JSON. Add `--pretty` anywhere in the command for human-readable formatting.

### `folders`
List all mail folders.
```
python outlook_bridge.py folders --pretty
```

### `list`
List recent emails in a folder.
```
python outlook_bridge.py list [--folder INBOX] [--count 20]
```

| Flag | Default | Description |
|------|---------|-------------|
| `--folder` | `inbox` | Folder name (`inbox`, `sent`, `drafts`, `deleted`) or path like `Projects/Client` |
| `--count` | `20` | Number of emails to return |

### `search`
Search emails by keyword (subject, body, sender, recipient).
```
python outlook_bridge.py search "renewal" [--folder INBOX] [--count 20]
```

### `read`
Read a full email including body and attachments. Requires an `EntryID` from `list` or `search`.
```
python outlook_bridge.py read <entry_id>
```

### `send`
Send a new email. **Sends immediately.**
```
python outlook_bridge.py send --to user@example.com --subject "Hello" --body "Message here"
python outlook_bridge.py send --to a@x.com b@x.com --cc c@x.com --subject "Hi" --body "..." --html
```

| Flag | Required | Description |
|------|----------|-------------|
| `--to` | Yes | One or more recipient addresses |
| `--subject` | Yes | Email subject |
| `--body` | Yes | Email body (plain text, or HTML if `--html`) |
| `--cc` | No | One or more CC addresses |
| `--html` | No | Treat `--body` as HTML |

### `reply`
Reply to an email. Use `--draft` to save to Drafts instead of sending.
```
python outlook_bridge.py reply <entry_id> --body "Thanks, noted." --draft
python outlook_bridge.py reply <entry_id> --body "Agreed." --all
```

| Flag | Required | Description |
|------|----------|-------------|
| `--body` | Yes | Reply text (prepended inline above quoted thread) |
| `--draft` | No | Save to Drafts instead of sending |
| `--all` | No | Reply all |
| `--html` | No | Treat `--body` as HTML |
| `--attach-email` | No | Entry IDs of emails to attach as `.msg` files |

### `forward`
Forward an email. Use `--draft` to save to Drafts instead of sending.
```
python outlook_bridge.py forward <entry_id> --to user@example.com --body "FYI" --draft
```

| Flag | Required | Description |
|------|----------|-------------|
| `--to` | Yes | One or more recipient addresses |
| `--body` | No | Optional note prepended above the forwarded content |
| `--draft` | No | Save to Drafts instead of sending |
| `--cc` | No | One or more CC addresses |

### `move`
Move an email to a folder.
```
python outlook_bridge.py move <entry_id> --folder "Archive"
python outlook_bridge.py move <entry_id> --folder "Projects/Client"
```

### `delete`
Move an email to Deleted Items.
```
python outlook_bridge.py delete <entry_id>
```

### `cal-list`
List calendar events in a date range.
```
python outlook_bridge.py cal-list --from 2025-01-01 --to 2025-01-31
python outlook_bridge.py cal-list --from 2025-01-01 --to 2025-01-31 --cal work
```

| Flag | Required | Description |
|------|----------|-------------|
| `--from` | Yes | Start date `YYYY-MM-DD` |
| `--to` | Yes | End date `YYYY-MM-DD` |
| `--cal` | No | `personal` (default) or a shared calendar name from your config |

---

## Shared calendar config

To query a shared calendar, create `outlook_bridge_config.json` in the same directory as the script (or at `~/.claude/outlook_bridge_config.json`):

```json
{
  "shared_calendars": {
    "work": "shared-calendar@yourorg.com"
  }
}
```

The key (`work`) becomes the value you pass to `--cal`. You can define as many shared calendars as you have access to in Outlook. If no config file exists, only `--cal personal` is available.

---

## Output format

All commands return a JSON object. On error, the script exits with code 1 and prints `{"error": "..."}`.

Use `--pretty` for human-readable output:
```
python outlook_bridge.py list --pretty
```

---

## Design notes

**COM automation vs Outlook Add-ins**

The Outlook Desktop App exposes two ways to build on top of it. *Add-ins* use the Office JS API — they run inside Outlook as a web-based extension and work across desktop, web, and Mac. *COM automation* (what this bridge uses) is a lower-level Windows interface that lets external programs talk directly to the Outlook process. COM gives broader access — full folder trees, arbitrary mail operations, shared calendars, the ability to send without a UI — but only works on Windows with the thick client running.

This bridge chose COM because the goal is automation from _outside_ Outlook, not extending the Outlook UI. An Add-in lives inside Outlook; this bridge lets Claude (an external agent) drive Outlook as a tool.

**How Python calls COM**

The bridge uses `pywin32`, a Python library that wraps the Windows COM interface. `win32com.client.Dispatch("Outlook.Application")` connects directly to the running Outlook process — in-process, no network, no API call. Everything from there is Python method calls on COM objects. Bash is only involved once: Claude executing `python outlook_bridge.py [command]`. Once Python is running, everything else is pure Python ↔ COM ↔ Outlook.

**Why not Microsoft Graph API?**

Graph works cross-platform and doesn't require Outlook to be open. The trade-off: it requires an Azure app registration, OAuth consent flow, and ongoing token management. COM requires nothing beyond Outlook being signed in. For a local AI assistant workflow where you're already at your Windows machine, COM is the simpler, faster path.

**Reply and forward — inline drafts**

Replies and forwards are constructed the way Outlook constructs them natively: the new content is prepended above the quoted thread HTML, so the draft appears inline in Outlook exactly as a user would compose it. The thread is preserved, formatting is intact, and you can review the full context before deciding whether to send.

**Why JSON output?**

Claude processes structured data better than prose. Every command returns a JSON object with consistent field names so Claude can reason about email content, extract IDs, chain operations (read → reply, search → forward), and summarise threads — without screen-scraping or parsing unstructured text.

---

## Limitations

- Windows only (requires COM / pywin32)
- Requires the Outlook Desktop App (the Windows thick client) — not compatible with the new Outlook web wrapper, OWA, or Mac Outlook
- Outlook must be open and signed in when the script runs
- Tested against classic Outlook (Microsoft 365 subscription); behaviour may vary on older perpetual-license versions
