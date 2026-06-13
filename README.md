# claude-outlook-bridge

**Let Claude Code (and other LLM agents) read and send Outlook email on Windows — no OAuth, no Microsoft Graph, no cloud intermediary.**

A single-file Python bridge that connects Claude Code to your Outlook Desktop App via Windows COM automation. If you're signed into Outlook, Claude has access. Read inboxes, search threads, draft replies, forward with notes, move messages, query calendars — all from inside a Claude Code conversation.

Built for: Claude Code on Windows, Cursor, Aider, or any agent that can shell out to `python`. A practical alternative to building an MCP server when Outlook is already open in front of you.

---

## At a glance

| | |
|---|---|
| **Platform** | Windows 10/11 only (uses COM / `pywin32`) |
| **Python** | 3.9+ |
| **Auth model** | None — inherits your signed-in Outlook session |
| **Outlook flavour** | Classic Outlook Desktop App (the Windows thick client). Not New Outlook, OWA, or Mac. |
| **Commands** | `folders`, `list`, `search`, `read`, `send`, `reply`, `forward`, `move`, `delete`, `cal-list` |
| **Output** | JSON (add `--pretty` for human formatting) |
| **Dependencies** | `pywin32` |
| **License** | MIT |
| **Agent config** | See [AGENTS.md](AGENTS.md) and [llms.txt](llms.txt) |

---

## Quickstart

```bash
git clone https://github.com/ChiefStarKid/claude-outlook-bridge.git
cd claude-outlook-bridge
pip install -r requirements.txt
python outlook_bridge.py list --pretty   # smoke test
```

Then drop this into `.claude/settings.json` in your project to let Claude Code invoke it:

```json
{ "permissions": { "allow": ["Bash(python *outlook_bridge.py*:*)"] } }
```

Coding agents should read [AGENTS.md](AGENTS.md) for the full auto-wire snippet. Humans, keep reading.

---

## When to use this vs. alternatives

| Approach | Auth setup | Platform | Best for |
|---|---|---|---|
| **claude-outlook-bridge (this repo)** | None — uses signed-in Outlook | Windows + classic Outlook | Local AI assistant on your own machine. Zero config beyond a pip install. |
| Microsoft Graph API | Azure app registration + OAuth | Any | Server-side automation, cross-platform, headless. |
| Outlook Add-in (Office JS) | Side-load or Store deploy | Anywhere Outlook runs | Extending the Outlook UI itself, not driving it from outside. |
| Generic email MCP server | Varies (often IMAP/OAuth) | Any | Multi-mailbox / multi-provider setups, or non-Outlook accounts. |

If Outlook is already open on your Windows box and you just want Claude to use it, this bridge is the shortest path. For anything cross-platform or headless, use Graph.

---

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
git clone https://github.com/ChiefStarKid/claude-outlook-bridge.git
cd claude-outlook-bridge
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

## FAQ

### Does this work with the new Outlook for Windows?
No. The new Outlook is a web wrapper without the COM surface. You need the classic Outlook Desktop App (the one bundled with Microsoft 365 / Office 2019+).

### Does this work on macOS or Linux?
No. COM is Windows-only. On macOS or for headless setups, use the Microsoft Graph API instead.

### Why not just use an MCP server?
MCP is great when you need a long-running server process and cross-client compatibility. This bridge is deliberately the opposite: one Python file, no daemon, no protocol overhead. Claude shells out, Outlook responds, JSON comes back. For a single-user Windows workflow it's much less to set up and less to break.

### Why not Microsoft Graph?
Graph requires an Azure app registration, OAuth consent flow, and token refresh. This bridge requires none of that — it rides your existing Outlook session. Trade-off: it only works while you're at your Windows machine with Outlook open.

### Is auto-send safe?
`send`, `reply`, and `forward` send immediately by default. **Always have Claude use `--draft`** unless you explicitly want an outgoing message. The recommended CLAUDE.md snippet in [CLAUDE_INTEGRATION.md](CLAUDE_INTEGRATION.md) bakes this in.

### Does Outlook need to be open?
Yes. The script attaches to the running Outlook process. If Outlook isn't open, the COM dispatch will fail.

### Can it read shared mailboxes?
Any mailbox or folder visible in your Outlook profile is reachable via the `--folder` argument using a path like `Shared Mailbox/Inbox`. Shared calendars need a one-line entry in `outlook_bridge_config.json`.

### Does it use AI / send my email to Anthropic?
No. The bridge itself is pure Python ↔ COM. It only returns email content to whatever process invokes it (typically Claude Code, which then handles it per your normal CC privacy settings).

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

---

## Sample output

`python outlook_bridge.py list --count 2 --pretty`

```json
[
  {
    "EntryID": "00000000ABC123...",
    "Subject": "Q2 renewal — action required",
    "SenderName": "Alice Tan",
    "SenderEmail": "alice@client.com",
    "ReceivedTime": "2026-06-12T09:14:00",
    "Unread": true,
    "HasAttachments": false,
    "BodyPreview": "Hi, following up on the renewal terms we discussed last week..."
  },
  {
    "EntryID": "00000000DEF456...",
    "Subject": "Re: staging deployment",
    "SenderName": "Bob Lee",
    "SenderEmail": "bob@internal.com",
    "ReceivedTime": "2026-06-11T17:42:00",
    "Unread": false,
    "HasAttachments": true,
    "BodyPreview": "Attached the updated config. Let me know if the env vars look right..."
  }
]
```

Pass any `EntryID` to `read` for the full message body, headers, and attachment list.

---

## Related docs

- [CLAUDE_INTEGRATION.md](CLAUDE_INTEGRATION.md) — recommended CLAUDE.md snippet, permission scoping, draft-by-default rules
- [EXAMPLES.md](EXAMPLES.md) — worked Claude prompts for common email workflows
- [AGENTS.md](AGENTS.md) — machine-readable manifest for coding agents (Claude Code, Cursor, Aider)
- [llms.txt](llms.txt) — [llmstxt.org](https://llmstxt.org) index for LLM ingestion

---

## Questions and feedback

- **General enquiries:** [joseph@kainosis.com](mailto:joseph@kainosis.com)
- **Bugs and feature requests:** [open an issue](../../issues)
