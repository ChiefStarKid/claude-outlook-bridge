# AGENTS.md

Instructions for coding agents (Claude Code, Cursor, Aider, etc.) wiring this bridge into a user's project.

## What this is

`claude-outlook-bridge` is a single Python file (`outlook_bridge.py`) that an agent invokes via Bash to read and send email through the user's already-open Outlook Desktop App on Windows. No auth setup. All output is JSON on stdout; errors exit non-zero with `{"error": "..."}`.

## Auto-wire checklist

When a user asks you to set up Outlook access in their project, do all of the following:

1. **Clone or download** to a stable path. Recommend `~/.claude/outlook-bridge/` so the script can be shared across projects:
   ```bash
   git clone https://github.com/ChiefStarKid/claude-outlook-bridge.git ~/.claude/outlook-bridge
   pip install -r ~/.claude/outlook-bridge/requirements.txt
   ```

2. **Smoke test** before wiring it in:
   ```bash
   python ~/.claude/outlook-bridge/outlook_bridge.py folders --pretty
   ```
   If this returns folder JSON, Outlook is reachable. If it errors, Outlook isn't open or `pywin32` isn't installed — fix before continuing.

3. **Add the permission** to the project's `.claude/settings.json` (create the file if missing):
   ```json
   {
     "permissions": {
       "allow": ["Bash(python *outlook_bridge.py*:*)"]
     }
   }
   ```

4. **Add a CLAUDE.md snippet** so the assistant knows the bridge exists and the safety rules:
   ```markdown
   ## Outlook email access

   Outlook is reachable via:
   `python ~/.claude/outlook-bridge/outlook_bridge.py <command> [flags]`

   Commands: `folders`, `list`, `search`, `read`, `send`, `reply`, `forward`, `move`, `delete`, `cal-list`.
   All output is JSON. Use `--pretty` for human display.

   **Safety:** `send`, `reply`, `forward` send immediately by default.
   ALWAYS pass `--draft` on `reply` and `forward` unless the user explicitly
   asks to send. For new outgoing mail, draft in Outlook and ask the user
   to review before using `send`.
   ```

## Command manifest

| Command | Required args | Side effects | Idempotent | Notes |
|---|---|---|---|---|
| `folders` | — | none | yes | Returns full folder tree as JSON. |
| `list` | — | none | yes | `--folder <name|path>` `--count <n>`. Returns array of `{EntryID, Subject, From, ReceivedTime, ...}`. |
| `search` | `<query>` | none | yes | Full-text across subject/body/sender/recipient. Same flags as `list`. |
| `read` | `<entry_id>` | none | yes | Returns full body + attachment metadata. `EntryID` comes from `list`/`search`. |
| `send` | `--to --subject --body` | **sends mail immediately** | no | No `--draft` flag — there is no draft mode for new mail in the current bridge. Confirm with user before invoking. |
| `reply` | `<entry_id> --body` | **sends mail immediately** unless `--draft` | no | Always pass `--draft` by default. `--all` for reply-all. |
| `forward` | `<entry_id> --to` | **sends mail immediately** unless `--draft` | no | Always pass `--draft` by default. `--body` adds an optional note. |
| `move` | `<entry_id> --folder` | mutates mailbox state | yes (same target) | Folder path supports nesting (`Projects/Client`). |
| `delete` | `<entry_id>` | moves to Deleted Items | yes | Recoverable from Deleted Items. |
| `cal-list` | `--from --to` | none | yes | Dates `YYYY-MM-DD`. `--cal <name>` for shared calendars (requires config). |

## Output contract

- Success: JSON object or array on stdout, exit 0.
- Failure: `{"error": "<message>"}` on stdout, exit 1.
- `--pretty` reformats stdout but does not change the schema. Do not pass `--pretty` when piping to another command — parse the compact JSON directly.

## Hard rules for agents

1. **Default to `--draft`** on `reply` and `forward`. Only omit it when the user has explicitly said "send" in the current turn.
2. **Confirm before `send`** (which has no draft mode). State the recipients, subject, and a one-line body summary, and wait for the user to approve.
3. **Confirm before `delete`** of anything that wasn't obviously transient (e.g. a notification the user told you to clean up).
4. **Don't loop `list` to scrape the inbox.** Use `search` with a query. Repeated `list` calls bloat context.
5. **Pass EntryIDs through verbatim.** They're opaque tokens — don't truncate, normalise, or quote-escape them.
6. **Don't assume Outlook is open.** If a command errors with a COM dispatch failure, surface the error to the user instead of retrying.

## Repo

Canonical source: <https://github.com/ChiefStarKid/claude-outlook-bridge>
