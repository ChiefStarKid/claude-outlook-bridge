# Integrating outlook-bridge with Claude Code

This guide explains how to wire `outlook_bridge.py` into Claude Code so Claude can read and act on your email during a conversation.

---

## How it works

Claude Code can run shell commands via the Bash tool. The bridge is a Python script that Claude calls like any CLI tool, receives structured JSON back, and reasons over the result. No MCP server, no webhook, no persistent process — just a script Claude can invoke on demand.

```
Claude Code  →  Bash("python ~/.claude/outlook_bridge.py list")  →  Outlook Desktop App
                                                                  ←  JSON response
```

---

## Setup

### 1. Place the script

Copy `outlook_bridge.py` to `~/.claude/`:

```
C:\Users\YourName\.claude\outlook_bridge.py
```

If you want shared calendar support, copy `outlook_bridge_config.json` to the same location and edit it with your shared calendar email addresses.

### 2. Grant Bash permission

Add a scoped permission to `.claude/settings.json` (project-level) or `~\AppData\Roaming\Claude\claude_desktop_config.json` (global):

```json
{
  "permissions": {
    "allow": ["Bash(python *outlook_bridge.py*:*)"]
  }
}
```

This allows Claude to run the bridge and nothing else automatically.

### 3. Tell Claude about the tool in CLAUDE.md

Add a section to your project's `CLAUDE.md` (or your global `~/.claude/CLAUDE.md`). Here is a ready-to-use snippet:

```markdown
## Outlook Bridge

Script: `C:/Users/YourName/.claude/outlook_bridge.py`

Use this to read and interact with Outlook email and calendar.

**Rules:**
- Always use `--draft` for reply and forward unless the user explicitly asks to send
- Never auto-send without user confirmation
- Use `--pretty` only for display; omit for chained operations

**Commands:**
| What | Command |
|------|---------|
| List folders | `python "C:/Users/YourName/.claude/outlook_bridge.py" folders` |
| List inbox | `python "C:/Users/YourName/.claude/outlook_bridge.py" list --folder inbox --count 20` |
| Search | `python "C:/Users/YourName/.claude/outlook_bridge.py" search "keyword"` |
| Read email | `python "C:/Users/YourName/.claude/outlook_bridge.py" read <entry_id>` |
| Send | `python "C:/Users/YourName/.claude/outlook_bridge.py" send --to addr --subject "..." --body "..."` |
| Reply (draft) | `python "C:/Users/YourName/.claude/outlook_bridge.py" reply <entry_id> --body "..." --draft` |
| Reply all (draft) | `python "C:/Users/YourName/.claude/outlook_bridge.py" reply <entry_id> --body "..." --all --draft` |
| Forward (draft) | `python "C:/Users/YourName/.claude/outlook_bridge.py" forward <entry_id> --to addr --body "..." --draft` |
| Move email | `python "C:/Users/YourName/.claude/outlook_bridge.py" move <entry_id> --folder "FolderName"` |
| Delete email | `python "C:/Users/YourName/.claude/outlook_bridge.py" delete <entry_id>` |
| Calendar | `python "C:/Users/YourName/.claude/outlook_bridge.py" cal-list --from YYYY-MM-DD --to YYYY-MM-DD` |
```

Replace `YourName` with your actual Windows username. Quoting the path handles spaces.

---

## Recommended defaults

**Always draft, never auto-send.** The bridge can send email immediately — which is powerful but unforgiving. The safest pattern is to instruct Claude in CLAUDE.md to always use `--draft` for `reply` and `forward`, and to confirm with you before running `send`. The example snippet above includes this rule.

**Pipe entry IDs, not subjects.** When chaining commands (e.g. search then reply), always pass the `id` field from the JSON result rather than trying to match by subject. Entry IDs are stable; subjects are not unique.

**Use project-level permissions.** If you only want Claude to have email access in specific projects, put the `allow` entry in `.claude/settings.json` inside that project folder — not in global settings. You can remove it at any time to revoke access.

---

## Troubleshooting

**`pywin32 not installed`** — Run `pip install pywin32` in the Python environment Claude is using.

**`Cannot connect to Outlook`** — Outlook Desktop App must be open and signed in before running any command.

**`Folder not found`** — Use the `folders` command to see the exact folder names in your mailbox. Folder names are case-insensitive but must match exactly (spaces included).

**`Shared calendar not found in config`** — Check that `outlook_bridge_config.json` exists in the same directory as the script, and that the key you're passing to `--cal` matches a key in `shared_calendars`.

**`Could not resolve shared calendar`** — The email address in your config must correspond to a calendar you have delegate/read access to in Outlook. Open the calendar in Outlook manually first to confirm access.
