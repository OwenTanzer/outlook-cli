# outlook-cli

COM-based command-line access to Outlook on Windows. Read emails, inspect threads, and attach them as PDFs to Linear issues — from the terminal or an AI agent.

## Requirements

- Windows with Outlook desktop installed and a configured account
- Python 3.12+
- Dependencies: `pip install pywin32 click requests`

## Installation

```
pip install pywin32 click requests
```

Add launchers to a directory on your PATH (e.g. `~/.local/bin`):

**PowerShell / cmd** (`outlook-cli.cmd`):
```bat
@echo off
python "C:\path\to\outlook_cli.py" %*
```

**Bash / Git Bash** (`outlook-cli`, no extension):
```sh
#!/bin/sh
python "C:/path/to/outlook_cli.py" "$@"
```
Then `chmod +x ~/.local/bin/outlook-cli`.

> **Note:** Use PowerShell for any command that operates on an individual email (`read`, `attach`, `mark-read`, `flag`, `complete-flag`, `unflag`, `flagged`). These open items via COM, which fails in Git Bash due to apartment model differences. `list` and `folders` work from either shell.

## Commands

### `flagged`
Show all actively flagged emails with body previews in a single pass. No separate `read` calls needed.

```
outlook-cli flagged
outlook-cli flagged --preview 0        # full body
```

Returns newest-first. Body is truncated to 500 characters by default; `--preview 0` returns the full body.

### `list`
List recent emails from your inbox as JSON, ordered newest-first.

```
outlook-cli list
outlook-cli list -n 10
outlook-cli list --unread-only
outlook-cli list --flagged
outlook-cli list --folder sent
```

Without filters, returns the 20 most recent emails. With `--flagged` or `--unread-only`, returns all matches (use `-n` to cap).

**Folders:** `inbox`, `sent`, `drafts`, `deleted`, `outbox`, `junk`

Each result includes a `message_id` field — use this for all per-email commands.

### `read`
Read the full body of an email by its `message_id`.

```
outlook-cli read <message_id>
```

### `attach`
Convert an email to PDF and attach it to an existing Linear issue.

```
outlook-cli attach <message_id> <issue_id>
outlook-cli attach <message_id> ONT-15 --mark-read
```

The PDF is generated via Edge headless, uploaded to Linear's file storage, and linked as a named attachment on the issue with sender and date as subtitle.

### `flag`
Flag an email for follow-up.

```
outlook-cli flag <message_id>
```

### `unflag`
Clear a flag entirely, removing the email from Outlook's flagged/To-Do view.

```
outlook-cli unflag <message_id>
```

Use `unflag` when an email is resolved — it fully removes it from Outlook's native flagged view. Setting `FlagStatus=2` ("complete") does not remove items from Outlook's To-Do list; only clearing all flag state does.

### `mark-read`
Mark an email as read.

```
outlook-cli mark-read <message_id>
```

### `folders`
List all Outlook folders and item counts across all mail stores.

```
outlook-cli folders
```

---

## For AI Assistants

All commands output JSON. Use PowerShell for any command that opens an individual email (`read`, `attach`, `mark-read`, `flag`, `complete-flag`, `unflag`, `flagged`).

### Reviewing flagged emails (recommended)

Use the `flagged` compound command — it returns everything in one COM session with no chaining:

```powershell
$emails = outlook-cli flagged | ConvertFrom-Json
$emails | ForEach-Object { "$($_.received.Substring(0,10)) | $($_.subject)" }
```

To read the full body of one:
```powershell
$emails[0].body
```

To attach one to a Linear issue:
```powershell
outlook-cli attach $emails[0].message_id ONT-15
```

### General workflow

**1. List emails**
```
outlook-cli list -n 20
```
Each item includes:
- `message_id` — stable RFC Message-ID, use for all subsequent commands
- `subject`, `sender`, `sender_email`, `received`, `unread`, `flag_status`, `has_attachments`, `size_bytes`

`flag_status` values: `"none"`, `"flagged"` (active), `"complete"`

**2. Read an email**
```powershell
outlook-cli read <message_id>
```
Same fields as list, plus `body` (plain text).

**3. Attach to a Linear issue**
```powershell
outlook-cli attach <message_id> <issue_id>
```
`issue_id` is a Linear identifier like `ONT-15`. Returns:
```json
{
  "status": "attached",
  "url": "https://uploads.linear.app/...",
  "attachment": { "id": "...", "title": "Email: ...", "url": "..." }
}
```

### Notes
- `message_id` is the RFC 2822 Message-ID — stable across sessions and process boundaries
- Outlook must be running (or will be launched) when any command executes
- `attach` requires Edge at `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`
- Linear OAuth token is read from `~\AppData\Roaming\linear-cli\config.toml`
