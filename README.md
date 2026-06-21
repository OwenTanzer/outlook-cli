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

> **Note:** Use PowerShell for any command that operates on an individual email (`read`, `attach`, `mark-read`, `flag`, `unflag`, `flagged`). These open items via COM, which fails in Git Bash due to apartment model differences. `list`, `cleanup`, and `folders` work from either shell.

## Commands

### `flagged`
Show all actively flagged emails with body previews in a single pass. No separate `read` calls needed.

```
outlook-cli flagged
outlook-cli flagged --preview 0        # full body
outlook-cli flagged --format table     # human-readable view with body snippet
```

Returns newest-first. Body is truncated to 500 characters by default; `--preview 0` returns the full body. `--format table` (default) prints a columnar view with a 120-character body snippet per email. Use `--format json` for machine-readable output.

### `list`
List recent emails from your inbox, ordered newest-first.

```
outlook-cli list
outlook-cli list -n 10
outlook-cli list --unread-only
outlook-cli list --flagged
outlook-cli list --folder sent
outlook-cli list --unread-only --format table
```

Without filters, returns the 20 most recent emails. With `--flagged` or `--unread-only`, returns all matches (use `-n` to cap).

**Folders:** `inbox`, `sent`, `drafts`, `deleted`, `outbox`, `junk`

`--format table` prints a human-readable columnar view (default). Columns: date, unread (•), flagged (F), has-attachment (@), sender, subject. Use `--format json` for machine-readable output.

Each JSON result includes a `message_id` field — use this for all per-email commands.

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
outlook-cli attach <message_id> ONT-15 --save-attachments
```

The PDF is generated via Edge headless, uploaded to Linear's file storage, and linked as a named attachment on the issue with sender and date as subtitle.

`--save-attachments` additionally uploads any file attachments from the email (PDFs, Word docs, etc.) as separate Linear attachments on the same issue. Use this when the email attachment is the primary content.

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
Mark one or more emails as read in a single COM session.

```
outlook-cli mark-read <message_id>
outlook-cli mark-read <id1> <id2> <id3> ...
```

Accepts any number of message IDs. All are processed in one Outlook session — no per-email startup cost.

### `linear-update`
Update a Linear issue via GraphQL for fields `linear-cli` doesn't expose — set/clear parent, update description or title.

```
outlook-cli linear-update ONT-40 --parent ONT-4
outlook-cli linear-update ONT-40 --unparent
outlook-cli linear-update ONT-40 --description "Some description text"
outlook-cli linear-update ONT-40 --title "New title" --parent ONT-4
```

Flags can be combined. Uses the same Linear token as `attach`.

### `cleanup`
Scan the inbox for stale `FlagStatus=2` items (emails previously marked complete via old CLI versions) and clear them, removing them from Outlook's To-Do view.

```
outlook-cli cleanup
```

### `folders`
List all Outlook folders and item counts across all mail stores.

```
outlook-cli folders
```

---

## For AI Assistants

All commands output JSON. Use PowerShell for any command that opens an individual email (`read`, `attach`, `mark-read`, `flag`, `unflag`, `flagged`).

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
outlook-cli attach $emails[0].message_id ONT-15 --save-attachments  # also uploads file attachments
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
outlook-cli attach <message_id> <issue_id> --save-attachments
```
`issue_id` is a Linear identifier like `ONT-15`. Returns:
```json
{
  "status": "attached",
  "url": "https://uploads.linear.app/...",
  "attachment": { "id": "...", "title": "Email: ...", "url": "..." },
  "file_attachments": ["manuscript.docx"]
}
```
`file_attachments` lists any email file attachments uploaded (empty list if `--save-attachments` not used or email has none).

### Notes
- `message_id` is the RFC 2822 Message-ID — stable across sessions and process boundaries
- Outlook must be running (or will be launched) when any command executes
- `attach` requires Edge at `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`
- Linear OAuth token is read from `~\AppData\Roaming\linear-cli\config.toml`
