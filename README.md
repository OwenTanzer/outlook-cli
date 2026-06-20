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

Add `outlook-cli.cmd` to a directory on your PATH (e.g. `~/.local/bin`):

```bat
@echo off
python "C:\path\to\outlook_cli.py" %*
```

## Commands

### `list`
List recent emails from your inbox as JSON, ordered newest-first.

```
outlook-cli list
outlook-cli list -n 10
outlook-cli list --unread-only
outlook-cli list --folder sent
```

**Folders:** `inbox`, `sent`, `drafts`, `deleted`, `outbox`, `junk`

### `read`
Read the full body of an email by its `entry_id`.

```
outlook-cli read <entry_id>
```

### `attach`
Convert an email to PDF and attach it to an existing Linear issue.

```
outlook-cli attach <entry_id> <issue_id>
outlook-cli attach <entry_id> ONT-15 --mark-read
```

The PDF is generated via Edge headless, uploaded to Linear's file storage, and linked as a named attachment on the issue with sender and date as subtitle.

### `mark-read`
Mark an email as read.

```
outlook-cli mark-read <entry_id>
```

### `folders`
List all Outlook folders and item counts across all mail stores.

```
outlook-cli folders
```

---

## For AI Assistants

All commands output JSON. The standard workflow is:

**1. List emails**
```
outlook-cli list -n 20
```
Returns an array. Each item includes:
- `index` — position in the sorted list (0 = newest)
- `entry_id` — stable identifier, use this for all subsequent commands
- `subject`, `sender`, `sender_email`, `received`, `unread`, `has_attachments`, `size_bytes`

**2. Read an email**
```
outlook-cli read <entry_id>
```
Same fields as list, plus `body` (plain text).

**3. Attach to a Linear issue**
```
outlook-cli attach <entry_id> <issue_id>
```
`issue_id` is a Linear identifier like `ONT-15`. Returns:
```json
{
  "status": "attached",
  "url": "https://uploads.linear.app/...",
  "attachment": { "id": "...", "title": "Email: ...", "url": "..." }
}
```

**Notes:**
- `entry_id` values are stable across sessions — safe to pass between tool calls
- Outlook must be running (or will be launched) when any command executes
- The `attach` command requires Edge (`C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`) and a valid Linear OAuth token in `~\AppData\Roaming\linear-cli\config.toml`
- Linear team in use: `Private Workflow` (identifiers prefixed `ONT-`)
