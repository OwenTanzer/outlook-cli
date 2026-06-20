#!/usr/bin/env python3
"""Outlook CLI — COM-based command-line access to Outlook on Windows."""

import sys
import json
import subprocess
from pathlib import Path

import click
import win32com.client

LINEAR_CLI = str(Path.home() / ".cargo" / "bin" / "linear-cli.exe")

# olDefaultFolders constants
FOLDER_IDS = {
    "inbox": 6,
    "sent": 5,
    "drafts": 16,
    "deleted": 3,
    "outbox": 4,
    "junk": 23,
}


def get_mapi():
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


def mail_to_dict(msg, index=None, include_body=False):
    try:
        received = msg.ReceivedTime.isoformat()
    except Exception:
        received = None
    d = {
        "index": index,
        "entry_id": msg.EntryID,
        "subject": msg.Subject or "(no subject)",
        "sender": msg.SenderName,
        "sender_email": msg.SenderEmailAddress,
        "received": received,
        "unread": bool(msg.UnRead),
        "has_attachments": bool(msg.Attachments.Count),
        "size_bytes": msg.Size,
    }
    if include_body:
        d["body"] = msg.Body
    return d


@click.group()
def cli():
    """Outlook CLI — COM-based command-line access to Outlook on Windows."""
    pass


def _find_folder(mapi, folder_name):
    """Find a named folder across all mail stores, newest-first."""
    target = folder_name.lower().replace(" ", "")
    aliases = {
        "inbox": ["inbox"],
        "sent": ["sentitems", "sent"],
        "drafts": ["drafts"],
        "deleted": ["deleteditems", "deleted"],
        "outbox": ["outbox"],
        "junk": ["junkemail", "junk", "spam"],
    }
    candidates = aliases.get(target, [target])
    for store in mapi.Folders:
        for sub in store.Folders:
            sub_key = sub.Name.lower().replace(" ", "")
            if sub_key in candidates:
                return sub
    # Fallback: default inbox
    return mapi.GetDefaultFolder(6)


@cli.command("list")
@click.option("--count", "-n", default=20, show_default=True, help="Max emails to return")
@click.option("--unread-only", is_flag=True, help="Only show unread emails")
@click.option("--folder", default="inbox", show_default=True,
              help="Folder name: inbox, sent, drafts, deleted, outbox, junk")
def list_emails(count, unread_only, folder):
    """List recent emails. Outputs JSON array ordered newest-first."""
    mapi = get_mapi()
    folder_obj = _find_folder(mapi, folder)
    items = folder_obj.Items
    items.Sort("[ReceivedTime]", True)

    results = []
    idx = 0
    for msg in items:
        if len(results) >= count:
            break
        try:
            if getattr(msg, "Class", None) != 43:  # 43 = olMail
                continue
            if unread_only and not msg.UnRead:
                continue
            results.append(mail_to_dict(msg, index=idx))
            idx += 1
        except Exception:
            continue

    print(json.dumps(results, indent=2, default=str))


@cli.command("read")
@click.argument("entry_id")
def read_email(entry_id):
    """Read full email body by EntryID. Get EntryIDs from 'list'."""
    mapi = get_mapi()
    msg = mapi.GetItemFromID(entry_id)
    data = mail_to_dict(msg, include_body=True)
    print(json.dumps(data, indent=2, default=str))


@cli.command("attach")
@click.argument("entry_id")
@click.argument("issue_id")
@click.option("--mark-read", is_flag=True, help="Mark the email as read after attaching")
def attach(entry_id, issue_id, mark_read):
    """Attach an email to an existing Linear issue as a comment.

    ENTRY_ID  — Outlook email EntryID (from 'list')
    ISSUE_ID  — Linear issue identifier, e.g. ONT-18
    """
    mapi = get_mapi()
    msg = mapi.GetItemFromID(entry_id)

    body = (
        f"**From:** {msg.SenderName} <{msg.SenderEmailAddress}>  \n"
        f"**Subject:** {msg.Subject or '(no subject)'}  \n"
        f"**Received:** {msg.ReceivedTime}  \n\n"
        f"---\n\n"
        f"{msg.Body}"
    )

    cmd = [LINEAR_CLI, "comments", "create", issue_id, "--body", body, "--output", "json", "--quiet"]
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        print(json.dumps({"error": result.stderr.strip()}), file=sys.stderr)
        sys.exit(1)

    if mark_read:
        msg.UnRead = False
        msg.Save()

    try:
        comment_data = json.loads(result.stdout)
        print(json.dumps({"status": "attached", "marked_read": mark_read, "comment": comment_data}, indent=2))
    except json.JSONDecodeError:
        print(result.stdout)


@cli.command("mark-read")
@click.argument("entry_id")
def mark_read_cmd(entry_id):
    """Mark an email as read by EntryID."""
    mapi = get_mapi()
    msg = mapi.GetItemFromID(entry_id)
    msg.UnRead = False
    msg.Save()
    print(json.dumps({"status": "ok", "entry_id": entry_id}))


@cli.command("folders")
def list_folders():
    """List top-level Outlook folders and subfolders with item counts."""
    mapi = get_mapi()
    folders = []
    for store in mapi.Folders:
        for sub in store.Folders:
            try:
                folders.append({
                    "store": store.Name,
                    "folder": sub.Name,
                    "count": sub.Items.Count,
                })
            except Exception:
                continue
    print(json.dumps(folders, indent=2))


if __name__ == "__main__":
    cli()
