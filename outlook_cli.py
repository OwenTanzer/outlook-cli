#!/usr/bin/env python3
"""Outlook CLI — COM-based command-line access to Outlook on Windows."""

import sys
import json
import subprocess
import tempfile
import os
import tomllib
from pathlib import Path

import click
import requests
import win32com.client

LINEAR_CLI = str(Path.home() / ".cargo" / "bin" / "linear-cli.exe")
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
LINEAR_API = "https://api.linear.app/graphql"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"


def _linear_token():
    config = Path.home() / "AppData/Roaming/linear-cli/config.toml"
    with open(config, "rb") as f:
        data = tomllib.load(f)
    workspace = data.get("current", "default")
    return data["workspaces"][workspace]["oauth"]["access_token"]


def _linear_query(query, variables=None):
    token = _linear_token()
    resp = requests.post(
        LINEAR_API,
        json={"query": query, "variables": variables or {}},
        headers={"Authorization": token},
    )
    resp.raise_for_status()
    return resp.json()


def get_mapi():
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


def _internet_message_id(msg):
    try:
        mid = msg.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
        return mid.strip("<>") if mid else None
    except Exception:
        return None


def _get_item(mapi, message_id):
    """Find a mail item by InternetMessageId — stable across sessions."""
    for store in mapi.Folders:
        for folder in store.Folders:
            try:
                for item in folder.Items:
                    try:
                        if _internet_message_id(item) == message_id:
                            return item
                    except Exception:
                        pass
            except Exception:
                pass
    raise click.ClickException(f"Email not found: {message_id}")


def mail_to_dict(msg, index=None, include_body=False):
    try:
        received = msg.ReceivedTime.isoformat()
    except Exception:
        received = None
    d = {
        "index": index,
        "message_id": _internet_message_id(msg),
        "subject": msg.Subject or "(no subject)",
        "sender": msg.SenderName,
        "sender_email": msg.SenderEmailAddress,
        "received": received,
        "unread": bool(msg.UnRead),
        "flag_status": {0: "none", 1: "flagged", 2: "complete"}.get(getattr(msg, "FlagStatus", 0), "none"),
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
            if sub.Name.lower().replace(" ", "") in candidates:
                return sub
    return mapi.GetDefaultFolder(6)


@cli.command("list")
@click.option("--count", "-n", default=None, type=int,
              help="Max emails to return (default: 20; unlimited when filtering)")
@click.option("--unread-only", is_flag=True, help="Only show unread emails")
@click.option("--flagged", is_flag=True, help="Only show actively flagged emails")
@click.option("--flag-complete", is_flag=True, help="Only show flag-completed emails")
@click.option("--folder", default="inbox", show_default=True,
              help="Folder: inbox, sent, drafts, deleted, outbox, junk")
def list_emails(count, unread_only, flagged, flag_complete, folder):
    """List recent emails as JSON, newest-first.

    Filtering options (--flagged, --flag-complete, --unread-only) return all
    matches by default. Use --count to cap results.
    """
    mapi = get_mapi()
    is_filtered = unread_only or flagged or flag_complete
    limit = count if count is not None else (None if is_filtered else 20)

    folder_obj = _find_folder(mapi, folder)
    items = folder_obj.Items
    items.Sort("[ReceivedTime]", True)

    results = []
    idx = 0
    for msg in items:
        if limit is not None and len(results) >= limit and not is_filtered:
            break
        try:
            if getattr(msg, "Class", None) != 43:
                continue
            if unread_only and not msg.UnRead:
                continue
            if flagged and getattr(msg, "FlagStatus", 0) != 1:
                continue
            if flag_complete and getattr(msg, "FlagStatus", 0) != 2:
                continue
            results.append(mail_to_dict(msg, index=idx))
            idx += 1
        except Exception:
            continue

    if limit is not None:
        results = results[:limit]

    print(json.dumps(results, indent=2, default=str))


@cli.command("flagged")
@click.option("--include-complete", is_flag=True, help="Also include flag-completed emails")
@click.option("--preview", default=500, show_default=True,
              help="Body preview length in characters (0 for full body)")
def show_flagged(include_complete, preview):
    """Show all flagged emails with body previews in a single pass.

    Returns everything needed for review — no separate 'read' calls required.
    """
    mapi = get_mapi()
    folder = _find_folder(mapi, "inbox")
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    results = []
    for msg in items:
        try:
            if getattr(msg, "Class", None) != 43:
                continue
            fs = getattr(msg, "FlagStatus", 0)
            if include_complete and fs not in (1, 2):
                continue
            if not include_complete and fs != 1:
                continue
            data = mail_to_dict(msg, include_body=True)
            if preview > 0 and data.get("body"):
                data["body"] = data["body"][:preview]
            results.append(data)
        except Exception:
            continue

    print(json.dumps(results, indent=2, default=str))


@cli.command("read")
@click.argument("message_id")
def read_email(message_id):
    """Read a full email by its message_id (from 'list' output)."""
    mapi = get_mapi()
    msg = _get_item(mapi, message_id)
    print(json.dumps(mail_to_dict(msg, include_body=True), indent=2, default=str))


@cli.command("attach")
@click.argument("message_id")
@click.argument("issue_id")
@click.option("--mark-read", is_flag=True, help="Mark the email as read after attaching")
def attach(message_id, issue_id, mark_read):
    """Attach an email to an existing Linear issue as a PDF.

    MESSAGE_ID — from 'list' or 'flagged' output (message_id field)
    ISSUE_ID   — Linear issue identifier, e.g. ONT-15
    """
    mapi = get_mapi()
    msg = _get_item(mapi, message_id)
    subject = msg.Subject or "email"

    with tempfile.TemporaryDirectory() as tmp:
        html_path = os.path.join(tmp, "email.html")
        pdf_path = os.path.join(tmp, "email.pdf")

        msg.SaveAs(html_path, 5)
        subprocess.run(
            [EDGE, "--headless", f"--print-to-pdf={pdf_path}", f"file:///{html_path}"],
            check=True, capture_output=True,
        )

        pdf_size = os.path.getsize(pdf_path)
        filename = f"{subject[:60].replace('/', '-')}.pdf"

        upload_resp = _linear_query("""
            mutation Upload($filename: String!, $size: Int!) {
              fileUpload(filename: $filename, contentType: "application/pdf", size: $size) {
                success
                uploadFile { uploadUrl assetUrl headers { key value } }
              }
            }
        """, {"filename": filename, "size": pdf_size})

        upload = upload_resp["data"]["fileUpload"]["uploadFile"]
        upload_headers = {h["key"]: h["value"] for h in upload["headers"]}
        upload_headers["Content-Type"] = "application/pdf"

        with open(pdf_path, "rb") as f:
            put_resp = requests.put(upload["uploadUrl"], data=f, headers=upload_headers)
        put_resp.raise_for_status()
        asset_url = upload["assetUrl"]

    result = subprocess.run(
        [LINEAR_CLI, "attachments", "create", issue_id,
         "--title", f"Email: {subject}",
         "--url", asset_url,
         "--subtitle", f"From: {msg.SenderName} • {msg.ReceivedTime}",
         "--output", "json", "--quiet"],
        capture_output=True, text=True,
    )

    if result.returncode != 0:
        print(json.dumps({"error": result.stderr.strip()}), file=sys.stderr)
        sys.exit(1)

    if mark_read:
        msg.UnRead = False
        msg.Save()

    try:
        att_data = json.loads(result.stdout)
        print(json.dumps({"status": "attached", "url": asset_url, "attachment": att_data}, indent=2))
    except json.JSONDecodeError:
        print(result.stdout)


@cli.command("mark-read")
@click.argument("message_id")
def mark_read_cmd(message_id):
    """Mark an email as read by message_id."""
    mapi = get_mapi()
    msg = _get_item(mapi, message_id)
    msg.UnRead = False
    msg.Save()
    print(json.dumps({"status": "ok", "message_id": message_id}))


@cli.command("folders")
def list_folders():
    """List Outlook folders and item counts."""
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
