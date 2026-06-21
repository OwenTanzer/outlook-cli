#!/usr/bin/env python3
"""Outlook CLI — COM-based command-line access to Outlook on Windows."""

import sys
import json
import subprocess
import tempfile
import os
import tomllib
import time
from pathlib import Path

import mimetypes
import click
import requests
import win32com.client

_debug = False
_t0 = time.perf_counter()

def _dbg(msg):
    if _debug:
        elapsed = time.perf_counter() - _t0
        print(f"[{elapsed:.3f}s] {msg}", file=sys.stderr)

LINEAR_CLI = str(Path.home() / ".cargo" / "bin" / "linear-cli.exe")
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
LINEAR_API = "https://api.linear.app/graphql"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_TODO_ITEM_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E2B0003"


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
    _dbg("COM dispatch start")
    outlook = win32com.client.Dispatch("Outlook.Application")
    _dbg("COM dispatch done")
    mapi = outlook.GetNamespace("MAPI")
    _dbg("MAPI namespace acquired")
    return mapi


def _internet_message_id(msg):
    try:
        mid = msg.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
        return mid.strip("<>") if mid else None
    except Exception:
        return None


_todo_calls = 0

def _todo_flagged(msg):
    """Return True if item is flagged via the M365 To-Do system (PR_TODO_ITEM_FLAGS bit 0)."""
    global _todo_calls
    _todo_calls += 1
    try:
        flags = msg.PropertyAccessor.GetProperty(PR_TODO_ITEM_FLAGS)
        return bool(flags and (flags & 1))
    except Exception:
        return False


def _flag_status(msg):
    """Resolve flag state, matching Outlook's native search behaviour.

    PR_TODO_ITEM_FLAGS (the M365 To-Do bit) is checked first because it is
    authoritative for what Outlook's 'Flagged' search view shows. Some items
    have FlagStatus=2 (classic 'complete') but PR_TODO=1 (still active in
    To-Do) — Outlook treats those as active flags, so we do too.
    """
    if _todo_flagged(msg):
        return "flagged"
    classic = getattr(msg, "FlagStatus", 0)
    if classic == 1:
        return "flagged"
    if classic == 2:
        return "complete"
    return "none"


def _get_item(mapi, message_id):
    """Find a mail item by InternetMessageId — Restrict-first, full-scan fallback."""
    filter_str = f'@SQL="{PR_INTERNET_MESSAGE_ID}" = \'<{message_id}>\''

    # Fast path: DASL Restrict on inbox (~15ms, server-side indexed lookup).
    _dbg("_get_item: trying DASL Restrict on inbox")
    try:
        inbox = _find_folder(mapi, "inbox")
        item = inbox.Items.Restrict(filter_str).GetFirst()
        if item:
            _dbg("_get_item: found via Restrict")
            return item
    except Exception as e:
        _dbg(f"_get_item: Restrict failed ({e})")

    # Fallback: full scan across all folders (catches sent items, subfolders, etc.)
    _dbg("_get_item: falling back to full scan")
    scanned = 0
    for store in mapi.Folders:
        for folder in store.Folders:
            try:
                for item in folder.Items:
                    scanned += 1
                    try:
                        if _internet_message_id(item) == message_id:
                            _dbg(f"_get_item: found via scan after {scanned} items")
                            return item
                    except Exception:
                        pass
            except Exception:
                pass
    _dbg(f"_get_item: exhausted {scanned} items — not found")
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
        "flag_status": _flag_status(msg),
        "has_attachments": bool(msg.Attachments.Count),
        "size_bytes": msg.Size,
    }
    if include_body:
        d["body"] = msg.Body
    return d


@click.group()
@click.option("--debug", is_flag=True, help="Print timing info to stderr")
def cli(debug):
    """Outlook CLI — COM-based command-line access to Outlook on Windows."""
    global _debug
    _debug = debug


def _find_folder(mapi, folder_name):
    _dbg(f"_find_folder: {folder_name}")
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
                _dbg(f"_find_folder: found '{sub.Name}' in store '{store.Name}'")
                return sub
    return mapi.GetDefaultFolder(6)


@cli.command("list")
@click.option("--count", "-n", default=None, type=int,
              help="Max emails to return (default: 20; unlimited when filtering)")
@click.option("--unread-only", is_flag=True, help="Only show unread emails")
@click.option("--flagged", is_flag=True, help="Only show actively flagged emails")
@click.option("--folder", default="inbox", show_default=True,
              help="Folder: inbox, sent, drafts, deleted, outbox, junk")
def list_emails(count, unread_only, flagged, folder):
    """List recent emails as JSON, newest-first.

    Filtering options (--flagged, --unread-only) return all matches by default.
    Use --count to cap results.
    """
    mapi = get_mapi()
    is_filtered = unread_only or flagged
    limit = count if count is not None else (None if is_filtered else 20)

    folder_obj = _find_folder(mapi, folder)
    items = folder_obj.Items
    _dbg(f"list: folder has {items.Count} items, limit={limit}, filtered={is_filtered}")

    collected = []
    scanned = 0

    if not is_filtered and limit is not None:
        # Fast path: Sort() + GetFirst/GetNext — stops at limit without scanning everything.
        # GetFirst/GetNext uses COM's own cursor and is documented to work correctly with
        # Sort(), unlike Python's for-loop enumerator which skips items on sorted collections.
        # Verified against scan-all-sort-in-Python via verify_sort.py before enabling.
        _dbg("list: using Sort()+GetFirst/GetNext fast path")
        items.Sort("[ReceivedTime]", True)
        msg = items.GetFirst()
        while msg and len(collected) < limit:
            scanned += 1
            try:
                if getattr(msg, "Class", None) == 43:
                    collected.append(mail_to_dict(msg))
            except Exception:
                pass
            msg = items.GetNext()
    elif flagged and not unread_only:
        # DASL Restrict fast path for --flagged: server-side filter on PR_TODO_ITEM_FLAGS.
        # Verified via verify_dasl_flagged.py: 29/29 correct, 0 misses, ~0.5s vs ~54s scan.
        # Falls back to full scan if Restrict raises.
        _dbg("list: using DASL Restrict fast path for --flagged")
        restrict_ok = False
        try:
            restricted = items.Restrict(f'@SQL="{PR_TODO_ITEM_FLAGS}" = 1')
            msg = restricted.GetFirst()
            while msg:
                scanned += 1
                try:
                    if getattr(msg, "Class", None) == 43:
                        collected.append(mail_to_dict(msg))
                except Exception:
                    pass
                msg = restricted.GetNext()
            restrict_ok = True
        except Exception as e:
            _dbg(f"list: Restrict failed ({e}), falling back to full scan")

        if not restrict_ok:
            for msg in items:
                scanned += 1
                try:
                    if getattr(msg, "Class", None) != 43:
                        continue
                    if not _todo_flagged(msg):
                        continue
                    collected.append(mail_to_dict(msg))
                except Exception:
                    continue

        collected.sort(key=lambda m: m.get("received") or "", reverse=True)
        if limit is not None:
            collected = collected[:limit]
    else:
        # --unread-only or combined filters: must scan everything.
        # Do NOT call items.Sort() — causes COM enumerator to skip items.
        for msg in items:
            scanned += 1
            try:
                if getattr(msg, "Class", None) != 43:
                    continue
                if unread_only and not msg.UnRead:
                    continue
                if flagged and not _todo_flagged(msg):
                    continue
                collected.append(mail_to_dict(msg))
            except Exception:
                continue
        collected.sort(key=lambda m: m.get("received") or "", reverse=True)
        if limit is not None:
            collected = collected[:limit]

    _dbg(f"list: scanned {scanned} items, {_todo_calls} PropertyAccessor calls, {len(collected)} matched")
    for idx, m in enumerate(collected):
        m["index"] = idx

    _dbg(f"list: done, returning {len(collected)} results")
    print(json.dumps(collected, indent=2, default=str))


@cli.command("flagged")
@click.option("--preview", default=500, show_default=True,
              help="Body preview length in characters (0 for full body)")
def show_flagged(preview):
    """Show all actively flagged emails with body previews in a single pass.

    Returns everything needed for review — no separate 'read' calls required.
    Use 'unflag' to dismiss an email from Outlook's flagged view when done.
    """
    mapi = get_mapi()
    folder = _find_folder(mapi, "inbox")
    items = folder.Items
    _dbg(f"flagged: folder has {items.Count} items")

    results = []
    scanned = 0

    # DASL Restrict fast path: server-side filter on PR_TODO_ITEM_FLAGS.
    # Verified via verify_dasl_flagged.py: 29/29 correct, 0 misses.
    _dbg("flagged: trying DASL Restrict fast path")
    try:
        restricted = items.Restrict(f'@SQL="{PR_TODO_ITEM_FLAGS}" = 1')
        msg = restricted.GetFirst()
        while msg:
            scanned += 1
            try:
                if getattr(msg, "Class", None) == 43:
                    data = mail_to_dict(msg, include_body=True)
                    if preview > 0 and data.get("body"):
                        data["body"] = data["body"][:preview]
                    results.append(data)
            except Exception:
                pass
            msg = restricted.GetNext()
        _dbg(f"flagged: Restrict returned {scanned} items, {len(results)} matched")
    except Exception as e:
        # Fallback: full scan. Do NOT call items.Sort() — causes COM enumerator to skip items.
        _dbg(f"flagged: Restrict failed ({e}), falling back to full scan")
        for msg in items:
            scanned += 1
            try:
                if getattr(msg, "Class", None) != 43:
                    continue
                if not _todo_flagged(msg):
                    continue
                data = mail_to_dict(msg, include_body=True)
                if preview > 0 and data.get("body"):
                    data["body"] = data["body"][:preview]
                results.append(data)
            except Exception:
                continue
        _dbg(f"flagged: scanned {scanned} items, {len(results)} matched")

    results.sort(key=lambda m: m.get("received") or "", reverse=True)
    _dbg("flagged: done")
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
@click.option("--save-attachments", is_flag=True, help="Also upload email file attachments to the Linear issue")
def attach(message_id, issue_id, mark_read, save_attachments):
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
    except json.JSONDecodeError:
        att_data = {}

    uploaded_attachments = []
    if save_attachments and msg.Attachments.Count > 0:
        with tempfile.TemporaryDirectory() as att_tmp:
            for i in range(1, msg.Attachments.Count + 1):
                att = msg.Attachments.Item(i)
                att_filename = att.FileName
                if not att_filename:
                    continue
                att_path = os.path.join(att_tmp, att_filename)
                try:
                    att.SaveAsFile(att_path)
                    att_size = os.path.getsize(att_path)
                    content_type = mimetypes.guess_type(att_filename)[0] or "application/octet-stream"

                    up = _linear_query("""
                        mutation Upload($filename: String!, $size: Int!, $ct: String!) {
                          fileUpload(filename: $filename, contentType: $ct, size: $size) {
                            success
                            uploadFile { uploadUrl assetUrl headers { key value } }
                          }
                        }
                    """, {"filename": att_filename, "size": att_size, "ct": content_type})

                    uf = up["data"]["fileUpload"]["uploadFile"]
                    uh = {h["key"]: h["value"] for h in uf["headers"]}
                    uh["Content-Type"] = content_type
                    with open(att_path, "rb") as f:
                        requests.put(uf["uploadUrl"], data=f, headers=uh).raise_for_status()

                    subprocess.run(
                        [LINEAR_CLI, "attachments", "create", issue_id,
                         "--title", att_filename,
                         "--url", uf["assetUrl"],
                         "--subtitle", f"Attachment from: {msg.SenderName}"],
                        capture_output=True,
                    )
                    uploaded_attachments.append(att_filename)
                    _dbg(f"attach: uploaded file attachment '{att_filename}'")
                except Exception as e:
                    _dbg(f"attach: failed to upload '{att_filename}' ({e})")

    print(json.dumps({
        "status": "attached",
        "url": asset_url,
        "attachment": att_data,
        "file_attachments": uploaded_attachments,
    }, indent=2))


@cli.command("mark-read")
@click.argument("message_id")
def mark_read_cmd(message_id):
    """Mark an email as read by message_id."""
    mapi = get_mapi()
    msg = _get_item(mapi, message_id)
    msg.UnRead = False
    msg.Save()
    print(json.dumps({"status": "ok", "message_id": message_id}))


@cli.command("flag")
@click.argument("message_id")
def flag_email(message_id):
    """Flag an email for follow-up by message_id."""
    mapi = get_mapi()
    msg = _get_item(mapi, message_id)
    msg.FlagStatus = 1
    msg.PropertyAccessor.SetProperty(PR_TODO_ITEM_FLAGS, 1)
    msg.Save()
    print(json.dumps({"status": "flagged", "message_id": message_id, "subject": msg.Subject}))


@cli.command("unflag")
@click.argument("message_id")
def unflag_email(message_id):
    """Remove a flag entirely, clearing the email from Outlook's flagged/To-Do view."""
    mapi = get_mapi()
    msg = _get_item(mapi, message_id)
    msg.FlagStatus = 0
    msg.PropertyAccessor.SetProperty(PR_TODO_ITEM_FLAGS, 0)
    msg.Save()
    print(json.dumps({"status": "unflagged", "message_id": message_id, "subject": msg.Subject}))


@cli.command("cleanup")
def cleanup_stale_flags():
    """Remove stale completed flags from Outlook's To-Do view in one pass.

    Finds all items in the To-Do folder with FlagStatus=2 (complete) but no
    active PR_TODO bit, and fully clears their flag state. These are emails that
    were previously marked complete but remain visible in Outlook's flagged view
    because FlagStatus=2 does not remove items from the To-Do list — only clearing
    all flag state does.
    """
    mapi = get_mapi()
    inbox = _find_folder(mapi, "inbox")
    items = inbox.Items
    _dbg(f"cleanup: scanning {items.Count} inbox items for stale FlagStatus=2")

    # Pass 1: scan inbox with GetFirst/GetNext. FlagStatus is a native COM attribute
    # (~1ms/item), so 3200 items takes ~3s. Collect refs without modifying.
    stale = []
    msg = items.GetFirst()
    while msg:
        try:
            if getattr(msg, "Class", None) == 43 and getattr(msg, "FlagStatus", 0) == 2:
                try:
                    tv = msg.PropertyAccessor.GetProperty(PR_TODO_ITEM_FLAGS)
                    todo_active = bool(tv and (tv & 1))
                except Exception:
                    todo_active = False
                if not todo_active:
                    stale.append(msg)
        except Exception:
            pass
        msg = items.GetNext()
    _dbg(f"cleanup: found {len(stale)} stale items")

    # Pass 2: modify after iteration is complete — no collection mutation mid-loop.
    cleared = []
    for item in stale:
        try:
            subj = item.Subject or "(no subject)"
            item.FlagStatus = 0
            item.PropertyAccessor.SetProperty(PR_TODO_ITEM_FLAGS, 0)
            item.Save()
            cleared.append(subj)
            _dbg(f"cleanup: cleared '{subj}'")
        except Exception as e:
            _dbg(f"cleanup: failed ({e})")

    _dbg(f"cleanup: done, cleared {len(cleared)} items")
    print(json.dumps({"status": "ok", "cleared": len(cleared), "subjects": cleared}, indent=2))


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
