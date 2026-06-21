#!/usr/bin/env python3
"""
Verify DASL Restrict filters for Fix 1: flagged email detection.

Builds ground truth using the current PropertyAccessor-based logic (known correct —
matches Outlook's native flagged view), then compares each candidate Restrict filter
against that set by message_id.

Tests four unknowns introduced by Fix 1 that were not covered by Fix 3:
  1. Integer comparison in DASL (vs string comparison tested in Fix 3)
  2. Bitwise semantics: does = 1 cover all cases our (flags & 1) check covers?
  3. OR logic combining two DASL conditions
  4. Whether FlagStatus is reachable via named property [FlagStatus] or needs DASL URL

Run in a fresh process:
    python verify_dasl_flagged.py
"""

import sys
import time
import win32com.client

PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_TODO_ITEM_FLAGS     = "http://schemas.microsoft.com/mapi/proptag/0x0E2B0003"
PR_FLAG_STATUS         = "http://schemas.microsoft.com/mapi/proptag/0x10900003"

# Candidate filters to test — each is a (label, filter_string) pair.
# We probe several variants to identify exactly which syntax works and
# whether any combination correctly covers all 37 flagged emails.
FILTERS = [
    (
        "PR_TODO only (int)",
        f'@SQL="{PR_TODO_ITEM_FLAGS}" = 1',
    ),
    (
        "FlagStatus named prop",
        '[FlagStatus] = 1',
    ),
    (
        "FlagStatus DASL URL (int)",
        f'@SQL="{PR_FLAG_STATUS}" = 1',
    ),
    (
        "OR: PR_TODO | FlagStatus named",
        f'@SQL="{PR_TODO_ITEM_FLAGS}" = 1 OR [FlagStatus] = 1',
    ),
    (
        "OR: PR_TODO | FlagStatus DASL (full DASL)",
        f'@SQL=("{PR_TODO_ITEM_FLAGS}" = 1 OR "{PR_FLAG_STATUS}" = 1)',
    ),
]


def get_mapi():
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


def get_inbox(mapi):
    for store in mapi.Folders:
        for sub in store.Folders:
            if sub.Name.lower().replace(" ", "") == "inbox":
                return sub
    raise RuntimeError("Inbox not found")


def internet_message_id(msg):
    try:
        mid = msg.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
        return mid.strip("<>") if mid else None
    except Exception:
        return None


def todo_flagged(msg):
    try:
        flags = msg.PropertyAccessor.GetProperty(PR_TODO_ITEM_FLAGS)
        return bool(flags and (flags & 1))
    except Exception:
        return False


def build_ground_truth(inbox):
    """Replicate current _flag_status logic to get the verified-correct set."""
    print("Building ground truth via PropertyAccessor scan (current correct logic)...")
    t0 = time.perf_counter()
    truth = {}
    for msg in inbox.Items:
        try:
            if getattr(msg, "Class", None) != 43:
                continue
            # Mirror _flag_status: PR_TODO first, then classic FlagStatus
            if todo_flagged(msg):
                flagged = True
            elif getattr(msg, "FlagStatus", 0) == 1:
                flagged = True
            else:
                flagged = False
            if flagged:
                mid = internet_message_id(msg)
                if mid:
                    truth[mid] = msg.Subject or "(no subject)"
        except Exception:
            continue
    elapsed = time.perf_counter() - t0
    print(f"  {len(truth)} flagged emails found in {elapsed:.1f}s\n")
    return truth


def run_filter(inbox_items, label, filter_str):
    """Apply Restrict filter, collect all matching message_ids."""
    t0 = time.perf_counter()
    results = {}
    try:
        restricted = inbox_items.Restrict(filter_str)
        msg = restricted.GetFirst()
        while msg:
            try:
                mid = internet_message_id(msg)
                if mid:
                    results[mid] = msg.Subject or "(no subject)"
            except Exception:
                pass
            msg = restricted.GetNext()
    except Exception as e:
        return None, 0, str(e)
    elapsed = time.perf_counter() - t0
    return results, elapsed, None


def compare(truth, results, label):
    """Compare result set to ground truth; return whether it's a full match."""
    missing  = {mid: subj for mid, subj in truth.items()   if mid not in results}
    extra    = {mid: subj for mid, subj in results.items() if mid not in truth}
    match    = not missing and not extra

    status = "PASS" if match else "FAIL"
    print(f"  {status}  |  truth={len(truth)}  filter={len(results)}  "
          f"missing={len(missing)}  extra={len(extra)}")

    if missing:
        print(f"  In ground truth but MISSING from filter ({len(missing)}):")
        for mid, subj in list(missing.items())[:5]:
            print(f"    - {subj[:60]}")
        if len(missing) > 5:
            print(f"    ... and {len(missing)-5} more")

    if extra:
        print(f"  In filter but NOT in ground truth ({len(extra)}) — false positives:")
        for mid, subj in list(extra.items())[:5]:
            print(f"    + {subj[:60]}")

    return match


def main():
    mapi   = get_mapi()
    inbox  = get_inbox(mapi)
    truth  = build_ground_truth(inbox)

    inbox_items = inbox.Items  # hold reference across all filter calls

    winners = []
    for label, filter_str in FILTERS:
        print(f"Filter: {label}")
        print(f"  Query: {filter_str}")
        results, elapsed, err = run_filter(inbox_items, label, filter_str)

        if err:
            print(f"  ERROR: {err}")
            print()
            continue

        print(f"  Time:  {elapsed:.2f}s")
        ok = compare(truth, results, label)
        if ok:
            winners.append((label, filter_str, elapsed))
        print()

    print("=" * 70)
    if winners:
        print(f"PASSING filters ({len(winners)}/{len(FILTERS)}):")
        for label, fstr, elapsed in winners:
            print(f"  [{elapsed:.2f}s]  {label}")
            print(f"           {fstr}")
        print()
        # Recommend the simplest passing filter
        print("Recommended for implementation: first passing filter above.")
        print("If 'PR_TODO only' passes, prefer it — simpler, no OR needed.")
        print("If not, use the combined OR filter that passes.")
    else:
        print("NO filters passed. Do not implement Fix 1 without further investigation.")


if __name__ == "__main__":
    main()
