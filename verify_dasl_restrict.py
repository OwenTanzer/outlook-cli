#!/usr/bin/env python3
"""
Verify DASL Restrict by exact message_id against the full _get_item scan.

Uses verify_ids.json (written by verify_getitemfromid.py collect) as ground truth.
Tests two filter formats (message_id with and without angle brackets), times both
approaches, and probes fallback behavior with a bad ID.

Run in a fresh process — does not depend on any prior session state:
    python verify_dasl_restrict.py [--n N]
"""

import sys
import json
import time
import argparse
import win32com.client

VERIFY_FILE = "verify_ids.json"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"


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


def restrict_lookup(inbox_items, message_id, bracket_format):
    """Try Restrict with the given format. Returns (item_or_None, elapsed_s, error_or_None)."""
    if bracket_format == "with_brackets":
        value = f"<{message_id}>"
    else:
        value = message_id
    # DASL: property URL in double quotes, value in single quotes
    filter_str = f"@SQL=\"{PR_INTERNET_MESSAGE_ID}\" = '{value}'"
    t0 = time.perf_counter()
    try:
        restricted = inbox_items.Restrict(filter_str)
        item = restricted.GetFirst()
        elapsed = time.perf_counter() - t0
        return item, elapsed, None
    except Exception as e:
        return None, time.perf_counter() - t0, str(e)


def scan_lookup(mapi, message_id):
    """Current _get_item O(n) scan. Returns (item_or_None, elapsed_s)."""
    t0 = time.perf_counter()
    for store in mapi.Folders:
        for folder in store.Folders:
            try:
                for item in folder.Items:
                    try:
                        if internet_message_id(item) == message_id:
                            return item, time.perf_counter() - t0
                    except Exception:
                        pass
            except Exception:
                pass
    return None, time.perf_counter() - t0


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--n", type=int, default=None,
                        help="Number of records to test (default: all in verify_ids.json)")
    parser.add_argument("--scan-for", type=int, default=3,
                        help="How many items to also time via full scan for comparison (default: 3)")
    args = parser.parse_args()

    try:
        with open(VERIFY_FILE) as f:
            records = json.load(f)
    except FileNotFoundError:
        print(f"ERROR: {VERIFY_FILE} not found. Run verify_getitemfromid.py collect first.")
        sys.exit(1)

    if args.n:
        records = records[:args.n]

    mapi = get_mapi()
    inbox = get_inbox(mapi)
    inbox_items = inbox.Items  # hold reference — avoids re-fetching each call

    # --- Probe which bracket format works ---
    print("Probing bracket formats on first record...\n")
    probe = records[0]
    for fmt in ("with_brackets", "without_brackets"):
        item, elapsed, err = restrict_lookup(inbox_items, probe["message_id"], fmt)
        if err:
            result = f"ERROR: {err}"
        elif item is None:
            result = "no item returned (Restrict returned empty)"
        else:
            got = internet_message_id(item)
            match = "MATCH" if got == probe["message_id"] else f"MISMATCH (got {got})"
            result = f"{match}  ({elapsed*1000:.1f}ms)"
        print(f"  Format '{fmt}': {result}")

    # Determine which format to use for the main run
    working_fmt = None
    for fmt in ("with_brackets", "without_brackets"):
        item, _, err = restrict_lookup(inbox_items, probe["message_id"], fmt)
        if item and internet_message_id(item) == probe["message_id"]:
            working_fmt = fmt
            break

    print()
    if working_fmt is None:
        print("NEITHER format returned correct results — DASL Restrict is not viable for this property.")
        print("Stick with _get_item scan or use GetItemFromID from verify_getitemfromid.py.")
        sys.exit(1)

    print(f"Using format: '{working_fmt}' for full verification run.\n")

    # --- Full verification run ---
    print(f"{'#':<4} {'Label':<14} {'Result':<8} {'Time':<10} {'Subject':<45}")
    print("-" * 90)

    pass_count = 0
    fail_count = 0
    error_count = 0
    restrict_times = []

    for i, rec in enumerate(records):
        message_id = rec["message_id"]
        subject = rec["subject"][:43]
        label = rec.get("label", f"item{i}")

        item, elapsed, err = restrict_lookup(inbox_items, message_id, working_fmt)
        restrict_times.append(elapsed)

        if err:
            status = "ERROR"
            error_count += 1
            detail = err
        elif item is None:
            status = "MISS"
            fail_count += 1
            detail = "Restrict returned no items"
        else:
            got = internet_message_id(item)
            if got == message_id:
                status = "OK"
                pass_count += 1
                detail = None
            else:
                status = "DIFF"
                fail_count += 1
                detail = f"expected {message_id}, got {got}"

        time_str = f"{elapsed*1000:.1f}ms"
        print(f"{i+1:<4} {label:<14} {status:<8} {time_str:<10} {subject}")
        if detail:
            print(f"     {detail}")

    # --- Fallback probe: bad ID should return None, not raise ---
    print()
    print("Probing fallback with deliberately bad message_id...")
    bad_item, bad_elapsed, bad_err = restrict_lookup(
        inbox_items, "definitely-not-a-real-id@nowhere.invalid", working_fmt
    )
    if bad_err:
        print(f"  ERROR raised: {bad_err}  (fallback to scan would be needed)")
    elif bad_item is None:
        print(f"  Returned None cleanly in {bad_elapsed*1000:.1f}ms — fallback path is safe")
    else:
        print(f"  Unexpectedly returned an item: {bad_item.Subject}")

    # --- Scan timing for comparison ---
    print()
    scan_subset = records[:args.scan_for]
    scan_times = []
    if scan_subset:
        print(f"Timing full _get_item scan for {len(scan_subset)} items (comparison)...")
        for rec in scan_subset:
            item, elapsed = scan_lookup(mapi, rec["message_id"])
            scan_times.append(elapsed)
            found = "found" if item else "NOT FOUND"
            print(f"  scan {rec['label']}: {elapsed*1000:.0f}ms ({found})")

    # --- Summary ---
    print()
    print("=" * 90)
    avg_restrict = sum(restrict_times) / len(restrict_times) * 1000
    print(f"DASL Restrict: {pass_count} pass, {fail_count} miss/diff, {error_count} error  "
          f"(avg {avg_restrict:.1f}ms per call)")
    if scan_times:
        avg_scan = sum(scan_times) / len(scan_times)
        speedup = avg_scan / (sum(restrict_times) / len(restrict_times))
        print(f"Full scan:     avg {avg_scan*1000:.0f}ms per call  (~{speedup:.0f}x slower than Restrict)")

    total = pass_count + fail_count + error_count
    print()
    if fail_count == 0 and error_count == 0:
        print(f"PASS — all {total} Restrict lookups returned correct items.")
        print("DASL Restrict by message_id is viable — no persistent ID storage needed.")
    else:
        print(f"FAIL — {fail_count + error_count}/{total} lookups did not return correct items.")
        print("Fall back to GetItemFromID or full scan.")


if __name__ == "__main__":
    main()
