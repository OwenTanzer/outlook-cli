#!/usr/bin/env python3
"""
Verification script: compare Sort()+GetFirst/GetNext against scan-everything-sort-in-Python.
Run before committing the list performance fix to confirm Sort() produces correct ordering.
"""

import sys
import time
import win32com.client


def get_inbox():
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    for store in mapi.Folders:
        for sub in store.Folders:
            if sub.Name.lower().replace(" ", "") == "inbox":
                return sub
    raise RuntimeError("Inbox not found")


def method_a_scan_all(inbox, n):
    """Current approach: iterate everything, sort in Python."""
    t0 = time.perf_counter()
    collected = []
    for msg in inbox.Items:
        try:
            if getattr(msg, "Class", None) != 43:
                continue
            collected.append((msg.ReceivedTime.isoformat(), msg.EntryID, msg.Subject))
        except Exception:
            continue
    collected.sort(key=lambda x: x[0], reverse=True)
    elapsed = time.perf_counter() - t0
    return collected[:n], elapsed


def method_b_sort_getnext(inbox, n):
    """Proposed approach: Sort() + GetFirst/GetNext, stop at n."""
    t0 = time.perf_counter()
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    collected = []
    msg = items.GetFirst()
    while msg and len(collected) < n:
        try:
            if getattr(msg, "Class", None) == 43:
                collected.append((msg.ReceivedTime.isoformat(), msg.EntryID, msg.Subject))
        except Exception:
            pass
        msg = items.GetNext()
    elapsed = time.perf_counter() - t0
    return collected, elapsed


def main():
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 20

    print(f"Comparing methods for top {n} emails...\n")
    inbox = get_inbox()

    print("Method A: scan all + sort in Python")
    results_a, time_a = method_a_scan_all(inbox, n)
    print(f"  {len(results_a)} results in {time_a:.2f}s\n")

    print("Method B: Sort() + GetFirst/GetNext")
    results_b, time_b = method_b_sort_getnext(inbox, n)
    print(f"  {len(results_b)} results in {time_b:.2f}s\n")

    print(f"Speedup: {time_a / time_b:.1f}x\n")

    # Compare
    mismatches = 0
    print(f"{'#':<4} {'Match':<6} {'Method A date':<22} {'Method B date':<22} Subject")
    print("-" * 100)
    for i, (a, b) in enumerate(zip(results_a, results_b)):
        date_a, id_a, subj_a = a
        date_b, id_b, subj_b = b
        match = "OK" if id_a == id_b else "DIFF"
        if match == "DIFF":
            mismatches += 1
        print(f"{i+1:<4} {match:<6} {date_a[:19]:<22} {date_b[:19]:<22} {subj_a[:40]}")
        if match == "DIFF":
            print(f"     {'':6} A: {subj_a[:60]}")
            print(f"     {'':6} B: {subj_b[:60]}")

    print()
    if mismatches == 0:
        print(f"PASS — all {n} results match. Sort() ordering is consistent.")
    else:
        print(f"FAIL — {mismatches}/{n} mismatches. Do NOT use Sort()+GetFirst/GetNext.")


if __name__ == "__main__":
    main()
