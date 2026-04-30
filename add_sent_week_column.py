"""add_sent_week_column.py — Pre-flight validation for the weekly snapshot.

The Sent Week column lives at col O (0-indexed 14) on the 3 cumulative tabs:
    _ApprovedData, _AdditionsData, _UnenrollData

_RejectedData is intentionally skipped — rejected rows don't go to support.

Because cumulative tabs don't have a header row (Apps Script appendRow starts
writing at row 1), there's no header to add — col O simply holds data per row
(blank string = unsent, "YYYY-MM-DD" = sent in that week). This script is
therefore a sanity check rather than a data-mutating migration:

  - Confirms all 3 target tabs exist
  - Reports row count and Sent Week state breakdown per tab
  - Warns if col O contains unexpected non-ISO-date values (potential prior test data)

Idempotent. Safe to run any time. No writes.

Usage:
    python add_sent_week_column.py
"""

import re
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build

from config import (
    SERVICE_ACCOUNT_KEY,
    SCOPES,
    OUTPUT_SPREADSHEET_ID,
    SENT_WEEK_COL,
    SENT_WEEK_HEADER,
    WEEKLY_SOURCE_TABS,
)
from retry_helper import retry_api  # v2.5.2

ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _col_letter(idx):
    """0-based column index → A1 letter. Handles up to ZZ."""
    if idx < 26:
        return chr(ord("A") + idx)
    return chr(ord("A") + idx // 26 - 1) + chr(ord("A") + idx % 26)


def main():
    print("=" * 60)
    print("  PRE-FLIGHT — Sent Week column state on cumulative tabs")
    print("=" * 60)

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEY, scopes=SCOPES
    )
    sheets = build("sheets", "v4", credentials=creds)

    # Confirm all tabs exist
    ss_meta = retry_api(
        lambda: sheets.spreadsheets()
        .get(spreadsheetId=OUTPUT_SPREADSHEET_ID, fields="sheets.properties.title")
        .execute(),
        label="get spreadsheet metadata",
    )
    tab_names = {s["properties"]["title"] for s in ss_meta.get("sheets", [])}

    col_letter = _col_letter(SENT_WEEK_COL)  # "O"
    target_tabs = sorted(set(WEEKLY_SOURCE_TABS.values()))

    missing = [t for t in target_tabs if t not in tab_names]
    if missing:
        print(f"\n  ERROR: missing target tabs: {missing}")
        print("         Run generate_corrections.py first to create them.")
        sys.exit(1)

    total_unsent = 0
    total_sent = 0
    total_bad = 0

    for tab in target_tabs:
        # Read A:O so we see actual data rows, not just col O (which may be empty)
        resp = retry_api(
            lambda t=tab: sheets.spreadsheets()
            .values()
            .get(
                spreadsheetId=OUTPUT_SPREADSHEET_ID,
                range=f"'{t}'!A:{col_letter}",
            )
            .execute(),
            label=f"read '{tab}' col O",
        )
        rows = resp.get("values", [])
        unsent = 0
        sent = 0
        bad = 0
        distinct_weeks = set()
        for row in rows:
            # Pad to at least col O
            if len(row) <= SENT_WEEK_COL:
                cell = ""
            else:
                cell = str(row[SENT_WEEK_COL] or "").strip()
            if not cell:
                unsent += 1
            elif ISO_DATE_RE.match(cell):
                sent += 1
                distinct_weeks.add(cell)
            else:
                bad += 1
        print(f"\n  {tab}")
        print(f"    data rows         : {len(rows)}")
        print(f"    Sent Week blank   : {unsent}")
        print(f"    Sent Week valid   : {sent}")
        if distinct_weeks:
            print(f"    distinct weeks    : {sorted(distinct_weeks)}")
        if bad:
            print(f"    [WARN] non-ISO-date values in col {col_letter}: {bad}")
        total_unsent += unsent
        total_sent += sent
        total_bad += bad

    print()
    print(f"  Total unsent rows across 3 tabs: {total_unsent}")
    print(f"  Total sent rows (any week): {total_sent}")
    if total_bad:
        print(
            f"  [WARN] {total_bad} row(s) have non-ISO-date values in col "
            f"{col_letter} — inspect before running weekly snapshot."
        )
    print(
        f"\n  Constants: SENT_WEEK_COL={SENT_WEEK_COL} ('{col_letter}'), "
        f"SENT_WEEK_HEADER='{SENT_WEEK_HEADER}'"
    )
    print("  Pre-flight complete. No writes performed.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n  FATAL: {e}")
        sys.exit(1)
