"""normalize_dates.py — One-time normalization of date strings in cumulative tabs.

Problem: Apps Script's onEdit appendRow + setNumberFormat races when multiple
edits fire nearly simultaneously. Some rows end up with locale-default format
(e.g. "4/23/2026 1:37:44") while others get the intended ISO format
("2026-04-23 01:37:47"). The inconsistency causes weird lexicographic sorts.

Fix: walk each cumulative tab (_ApprovedData, _AdditionsData, _UnenrollData,
_RejectedData), parse column A (the Date column) regardless of current format,
and rewrite it as a canonical "yyyy-MM-dd HH:mm:ss" string. Also apply the
same number format at the column level so future writes display consistently.

Idempotent — safe to re-run.
"""

import re
import sys
import time
from datetime import datetime

from google.oauth2 import service_account
from googleapiclient.discovery import build

from config import SERVICE_ACCOUNT_KEY, SCOPES, OUTPUT_SPREADSHEET_ID
from retry_helper import retry_api  # v2.5.2

CUMULATIVE_TABS = [
    "_ApprovedData",
    "_AdditionsData",
    "_UnenrollData",
    "_RejectedData",
]

CANONICAL_FMT = "%Y-%m-%d %H:%M:%S"

# Parse patterns, ordered most-specific-first. Add more if new formats appear.
_PARSERS = [
    ("%Y-%m-%d %H:%M:%S", re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")),
    ("%Y-%m-%d %H:%M", re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$")),
    ("%m/%d/%Y %H:%M:%S", re.compile(r"^\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2}:\d{2}$")),
    ("%m/%d/%Y %H:%M", re.compile(r"^\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2}$")),
    ("%Y-%m-%d", re.compile(r"^\d{4}-\d{2}-\d{2}$")),
    ("%m/%d/%Y", re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")),
]


def parse_date(value):
    """Parse any of the known date string formats into a datetime. Returns None on failure."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    # Try all known patterns
    for fmt, pattern in _PARSERS:
        if pattern.match(s):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
    return None


def get_sheet_id(sheets, spreadsheet_id, tab_name):
    resp = retry_api(
        lambda: sheets.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
        .execute(),
        label="get sheet metadata",
    )
    for s in resp.get("sheets", []):
        if s["properties"]["title"] == tab_name:
            return s["properties"]["sheetId"]
    return None


def normalize_tab(sheets, spreadsheet_id, tab_name):
    # Read the Date column (A) with FORMATTED_VALUE so we see how the user sees it.
    resp = retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=f"'{tab_name}'!A1:A")
        .execute(),
        label=f"read '{tab_name}' col A",
    )
    rows = resp.get("values", [])
    if not rows:
        print(f"  {tab_name}: empty, skipping")
        return

    normalized = []
    unparseable = 0
    mutated = 0
    for row in rows:
        if not row:
            normalized.append([""])
            continue
        val = row[0]
        dt = parse_date(val)
        if dt is None:
            # Leave as-is (don't corrupt unknown values)
            normalized.append([val])
            if val:
                unparseable += 1
            continue
        new_val = dt.strftime(CANONICAL_FMT)
        normalized.append([new_val])
        if str(val).strip() != new_val:
            mutated += 1

    if mutated == 0:
        print(f"  {tab_name}: already normalized ({len(rows)} rows)")
        return

    print(
        f"  {tab_name}: normalizing {mutated}/{len(rows)} rows ({unparseable} unparseable)"
    )

    # Write back column A. Use RAW so Google doesn't try to re-interpret the string.
    retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab_name}'!A1",
            valueInputOption="RAW",
            body={"values": normalized},
        )
        .execute(),
        label=f"write normalized dates to '{tab_name}'",
    )

    # Apply consistent number format at the cell level. Since the values are now
    # strings (not serial numbers), the format doesn't reinterpret them — it just
    # keeps display consistent if Google later tries to auto-parse.
    sheet_id = get_sheet_id(sheets, spreadsheet_id, tab_name)
    if sheet_id is not None:
        retry_api(
            lambda: sheets.spreadsheets()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "requests": [
                        {
                            "repeatCell": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": 0,
                                    "endRowIndex": len(rows) + 100,
                                    "startColumnIndex": 0,
                                    "endColumnIndex": 1,
                                },
                                "cell": {
                                    "userEnteredFormat": {
                                        "numberFormat": {
                                            "type": "TEXT",
                                        }
                                    }
                                },
                                "fields": "userEnteredFormat.numberFormat",
                            }
                        }
                    ]
                },
            )
            .execute(),
            label=f"apply TEXT format to '{tab_name}'",
        )


def main():
    start = time.time()
    print("=" * 70)
    print("  NORMALIZE DATES IN CUMULATIVE TABS")
    print("=" * 70)

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEY, scopes=SCOPES
    )
    sheets = build("sheets", "v4", credentials=creds)

    for tab in CUMULATIVE_TABS:
        normalize_tab(sheets, OUTPUT_SPREADSHEET_ID, tab)

    print(f"\n  Completed in {time.time() - start:.1f}s")


if __name__ == "__main__":
    main()
