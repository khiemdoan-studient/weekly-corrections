"""generate_weekly_snapshot.py — Create/update a weekly Google Sheet that
bundles all corrections not yet sent to support.

Runs every Monday at 07:00 ET via GitHub Actions, and can be re-run manually
any day. When re-run in the same week, it updates the existing weekly file
in place (same file ID — any existing shares with support remain valid).

Behavior
--------
1. Compute current Monday in America/New_York.
2. Read each of the 3 source cumulative tabs (_ApprovedData,
   _AdditionsData, _UnenrollData) and filter for rows where col O
   (Sent Week) is blank OR equals the current Monday ISO date.
3. If 0 unsent rows total → skip file creation, exit cleanly. Drive stays
   clean for weeks where no IM accepted a correction. (v2.5.1: this was
   the v2.5.0 bug — empty file → all 3 tabs hidden → deleteSheet on the
   default Sheet1 → API error "can't remove all visible sheets".)
4. Otherwise, find-or-create "M/D Corrections" spreadsheet inside the
   WEEKLY_SHARED_DRIVE_ID Shared Drive. Files in a Shared Drive are owned
   by the drive itself, so no per-user quota applies and anyone with drive
   membership automatically has access — no per-file sharing needed.
5. Write the selected rows into the weekly sheet under tabs:
       Correction List        ← _ApprovedData
       Roster Additions       ← _AdditionsData
       Roster Unenrollments   ← _UnenrollData
   Hide any tab that has 0 data rows. Delete the default "Sheet1" tab
   created with the spreadsheet.
6. Stamp col O of the selected rows in the cumulative tabs with the
   current Monday ISO (e.g. "2026-04-20"), so next week's run excludes
   them automatically.

Design notes
------------
- Uses a Shared Drive (not a user's personal Drive) because the service
  account has 0 bytes of storage quota and cannot own regular Drive files.
  The SA is added as Content Manager to the Shared Drive; files created
  there are owned by the drive itself.
- _RejectedData is NOT included (rejected rows don't go to support).
- Cumulative tabs have no header row — col O simply holds data per row
  ("" = unsent, "YYYY-MM-DD" = sent in that week).
- Re-running the same week picks up both already-sent rows (Sent Week ==
  current Monday) AND newly-accepted unsent rows, so the snapshot always
  reflects the full week up to now.

Usage
-----
    python generate_weekly_snapshot.py
"""

import argparse
import re
import sys
import time
from datetime import datetime, timedelta, date
from zoneinfo import ZoneInfo

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from config import (
    SERVICE_ACCOUNT_KEY,
    SCOPES,
    OUTPUT_SPREADSHEET_ID,
    WEEKLY_SHARED_DRIVE_ID,
    WEEKLY_SHARED_DRIVE_NAME,
    WEEKLY_TIMEZONE,
    SENT_WEEK_COL,
    WEEKLY_SOURCE_TABS,
    WEEKLY_HEADERS,
    WEEKLY_TAB_INSTRUCTIONS,
)

ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

SHEET_MIME = "application/vnd.google-apps.spreadsheet"


# ── Helpers ────────────────────────────────────────────────────────────────


def _col_letter(idx):
    """0-based column index → A1 letter. Handles up to ZZ."""
    if idx < 26:
        return chr(ord("A") + idx)
    return chr(ord("A") + idx // 26 - 1) + chr(ord("A") + idx % 26)


def compute_monday(tz_name):
    """Return (date, 'M/D', 'YYYY-MM-DD') for Monday of the current week in tz."""
    now = datetime.now(ZoneInfo(tz_name))
    monday = (now - timedelta(days=now.weekday())).date()
    return monday, f"{monday.month}/{monday.day}", monday.isoformat()


# v2.5.2: retry helper now imported from retry_helper.py — was 3-attempts-
# linear-backoff with HttpError-only catch, which missed TimeoutError and
# couldn't span sustained transient outages. New helper does 5 attempts
# exponential + jitter and catches HttpError 5xx/429/408, TimeoutError,
# socket.timeout, and ConnectionError.
from retry_helper import retry_api as _retry  # noqa: E402

# ── Drive operations (Shared Drive mode) ───────────────────────────────────
# All Drive API calls need supportsAllDrives=True and files.list additionally
# needs driveId + corpora='drive' + includeItemsFromAllDrives=True to
# scope to our Shared Drive.


def find_sheet_in_shared_drive(drive, shared_drive_id, name):
    """Return spreadsheet ID if a sheet with this name exists in the drive, else None."""
    safe_name = name.replace("'", "\\'")
    q = f"name='{safe_name}' and mimeType='{SHEET_MIME}' and trashed=false"
    resp = _retry(
        lambda: drive.files()
        .list(
            q=q,
            driveId=shared_drive_id,
            corpora="drive",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            fields="files(id,name)",
        )
        .execute()
    )
    files = resp.get("files", [])
    return files[0]["id"] if files else None


def create_sheet_in_shared_drive(drive, shared_drive_id, name):
    """Create a new Google Sheets file at the root of the Shared Drive."""
    result = _retry(
        lambda: drive.files()
        .create(
            body={
                "name": name,
                "mimeType": SHEET_MIME,
                "parents": [shared_drive_id],  # Shared Drive ID acts as parent
            },
            supportsAllDrives=True,
            fields="id",
        )
        .execute()
    )
    return result["id"]


# ── Sheets operations ──────────────────────────────────────────────────────


def get_sheet_tabs(sheets, spreadsheet_id):
    """Return dict of tab_name → sheetId."""
    resp = _retry(
        lambda: sheets.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
        .execute()
    )
    return {
        s["properties"]["title"]: s["properties"]["sheetId"]
        for s in resp.get("sheets", [])
    }


def ensure_tab(sheets, spreadsheet_id, name, existing_tabs):
    """Ensure a tab with this name exists. Returns sheetId. Updates existing_tabs in place."""
    if name in existing_tabs:
        return existing_tabs[name]
    result = _retry(
        lambda: sheets.spreadsheets()
        .batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": name}}}]},
        )
        .execute()
    )
    new_id = result["replies"][0]["addSheet"]["properties"]["sheetId"]
    existing_tabs[name] = new_id
    return new_id


def read_cumulative_tab(sheets, spreadsheet_id, tab_name, sent_week_col):
    """Read all rows from a cumulative tab.

    Returns list of (row_index_1based, row_values_padded_to_15_cols).
    """
    col_letter = _col_letter(sent_week_col)
    resp = (
        sheets.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab_name}'!A:{col_letter}",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        )
        .execute()
    )
    rows = resp.get("values", [])
    result = []
    for i, row in enumerate(rows):
        padded = list(row) + [""] * (sent_week_col + 1 - len(row))
        result.append((i + 1, padded))  # 1-based row
    return result


def filter_for_week(rows, current_monday_iso, sent_week_col, all_unsent=False):
    """Return subset of rows whose Sent Week (col O) qualifies.

    Default mode: blank OR == current_monday_iso. Lets within-week re-runs
    pick up the rows they already stamped (idempotent) plus any newly-blank
    ones added since the last run.

    all_unsent=True (v2.6.1 support-packet mode): blank only. Bundles every
    correction that has never been sent to support, regardless of week.
    Used by `python generate_weekly_snapshot.py --all-unsent` for ad-hoc
    support packets.
    """
    selected = []
    for row_num, row in rows:
        sent = str(row[sent_week_col] or "").strip()
        if all_unsent:
            include = not sent
        else:
            include = (not sent) or (sent == current_monday_iso)
        if include:
            # Also skip truly empty rows (no Date, no Campus)
            if str(row[0] or "").strip() or str(row[2] or "").strip():
                selected.append((row_num, row))
    return selected


# ── Formatting ─────────────────────────────────────────────────────────────


def _rgb(h):
    h = h.lstrip("#")
    return {
        "red": int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue": int(h[4:6], 16) / 255,
    }


NAVY_DARK = _rgb("0F1B33")
WHITE = _rgb("FFFFFF")
ALT_ROW = _rgb("EDF2F7")


def build_tab_format_requests(sheet_id, num_cols, num_data_rows):
    """Header bold + freeze + banding + column widths for a weekly tab."""
    requests = [
        # Header row: navy bg, white bold 10pt
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": num_cols,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": NAVY_DARK,
                        "textFormat": {
                            "foregroundColor": WHITE,
                            "bold": True,
                            "fontSize": 10,
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        # Freeze header
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 1},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        },
    ]
    # Banding — but only if there's data
    if num_data_rows > 0:
        end_row = max(1 + num_data_rows, 2)
        requests.append(
            {
                "addBanding": {
                    "bandedRange": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": end_row,
                            "startColumnIndex": 0,
                            "endColumnIndex": num_cols,
                        },
                        "rowProperties": {
                            "headerColor": NAVY_DARK,
                            "firstBandColor": _rgb("FFFFFF"),
                            "secondBandColor": ALT_ROW,
                        },
                    }
                }
            }
        )

    # Column widths — Date, Mismatch, then 12 fields
    widths = [130, 180, 150, 60, 80, 110, 110, 220, 150, 110, 110, 220, 110, 130]
    for i, w in enumerate(widths[:num_cols]):
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": i,
                        "endIndex": i + 1,
                    },
                    "properties": {"pixelSize": w},
                    "fields": "pixelSize",
                }
            }
        )
    return requests


# ── Instructions tab (v2.6.1 support packet) ──────────────────────────────


def _build_instructions_rows(generation_iso, total_rows, per_tab_counts):
    """Build the (text, style) tuple list for the Instructions tab.

    Style is one of: "title", "h2", "h3", "body", "blank". The
    `_instructions_format_requests` helper maps each style to per-row
    formatting (bold/size/background).
    """
    corrections_n = per_tab_counts.get("Correction List", 0)
    additions_n = per_tab_counts.get("Roster Additions", 0)
    unenrolls_n = per_tab_counts.get("Roster Unenrollments", 0)

    rows = [
        ("Studient — Roster Correction Packet", "title"),
        (f"Generated {generation_iso} (America/New_York)", "body"),
        ("", "blank"),
        ("What this is", "h2"),
        (
            "This file is a bundle of every roster correction the Studient team has "
            "approved that has NOT yet been processed by support. Each tab is a "
            "separate worklist. Please process every row — once processed, the row "
            "will not appear in the next packet.",
            "body",
        ),
        ("", "blank"),
        (
            f"Totals: {total_rows} row(s) across {sum(1 for c in (corrections_n, additions_n, unenrolls_n) if c > 0)} tab(s).",
            "body",
        ),
        (f"  - Correction List: {corrections_n} student(s) — field updates", "body"),
        (
            f"  - Roster Additions: {additions_n} student(s) — new students to add",
            "body",
        ),
        (
            f"  - Roster Unenrollments: {unenrolls_n} student(s) — students to remove",
            "body",
        ),
        ("", "blank"),
        ("How to use each tab", "h2"),
        ("", "blank"),
        ("Tab: Correction List", "h3"),
        (
            "Each row is a student whose record in our roster (MAP) does NOT match "
            "the SIS. Update the SIS so the values listed (First Name, Last Name, "
            "Email, Grade, Level, Guide, etc.) match exactly what is shown.",
            "body",
        ),
        (
            "Mismatch Summary (column B) names which fields differ. Use Student_ID "
            "(column M) or External Student ID (column N) to find the student in "
            "the SIS.",
            "body",
        ),
        ("", "blank"),
        ("Tab: Roster Additions", "h3"),
        (
            "Each row is a student who exists in MAP but NOT in the SIS. Add the "
            "student to the SIS using the values shown (Campus, Grade, Level, "
            "First/Last Name, Email, Guide, Student_ID, External Student ID).",
            "body",
        ),
        ("", "blank"),
        ("Tab: Roster Unenrollments", "h3"),
        (
            "Each row is a student whose MAP record was marked Unenroll = TRUE by "
            "an instructional manager. Mark these students as withdrawn / "
            "unenrolled in the SIS.",
            "body",
        ),
        ("", "blank"),
        ("Column reference", "h2"),
        (
            "A: Date Approved   B: Mismatch Summary   C: Campus   D: Grade   "
            "E: Level   F: First Name   G: Last Name   H: Email   I: Student Group   "
            "J: Guide First   K: Guide Last   L: Guide Email   M: Student_ID   "
            "N: External Student ID",
            "body",
        ),
        ("", "blank"),
        ("Questions / problems", "h2"),
        (
            "Reply to the Studient team that sent this packet. Do not edit this "
            "file — it is regenerated automatically and your edits will be "
            "overwritten on the next run.",
            "body",
        ),
    ]
    return rows


def _instructions_format_requests(sheet_id, rows):
    """Format the Instructions tab and pin it as the first tab (index 0).

    Per-row formatting applied based on style tag:
      - title: navy bg, white bold 16pt, row height 48
      - h2:    bold 13pt, row height 32
      - h3:    bold 11pt, row height 26
      - body:  10pt normal, row height auto (default)
      - blank: row height 14
    """
    requests = []

    # Move Instructions tab to index 0 (first tab when sheet opens)
    requests.append(
        {
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "index": 0},
                "fields": "index",
            }
        }
    )

    # Column A width = 900px so wrapped paragraphs are readable
    requests.append(
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 1,
                },
                "properties": {"pixelSize": 900},
                "fields": "pixelSize",
            }
        }
    )

    # Wrap text on column A for all rows
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": len(rows),
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "wrapStrategy": "WRAP",
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment",
            }
        }
    )

    # Per-row style
    for i, (_text, style) in enumerate(rows):
        if style == "title":
            requests.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": i,
                            "endRowIndex": i + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": NAVY_DARK,
                                "textFormat": {
                                    "foregroundColor": WHITE,
                                    "bold": True,
                                    "fontSize": 16,
                                },
                                "verticalAlignment": "MIDDLE",
                                "wrapStrategy": "WRAP",
                            }
                        },
                        "fields": "userEnteredFormat",
                    }
                }
            )
            requests.append(
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": i,
                            "endIndex": i + 1,
                        },
                        "properties": {"pixelSize": 48},
                        "fields": "pixelSize",
                    }
                }
            )
        elif style == "h2":
            requests.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": i,
                            "endRowIndex": i + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {"bold": True, "fontSize": 13},
                                "verticalAlignment": "MIDDLE",
                                "wrapStrategy": "WRAP",
                            }
                        },
                        "fields": "userEnteredFormat",
                    }
                }
            )
            requests.append(
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": i,
                            "endIndex": i + 1,
                        },
                        "properties": {"pixelSize": 32},
                        "fields": "pixelSize",
                    }
                }
            )
        elif style == "h3":
            requests.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": i,
                            "endRowIndex": i + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {"bold": True, "fontSize": 11},
                                "verticalAlignment": "MIDDLE",
                                "wrapStrategy": "WRAP",
                            }
                        },
                        "fields": "userEnteredFormat",
                    }
                }
            )
            requests.append(
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": i,
                            "endIndex": i + 1,
                        },
                        "properties": {"pixelSize": 26},
                        "fields": "pixelSize",
                    }
                }
            )
        elif style == "blank":
            requests.append(
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": i,
                            "endIndex": i + 1,
                        },
                        "properties": {"pixelSize": 14},
                        "fields": "pixelSize",
                    }
                }
            )
        # body: no extra formatting beyond the column-wide wrap+middle align

    # Hide all columns past A (this is a single-column instruction sheet)
    requests.append(
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": 26,
                },
                "properties": {"hiddenByUser": True},
                "fields": "hiddenByUser",
            }
        }
    )

    return requests


# ── Main ───────────────────────────────────────────────────────────────────


def main(all_unsent=False):
    start = time.time()
    print("=" * 70)
    if all_unsent:
        print("  WEEKLY SNAPSHOT — SUPPORT PACKET MODE (--all-unsent)")
    else:
        print("  WEEKLY SNAPSHOT — corrections bundle for support")
    print("=" * 70)

    monday_date, monday_label, monday_iso = compute_monday(WEEKLY_TIMEZONE)
    sheet_name = f"{monday_label} Corrections"
    print(f"  Current Monday ({WEEKLY_TIMEZONE}): {monday_iso}")
    print(f"  Target sheet name: '{sheet_name}'")
    if all_unsent:
        print(
            f"  Filter mode: ALL UNSENT (any row with blank Sent Week, "
            f"regardless of week)"
        )
    else:
        print(f"  Filter mode: this week (blank OR Sent Week == {monday_iso})")

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEY, scopes=SCOPES
    )
    drive = build("drive", "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)

    # ── Verify Shared Drive access ───────────────────────────────────
    print(
        f"\n  Target Shared Drive: '{WEEKLY_SHARED_DRIVE_NAME}' "
        f"id={WEEKLY_SHARED_DRIVE_ID}"
    )

    # ── Read cumulative tabs FIRST to know if there's anything to send ──
    # Done before any file create/find so we don't make an empty file when
    # there's nothing this week (the v2.5.0 bug — empty file → all 3 tabs
    # hidden → deleteSheet on Sheet1 → API error "can't remove all visible").
    print("\n  Reading cumulative tabs...")
    collected = {}  # weekly_tab_name → list of (orig_row_num, row_values)
    total_rows = 0
    for weekly_tab, source_tab in WEEKLY_SOURCE_TABS.items():
        rows = read_cumulative_tab(
            sheets, OUTPUT_SPREADSHEET_ID, source_tab, SENT_WEEK_COL
        )
        filtered = filter_for_week(
            rows, monday_iso, SENT_WEEK_COL, all_unsent=all_unsent
        )
        collected[weekly_tab] = filtered
        scope_label = "unsent (all-time)" if all_unsent else "selected for this week"
        print(
            f"    {source_tab} -> {weekly_tab}: "
            f"{len(rows)} total, {len(filtered)} {scope_label}"
        )
        total_rows += len(filtered)

    # ── Empty-week short-circuit ─────────────────────────────────────
    # If 0 unsent rows, don't create a file. Drive stays clean for weeks
    # where no IM accepted a correction. If a file already exists for this
    # week (e.g. an earlier run had rows but they've since been re-stamped,
    # or a leftover orphan from a failed prior run), leave it untouched —
    # manual cleanup if needed.
    if total_rows == 0:
        existing = find_sheet_in_shared_drive(drive, WEEKLY_SHARED_DRIVE_ID, sheet_name)
        elapsed = time.time() - start
        print()
        if existing is None:
            print(
                f"  No corrections to send this week. File not created. "
                f"({elapsed:.1f}s)"
            )
        else:
            print(f"  No new corrections this week. Existing file left untouched.")
            print(
                f"  URL: https://docs.google.com/spreadsheets/d/{existing}  "
                f"({elapsed:.1f}s)"
            )
        return

    # ── Find or create the weekly sheet inside the Shared Drive ──────
    ssid = find_sheet_in_shared_drive(drive, WEEKLY_SHARED_DRIVE_ID, sheet_name)
    created_new = False
    if ssid is None:
        ssid = create_sheet_in_shared_drive(drive, WEEKLY_SHARED_DRIVE_ID, sheet_name)
        created_new = True
        print(f"\n  Sheet created id={ssid} (new)")
    else:
        print(f"\n  Sheet found id={ssid} (updating in place)")

    # ── Write the weekly sheet ───────────────────────────────────────
    print(f"\n  Writing weekly sheet (total {total_rows} rows)...")
    existing_tabs = get_sheet_tabs(sheets, ssid)

    # Make sure each of our 3 tabs exists
    for tab_name in WEEKLY_SOURCE_TABS.keys():
        ensure_tab(sheets, ssid, tab_name, existing_tabs)

    # Clear each weekly tab, then write header + data
    value_payloads = []
    visibility_requests = []
    for weekly_tab in WEEKLY_SOURCE_TABS.keys():
        rows_for_tab = collected[weekly_tab]
        data_rows = [row[:SENT_WEEK_COL] for _, row in rows_for_tab]  # drop col O
        payload = [WEEKLY_HEADERS] + data_rows
        value_payloads.append({"range": f"'{weekly_tab}'!A1", "values": payload})
        # Hide tab if no data rows
        hidden = len(data_rows) == 0
        visibility_requests.append(
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": existing_tabs[weekly_tab],
                        "hidden": hidden,
                    },
                    "fields": "hidden",
                }
            }
        )

    # Clear old data
    _retry(
        lambda: sheets.spreadsheets()
        .values()
        .batchClear(
            spreadsheetId=ssid,
            body={"ranges": [f"'{t}'!A:Z" for t in WEEKLY_SOURCE_TABS.keys()]},
        )
        .execute()
    )

    # Write new data
    _retry(
        lambda: sheets.spreadsheets()
        .values()
        .batchUpdate(
            spreadsheetId=ssid,
            body={"valueInputOption": "RAW", "data": value_payloads},
        )
        .execute()
    )

    # Format each tab. First delete any existing bandings on our target tabs
    # so addBanding (not idempotent — errors if range already banded) re-runs
    # cleanly.
    target_sheet_ids = {existing_tabs[t] for t in WEEKLY_SOURCE_TABS.keys()}
    bandings_resp = _retry(
        lambda: sheets.spreadsheets()
        .get(
            spreadsheetId=ssid,
            fields="sheets(properties.sheetId,bandedRanges)",
        )
        .execute()
    )
    format_requests = []
    for sheet in bandings_resp.get("sheets", []):
        sid_here = sheet["properties"]["sheetId"]
        if sid_here not in target_sheet_ids:
            continue
        for br in sheet.get("bandedRanges", []):
            format_requests.append(
                {"deleteBanding": {"bandedRangeId": br["bandedRangeId"]}}
            )

    for weekly_tab in WEEKLY_SOURCE_TABS.keys():
        sheet_id = existing_tabs[weekly_tab]
        num_rows = len(collected[weekly_tab])
        format_requests.extend(
            build_tab_format_requests(sheet_id, len(WEEKLY_HEADERS), num_rows)
        )
    # Visibility + delete default Sheet1 if present
    format_requests.extend(visibility_requests)
    if "Sheet1" in existing_tabs:
        format_requests.append({"deleteSheet": {"sheetId": existing_tabs["Sheet1"]}})

    _retry(
        lambda: sheets.spreadsheets()
        .batchUpdate(spreadsheetId=ssid, body={"requests": format_requests})
        .execute()
    )

    # ── Instructions tab (v2.6.1 support-packet mode) ─────────────────
    # Only added when --all-unsent flag is used. Contains plain-language
    # support guidance and is pinned as the first tab (index 0) so support
    # sees it on open. Done as a SEPARATE batchUpdate after the data-tab
    # format batch so the index-0 move doesn't conflict with the Sheet1
    # delete, and so a stale Instructions tab from a prior --all-unsent
    # run can be cleared and rewritten cleanly.
    if all_unsent:
        print("\n  Writing Instructions tab (support-packet mode)...")
        # Re-read tabs so any addSheet from this run is included
        existing_tabs = get_sheet_tabs(sheets, ssid)
        instructions_sheet_id = ensure_tab(
            sheets, ssid, WEEKLY_TAB_INSTRUCTIONS, existing_tabs
        )

        per_tab_counts = {t: len(collected[t]) for t in WEEKLY_SOURCE_TABS.keys()}
        generation_iso = datetime.now(ZoneInfo(WEEKLY_TIMEZONE)).strftime(
            "%Y-%m-%d %H:%M %Z"
        )
        instructions_rows = _build_instructions_rows(
            generation_iso, total_rows, per_tab_counts
        )

        # Clear any old content
        _retry(
            lambda: sheets.spreadsheets()
            .values()
            .clear(
                spreadsheetId=ssid,
                range=f"'{WEEKLY_TAB_INSTRUCTIONS}'!A:Z",
                body={},
            )
            .execute()
        )
        # Write text rows (col A only)
        _retry(
            lambda: sheets.spreadsheets()
            .values()
            .update(
                spreadsheetId=ssid,
                range=f"'{WEEKLY_TAB_INSTRUCTIONS}'!A1",
                valueInputOption="RAW",
                body={"values": [[text] for text, _ in instructions_rows]},
            )
            .execute()
        )
        # Apply per-row styling and pin to index 0
        instructions_format_requests = _instructions_format_requests(
            instructions_sheet_id, instructions_rows
        )
        _retry(
            lambda: sheets.spreadsheets()
            .batchUpdate(
                spreadsheetId=ssid,
                body={"requests": instructions_format_requests},
            )
            .execute()
        )
        print(f"    Instructions tab populated ({len(instructions_rows)} rows)")

    # ── Mark cumulative-tab rows as sent ─────────────────────────────
    # v2.5.3: stamp by student_id lookup, NOT by stored row number. The
    # stored row numbers from collected[...] are from the earlier read pass
    # (~5s ago); Apps Script's removeStudentFromCumulativeTabs_ may have
    # deleted/shifted rows since then, which would cause stale row numbers
    # to land on the WRONG row. Re-reading col M (Student_ID) right before
    # stamping shrinks the race window from ~5s to ~milliseconds.
    print("\n  Marking selected rows with Sent Week...")
    col_letter = _col_letter(SENT_WEEK_COL)
    SID_COL_INDEX = 12  # col M = Student_ID (0-indexed) in the 14-col layout
    mark_data = []
    mark_counts = {}
    skipped_deleted = (
        {}
    )  # source_tab -> count of rows that vanished between read and stamp
    for weekly_tab, source_tab in WEEKLY_SOURCE_TABS.items():
        count = 0
        skipped = 0
        if not collected[weekly_tab]:
            mark_counts[source_tab] = 0
            skipped_deleted[source_tab] = 0
            continue

        # Re-read col M to get CURRENT row positions, then stamp by sid lookup.
        sid_resp = _retry(
            lambda t=source_tab: sheets.spreadsheets()
            .values()
            .get(spreadsheetId=OUTPUT_SPREADSHEET_ID, range=f"'{t}'!M:M")
            .execute(),
            label=f"re-read {source_tab} student_ids for stamping",
        )
        current_sid_to_row = {}
        for i, sid_row in enumerate(sid_resp.get("values", []), start=1):
            if sid_row and sid_row[0]:
                current_sid_to_row[str(sid_row[0]).strip()] = i

        for _orig_row_num, row in collected[weekly_tab]:
            sid = str(row[SID_COL_INDEX] or "").strip()
            if not sid or sid not in current_sid_to_row:
                # Row was deleted/shifted by Apps Script between read and stamp.
                # Skip silently — the row will be picked up next snapshot run if
                # it still exists with blank Sent Week.
                skipped += 1
                continue
            current_row_num = current_sid_to_row[sid]
            # Only write if not already marked with this week's date
            if str(row[SENT_WEEK_COL] or "").strip() != monday_iso:
                mark_data.append(
                    {
                        "range": f"'{source_tab}'!{col_letter}{current_row_num}",
                        "values": [[monday_iso]],
                    }
                )
                count += 1
        mark_counts[source_tab] = count
        skipped_deleted[source_tab] = skipped

    if mark_data:
        _retry(
            lambda: sheets.spreadsheets()
            .values()
            .batchUpdate(
                spreadsheetId=OUTPUT_SPREADSHEET_ID,
                body={"valueInputOption": "RAW", "data": mark_data},
            )
            .execute()
        )
    for src, cnt in mark_counts.items():
        print(f"    {src}: stamped {cnt} row(s) with {monday_iso}")

    # ── Done ─────────────────────────────────────────────────────────
    elapsed = time.time() - start
    print("\n" + "-" * 70)
    print(
        f"  '{sheet_name}' "
        f"({'CREATED' if created_new else 'UPDATED'}) — {elapsed:.1f}s"
    )
    per_tab = ", ".join(f"{t}: {len(collected[t])}" for t in WEEKLY_SOURCE_TABS.keys())
    print(f"    Rows per tab: {per_tab}")
    print(f"    Location: Shared Drive '{WEEKLY_SHARED_DRIVE_NAME}'")
    print(f"    URL: https://docs.google.com/spreadsheets/d/{ssid}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=(
            "Create/update the weekly corrections snapshot in the " "Shared Drive."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python generate_weekly_snapshot.py\n"
            "      Default: bundle blank-Sent-Week rows + rows already\n"
            "      stamped with the current Monday. This is what the\n"
            "      Monday cron runs.\n\n"
            "  python generate_weekly_snapshot.py --all-unsent\n"
            "      Support-packet mode: bundle EVERY row across all\n"
            "      cumulative tabs whose Sent Week is blank, regardless\n"
            "      of week. Adds an 'Instructions' tab pinned as the\n"
            "      first tab so support has plain-language guidance.\n"
            "      Use this when generating an ad-hoc handoff.\n"
        ),
    )
    parser.add_argument(
        "--all-unsent",
        action="store_true",
        help=(
            "Include every row with a blank Sent Week (any week, not just "
            "this week). Adds an Instructions tab. Use for ad-hoc support "
            "packets."
        ),
    )
    args = parser.parse_args()

    try:
        main(all_unsent=args.all_unsent)
    except Exception as e:
        print(f"\n  FATAL: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)
