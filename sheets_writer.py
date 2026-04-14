"""Google Sheets API writer for the Weekly Corrections tool."""

import time

from config import (
    OUTPUT_SPREADSHEET_ID,
    TAB_CORRECTED,
    TAB_SIS,
    TAB_APPROVED,
    OUTPUT_FIELDS,
)

# ── API helpers (from email_winners.py pattern) ────────────────────────────


def _retry_api(fn, max_retries=3, delay=5):
    """Retry a Sheets API call with backoff."""
    for attempt in range(max_retries):
        try:
            return fn()
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            print(f"     Warning: API call failed (attempt {attempt+1}): {e}")
            time.sleep(delay * (attempt + 1))


def _ensure_tab_exists(sheets_service, spreadsheet_id, tab_name):
    """Create the tab if it doesn't exist; return its sheetId."""
    resp = _retry_api(
        lambda: sheets_service.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
        .execute()
    )
    for sheet in resp.get("sheets", []):
        props = sheet["properties"]
        if props["title"] == tab_name:
            return props["sheetId"]

    result = _retry_api(
        lambda: sheets_service.spreadsheets()
        .batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]},
        )
        .execute()
    )
    sheet_id = result["replies"][0]["addSheet"]["properties"]["sheetId"]
    print(f"     Created tab '{tab_name}' (sheetId={sheet_id})")
    return sheet_id


def _rgb(hex_color):
    """Convert hex color string to Sheets API RGB dict."""
    h = hex_color.lstrip("#")
    return {
        "red": int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue": int(h[4:6], 16) / 255,
    }


def _clear_banding(sheets_service, spreadsheet_id, sheet_ids):
    """Remove existing banded ranges from the given sheets."""
    resp = _retry_api(
        lambda: sheets_service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields="sheets.bandedRanges,sheets.properties.sheetId",
        )
        .execute()
    )
    requests = []
    for sheet in resp.get("sheets", []):
        sid = sheet["properties"]["sheetId"]
        if sid not in sheet_ids:
            continue
        for br in sheet.get("bandedRanges", []):
            requests.append({"deleteBanding": {"bandedRangeId": br["bandedRangeId"]}})
    if requests:
        _retry_api(
            lambda: sheets_service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests})
            .execute()
        )


# ── Main write function ───────────────────────────────────────────────────


def write_corrections(sheets_service, corrections_map, corrections_sis):
    """Write comparison results to the output spreadsheet.

    Args:
        sheets_service: Google Sheets API v4 service object
        corrections_map: list of dicts — MAP roster data for mismatched students
        corrections_sis: list of dicts — SIS data for same students, same order
    """
    sid = OUTPUT_SPREADSHEET_ID
    print(f"\n  Writing {len(corrections_map)} correction rows...")

    # Ensure all 3 tabs exist
    sheet1_id = _ensure_tab_exists(sheets_service, sid, TAB_CORRECTED)
    sheet2_id = _ensure_tab_exists(sheets_service, sid, TAB_SIS)
    _ensure_tab_exists(sheets_service, sid, TAB_APPROVED)

    # ── Clear Sheet 1 and Sheet 2 (NOT Sheet 3 — cumulative) ──────────
    # Remove existing banding + conditional formatting before re-applying
    _clear_banding(sheets_service, sid, [sheet1_id, sheet2_id])
    for tab in [TAB_CORRECTED, TAB_SIS]:
        _retry_api(
            lambda t=tab: sheets_service.spreadsheets()
            .values()
            .clear(spreadsheetId=sid, range=f"'{t}'!A:Z")
            .execute()
        )

    if not corrections_map:
        print("  No mismatches found — sheets cleared.")
        _write_empty_state(sheets_service, sid, sheet1_id, sheet2_id)
        return

    # ── Build Sheet 1 rows (checkbox + MAP data + mismatch summary) ───
    header1 = ["✓"] + OUTPUT_FIELDS + ["Mismatch Summary"]
    rows1 = [header1]
    for rec in corrections_map:
        row = [False]  # checkbox default unchecked
        for field in OUTPUT_FIELDS:
            row.append(rec.get(field, ""))
        row.append(rec.get("mismatch_summary", ""))
        rows1.append(row)

    # ── Build Sheet 2 rows (SIS data, no checkbox) ────────────────────
    header2 = OUTPUT_FIELDS[:]
    rows2 = [header2]
    for rec in corrections_sis:
        row = []
        for field in OUTPUT_FIELDS:
            row.append(rec.get(field, ""))
        rows2.append(row)

    # ── Write data ────────────────────────────────────────────────────
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range=f"'{TAB_CORRECTED}'!A1",
            valueInputOption="RAW",
            body={"values": rows1},
        )
        .execute()
    )

    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range=f"'{TAB_SIS}'!A1",
            valueInputOption="RAW",
            body={"values": rows2},
        )
        .execute()
    )

    # ── Format both sheets ────────────────────────────────────────────
    fmt_requests = []
    fmt_requests.extend(
        _format_sheet(sheet1_id, len(rows1), len(header1), has_checkbox=True)
    )
    fmt_requests.extend(
        _format_sheet(sheet2_id, len(rows2), len(header2), has_checkbox=False)
    )

    # Add checkbox data validation on Sheet 1 column A (rows 2+)
    if len(rows1) > 1:
        fmt_requests.append(
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet1_id,
                        "startRowIndex": 1,
                        "endRowIndex": len(rows1),
                        "startColumnIndex": 0,
                        "endColumnIndex": 1,
                    },
                    "rule": {
                        "condition": {"type": "BOOLEAN"},
                        "showCustomUi": True,
                    },
                }
            }
        )

    if fmt_requests:
        _retry_api(
            lambda: sheets_service.spreadsheets()
            .batchUpdate(
                spreadsheetId=sid,
                body={"requests": fmt_requests},
            )
            .execute()
        )

    # ── Ensure Sheet 3 has headers if empty ───────────────────────────
    _ensure_approved_headers(sheets_service, sid)

    print(f"  Done — {len(corrections_map)} students written to corrections sheet.")


def _write_empty_state(sheets_service, sid, sheet1_id, sheet2_id):
    """Write a 'no mismatches' message when there are no corrections."""
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range=f"'{TAB_CORRECTED}'!A1",
            valueInputOption="RAW",
            body={"values": [["No mismatches found between MAP roster and SIS data."]]},
        )
        .execute()
    )
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range=f"'{TAB_SIS}'!A1",
            valueInputOption="RAW",
            body={"values": [["No mismatches found between MAP roster and SIS data."]]},
        )
        .execute()
    )


def _format_sheet(sheet_id, num_rows, num_cols, has_checkbox=False):
    """Build formatting requests for a corrections sheet."""
    requests = []

    # Header formatting — navy background, white bold text
    header_bg = _rgb("1E3A5F")
    header_fg = _rgb("FFFFFF")
    requests.append(
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
                        "backgroundColor": header_bg,
                        "textFormat": {
                            "foregroundColor": header_fg,
                            "bold": True,
                            "fontSize": 10,
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
            }
        }
    )

    # Freeze header row
    requests.append(
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 1},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        }
    )

    # Alternating row colors (light grey on even rows)
    if num_rows > 1:
        requests.append(
            {
                "addBanding": {
                    "bandedRange": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": num_rows,
                            "startColumnIndex": 0,
                            "endColumnIndex": num_cols,
                        },
                        "rowProperties": {
                            "headerColor": header_bg,
                            "firstBandColor": _rgb("FFFFFF"),
                            "secondBandColor": _rgb("EDF2F7"),
                        },
                    }
                }
            }
        )

    # Column widths
    col_widths = {
        0: 30 if has_checkbox else 150,  # checkbox or Campus
        1: 150 if has_checkbox else 60,  # Campus or Grade
    }
    # Set reasonable widths for data columns
    data_start = 1 if has_checkbox else 0
    field_widths = [150, 60, 80, 100, 100, 220, 150, 100, 100, 220, 120, 140]
    for i, w in enumerate(field_widths):
        col_idx = data_start + i
        if col_idx < num_cols:
            col_widths[col_idx] = w

    # Mismatch summary column (last column on Sheet 1)
    if has_checkbox and num_cols > len(field_widths) + 1:
        col_widths[num_cols - 1] = 200

    for col, width in col_widths.items():
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": col,
                        "endIndex": col + 1,
                    },
                    "properties": {"pixelSize": width},
                    "fields": "pixelSize",
                }
            }
        )

    return requests


def _ensure_approved_headers(sheets_service, sid):
    """Write headers to Sheet 3 if it's empty."""
    resp = _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range=f"'{TAB_APPROVED}'!A1:A1")
        .execute()
    )
    if resp.get("values"):
        return  # already has content

    header = ["Date Approved"] + OUTPUT_FIELDS
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range=f"'{TAB_APPROVED}'!A1",
            valueInputOption="RAW",
            body={"values": [header]},
        )
        .execute()
    )

    # Format header
    sheet3_id = _ensure_tab_exists(sheets_service, sid, TAB_APPROVED)
    header_bg = _rgb("1E3A5F")
    header_fg = _rgb("FFFFFF")
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .batchUpdate(
            spreadsheetId=sid,
            body={
                "requests": [
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": sheet3_id,
                                "startRowIndex": 0,
                                "endRowIndex": 1,
                                "startColumnIndex": 0,
                                "endColumnIndex": len(header),
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": header_bg,
                                    "textFormat": {
                                        "foregroundColor": header_fg,
                                        "bold": True,
                                        "fontSize": 10,
                                    },
                                    "horizontalAlignment": "CENTER",
                                }
                            },
                            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                        }
                    },
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": sheet3_id,
                                "gridProperties": {"frozenRowCount": 1},
                            },
                            "fields": "gridProperties.frozenRowCount",
                        }
                    },
                ]
            },
        )
        .execute()
    )
