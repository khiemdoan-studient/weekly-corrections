"""build_unenroll_queue.py — Create/refresh the real-time "Unenroll Queue (Live)" sheet.

This visible sheet lives in the corrections spreadsheet and pulls IM-flagged
Unenroll students from all 9 CMR campus tabs in real-time via QUERY + IMPORTRANGE.

Latency: ~1 minute (IMPORTRANGE refresh window). Shows flagged students
immediately, but does NOT do SIS comparison — that's the hourly Python
pipeline's job.

Run once to create the sheet. Re-runs are idempotent.
    python build_unenroll_queue.py
"""

import time

from google.oauth2 import service_account
from googleapiclient.discovery import build

from config import (
    SERVICE_ACCOUNT_KEY,
    SCOPES,
    MAP_SPREADSHEET_ID,
    OUTPUT_SPREADSHEET_ID,
    ISR_CONFIG,
)

TAB_NAME = "Unenroll Queue (Live)"


def _rgb(h):
    h = h.lstrip("#")
    return {
        "red": int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue": int(h[4:6], 16) / 255,
    }


NAVY_DARK = _rgb("0F1B33")
NAVY_MED = _rgb("263D66")
GREY_LABEL = _rgb("94A3B8")
WHITE = _rgb("FFFFFF")
FILTER_BG = _rgb("1E3A5F")
RED_LIGHT = _rgb("FEE2E2")  # light red for Unenrolling


def ensure_tab(sheets, spreadsheet_id, tab_name):
    resp = (
        sheets.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
        .execute()
    )
    for s in resp.get("sheets", []):
        if s["properties"]["title"] == tab_name:
            return s["properties"]["sheetId"]
    result = (
        sheets.spreadsheets()
        .batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]},
        )
        .execute()
    )
    return result["replies"][0]["addSheet"]["properties"]["sheetId"]


def build_campus_formula(cmr_tab, unenroll_col_letter):
    """Build QUERY+IMPORTRANGE formula for a single campus.

    Pulls flagged rows (Unenroll=TRUE) and returns:
    [Campus, Student_ID, Student_Email, Last_Name, First_Name, Grade, Unenroll_Status].

    Handles varying column layouts:
    - Reading CCSD has Full Name at col G, shifting Grade from G→H
    - Others: Student_ID=A, Email=B, Campus=C, Last=E, First=F, Grade=G
    """
    # Reading CCSD's grade is at col H (not G) because of inserted Full Name column
    grade_col = "Col8" if cmr_tab == "Reading CCSD (Dash)" else "Col7"

    importrange = (
        f'IMPORTRANGE("https://docs.google.com/spreadsheets/d/{MAP_SPREADSHEET_ID}",'
        f'"{cmr_tab}!A2:AE")'
    )
    unenroll_col_num = (
        ord(unenroll_col_letter[-1]) - ord("A") + 1
        if len(unenroll_col_letter) == 1
        else 26 + (ord(unenroll_col_letter[1]) - ord("A") + 1)
    )

    # NOTE: No IFERROR wrapper — if IMPORTRANGE fails (e.g. auth not granted yet),
    # the user sees #REF! which is the visible signal to click Allow access.
    formula = (
        f"=QUERY({importrange}, "
        f'"SELECT Col3, Col1, Col2, Col5, Col6, {grade_col}, Col{unenroll_col_num} '
        f"WHERE Col{unenroll_col_num} = TRUE "
        f"LABEL Col3 '', Col1 '', Col2 '', Col5 '', Col6 '', {grade_col} '', Col{unenroll_col_num} ''\", 0)"
    )
    return formula


def main():
    start = time.time()
    print("=" * 70)
    print("  BUILD UNENROLL QUEUE (LIVE)")
    print("=" * 70)

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEY, scopes=SCOPES
    )
    sheets = build("sheets", "v4", credentials=creds)

    sheet_id = ensure_tab(sheets, OUTPUT_SPREADSHEET_ID, TAB_NAME)
    print(f"  Tab '{TAB_NAME}' sheetId={sheet_id}")

    # Clear existing content
    sheets.spreadsheets().values().clear(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        range=f"'{TAB_NAME}'!A:Z",
    ).execute()

    # Build the layout: title, caption, headers, 9 campus formulas stacked vertically.
    NC = 7  # Campus, Student_ID, Email, Last, First, Grade, Unenroll

    # Collect values to write in batch
    values_data = []

    # Row 1: Title (merged)
    values_data.append(
        {"range": f"'{TAB_NAME}'!A1", "values": [["Unenroll Queue (Live)"]]}
    )
    # Row 2: Caption
    values_data.append(
        {
            "range": f"'{TAB_NAME}'!A2",
            "values": [
                [
                    "Real-time feed of IM-flagged Unenroll students across all 9 campuses (~1 min latency). "
                    "IF DATA BELOW IS EMPTY OR SHOWS #REF!, CLICK 'ALLOW ACCESS' ON THE IMPORTRANGE PROMPT ONCE."
                ]
            ],
        }
    )
    # Row 3: spacer
    values_data.append({"range": f"'{TAB_NAME}'!A3", "values": [[""]]})
    # Row 4: column headers
    values_data.append(
        {
            "range": f"'{TAB_NAME}'!A4",
            "values": [
                [
                    "Campus",
                    "Student_ID",
                    "Student_Email",
                    "Last Name",
                    "First Name",
                    "Grade",
                    "Unenroll",
                ]
            ],
        }
    )

    # Rows 5+: one campus block per campus. Each block is a QUERY formula that
    # auto-expands to however many flagged rows exist. We stack them by writing
    # each formula at a different row so they don't overlap.
    # Strategy: allocate 50 rows per campus block (more than enough for any
    # realistic flagged-count), and put each formula at the start of its block.
    ROWS_PER_CAMPUS = 50
    current_row = 5
    formulas_data = []
    for cmr_tab, conf in ISR_CONFIG.items():
        # Look up CMR Unenroll column position by reading the CMR header
        headers_resp = (
            sheets.spreadsheets()
            .values()
            .get(spreadsheetId=MAP_SPREADSHEET_ID, range=f"'{cmr_tab}'!A1:AE1")
            .execute()
        )
        headers = headers_resp.get("values", [[]])[0]
        unenroll_col_idx = None
        for i, h in enumerate(headers):
            if str(h).strip().lower() == "unenroll":
                unenroll_col_idx = i
                break
        if unenroll_col_idx is None:
            print(f"  [SKIP] {cmr_tab}: no Unenroll header")
            continue

        unenroll_letter = (
            chr(65 + unenroll_col_idx)
            if unenroll_col_idx < 26
            else chr(65 + unenroll_col_idx // 26 - 1) + chr(65 + unenroll_col_idx % 26)
        )
        formula = build_campus_formula(cmr_tab, unenroll_letter)
        formulas_data.append(
            {"range": f"'{TAB_NAME}'!A{current_row}", "values": [[formula]]}
        )
        print(f"  Row {current_row}: {cmr_tab} (Unenroll at col {unenroll_letter})")
        current_row += ROWS_PER_CAMPUS

    # Batch write headers
    sheets.spreadsheets().values().batchUpdate(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        body={
            "valueInputOption": "RAW",
            "data": values_data,
        },
    ).execute()
    # Batch write formulas (USER_ENTERED to evaluate)
    sheets.spreadsheets().values().batchUpdate(
        spreadsheetId=OUTPUT_SPREADSHEET_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": formulas_data,
        },
    ).execute()

    # Apply formatting
    fmt = [
        # Title row: merged, navy dark, 20pt bold white
        {
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": NC,
                },
                "mergeType": "MERGE_ALL",
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": NC,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": NAVY_DARK,
                        "textFormat": {
                            "foregroundColor": WHITE,
                            "bold": True,
                            "fontSize": 20,
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 0,
                    "endIndex": 1,
                },
                "properties": {"pixelSize": 55},
                "fields": "pixelSize",
            }
        },
        # Caption row: merged, navy med, italic grey
        {
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 0,
                    "endColumnIndex": NC,
                },
                "mergeType": "MERGE_ALL",
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 0,
                    "endColumnIndex": NC,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": NAVY_MED,
                        "textFormat": {
                            "foregroundColor": GREY_LABEL,
                            "italic": True,
                            "fontSize": 12,
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "wrapStrategy": "WRAP",
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 1,
                    "endIndex": 2,
                },
                "properties": {"pixelSize": 60},
                "fields": "pixelSize",
            }
        },
        # Header row (row 4 idx 3): navy bold white
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 3,
                    "endRowIndex": 4,
                    "startColumnIndex": 0,
                    "endColumnIndex": NC,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": FILTER_BG,
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
        # Freeze header rows
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 4},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        },
        # Data area: light yellow background for Unenrolling type
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 4,
                    "endRowIndex": 500,
                    "startColumnIndex": 0,
                    "endColumnIndex": NC,
                },
                "cell": {"userEnteredFormat": {"backgroundColor": RED_LIGHT}},
                "fields": "userEnteredFormat.backgroundColor",
            }
        },
    ]

    # Column widths
    field_widths = [220, 120, 220, 120, 120, 60, 80]
    for i, w in enumerate(field_widths):
        fmt.append(
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

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=OUTPUT_SPREADSHEET_ID, body={"requests": fmt}
    ).execute()

    elapsed = time.time() - start
    print(f"\n  Completed in {elapsed:.1f}s")
    print(f"\n  Output: https://docs.google.com/spreadsheets/d/{OUTPUT_SPREADSHEET_ID}")
    print("  Open the 'Unenroll Queue (Live)' tab. First load may need 'Allow access'")
    print("  clicks on IMPORTRANGE prompts (once per source ISR).")


if __name__ == "__main__":
    main()
