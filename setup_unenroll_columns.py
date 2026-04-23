"""setup_unenroll_columns.py — One-time setup for the Unenroll workflow.

For each of 9 ISRs (Individual Student Rosters):
  1. Student Roster (SR) tab: ensure "Unenroll" column exists at ISR_CONFIG
     position with checkbox data validation on all data rows.
  2. MAP Roster (MR) tab: add "Unenroll" column (mirror formula) at the end
     so it auto-reflects SR's Unenroll value.

For the Combined MAP Roster (CMR):
  3. For each of 9 campus tabs: write IMPORTRANGE formula in the existing
     "Unenroll" column header, pulling the MR's Unenroll column from the ISR.

Run once after config.py ISR_CONFIG is set:
    python setup_unenroll_columns.py
"""

import sys
import time

from google.oauth2 import service_account
from googleapiclient.discovery import build

from config import (
    SERVICE_ACCOUNT_KEY,
    SCOPES,
    MAP_SPREADSHEET_ID,
    ISR_CONFIG,
)


def col_letter(i):
    """0-based index -> A1 notation column letter."""
    if i < 26:
        return chr(65 + i)
    return chr(65 + i // 26 - 1) + chr(65 + i % 26)


def get_sheet_id(sheets, spreadsheet_id, tab_name):
    resp = (
        sheets.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
        .execute()
    )
    for s in resp.get("sheets", []):
        if s["properties"]["title"] == tab_name:
            return s["properties"]["sheetId"]
    raise ValueError(f"Tab '{tab_name}' not found in {spreadsheet_id}")


def get_sheet_props(sheets, spreadsheet_id, tab_name):
    """Return (sheetId, rowCount, colCount) for the named tab."""
    resp = (
        sheets.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
        .execute()
    )
    for s in resp.get("sheets", []):
        if s["properties"]["title"] == tab_name:
            p = s["properties"]
            grid = p.get("gridProperties", {})
            return p["sheetId"], grid.get("rowCount", 1000), grid.get("columnCount", 26)
    raise ValueError(f"Tab '{tab_name}' not found in {spreadsheet_id}")


def ensure_grid_cols(sheets, spreadsheet_id, sheet_id, current_cols, min_cols):
    """Append columns if the sheet has fewer than min_cols."""
    if current_cols >= min_cols:
        return
    reqs = [
        {
            "appendDimension": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "length": min_cols - current_cols,
            }
        }
    ]
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": reqs}
    ).execute()


def ensure_sr_unenroll(sheets, isr_id, sr_col_idx, row_count):
    """Ensure SR's Unenroll column has header + checkbox validation.

    Handles 2 edge cases:
    - Grid too narrow: appends columns to fit sr_col_idx
    - Column is inside a Google Sheets Table (typed): skip data validation
      (the Table definition already provides checkbox rendering)
    """
    sr_gid, _, sr_cols = get_sheet_props(sheets, isr_id, "Student Roster")
    ensure_grid_cols(sheets, isr_id, sr_gid, sr_cols, sr_col_idx + 1)
    col = col_letter(sr_col_idx)

    # Write header
    sheets.spreadsheets().values().update(
        spreadsheetId=isr_id,
        range=f"'Student Roster'!{col}1",
        valueInputOption="RAW",
        body={"values": [["Unenroll"]]},
    ).execute()

    # Try to apply checkbox data validation; if it fails because column is in
    # a Table (typed column), skip — the Table already provides checkboxes.
    reqs = [
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sr_gid,
                    "startRowIndex": 1,
                    "endRowIndex": row_count,
                    "startColumnIndex": sr_col_idx,
                    "endColumnIndex": sr_col_idx + 1,
                },
                "rule": {
                    "condition": {"type": "BOOLEAN"},
                    "showCustomUi": True,
                },
            }
        }
    ]
    try:
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=isr_id, body={"requests": reqs}
        ).execute()
    except Exception as e:
        if "typed columns" in str(e):
            print(f"    (skipping data validation — col inside a Table)")
        else:
            raise


def ensure_mr_unenroll(sheets, isr_id, sr_col_idx, mr_col_idx, row_count):
    """Add MR's Unenroll mirror column: =IF('Student Roster'!X2, TRUE, FALSE)."""
    mr_gid, _, mr_cols = get_sheet_props(sheets, isr_id, "MAP Roster")
    ensure_grid_cols(sheets, isr_id, mr_gid, mr_cols, mr_col_idx + 1)
    col = col_letter(mr_col_idx)
    sr_col = col_letter(sr_col_idx)

    # Write header
    sheets.spreadsheets().values().update(
        spreadsheetId=isr_id,
        range=f"'MAP Roster'!{col}1",
        valueInputOption="RAW",
        body={"values": [["Unenroll"]]},
    ).execute()

    # Write mirror formulas for every data row
    # Use IF to ensure we output boolean TRUE/FALSE and not blank "" from empty SR cells.
    formulas = []
    for r in range(2, row_count + 1):
        formulas.append([f"=IF('Student Roster'!{sr_col}{r}=TRUE, TRUE, FALSE)"])
    sheets.spreadsheets().values().update(
        spreadsheetId=isr_id,
        range=f"'MAP Roster'!{col}2:{col}{row_count}",
        valueInputOption="USER_ENTERED",
        body={"values": formulas},
    ).execute()

    # Apply checkbox data validation so MR column renders as checkboxes
    reqs = [
        {
            "setDataValidation": {
                "range": {
                    "sheetId": mr_gid,
                    "startRowIndex": 1,
                    "endRowIndex": row_count,
                    "startColumnIndex": mr_col_idx,
                    "endColumnIndex": mr_col_idx + 1,
                },
                "rule": {
                    "condition": {"type": "BOOLEAN"},
                    "showCustomUi": True,
                },
            }
        }
    ]
    try:
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=isr_id, body={"requests": reqs}
        ).execute()
    except Exception as e:
        if "typed columns" in str(e):
            print(f"    (skipping MR data validation — col inside a Table)")
        else:
            raise


def setup_cmr_importrange(sheets, cmr_tab, isr_id, mr_col_idx):
    """Write IMPORTRANGE formula in CMR's Unenroll column, pulling from ISR's MR."""
    # Find the CMR's Unenroll column (varies per campus: AD [29] or AC [28])
    resp = (
        sheets.spreadsheets()
        .values()
        .get(spreadsheetId=MAP_SPREADSHEET_ID, range=f"'{cmr_tab}'!A1:AE1")
        .execute()
    )
    headers = resp.get("values", [[]])[0]
    cmr_unenroll_col = None
    for i, h in enumerate(headers):
        if str(h).strip().lower() == "unenroll":
            cmr_unenroll_col = i
            break
    if cmr_unenroll_col is None:
        print(f"  [SKIP] {cmr_tab}: no Unenroll column header found")
        return

    cmr_col = col_letter(cmr_unenroll_col)
    mr_col = col_letter(mr_col_idx)

    # Build IMPORTRANGE formula
    formula = (
        f'=IMPORTRANGE("https://docs.google.com/spreadsheets/d/{isr_id}",'
        f'"MAP Roster!{mr_col}2:{mr_col}")'
    )

    sheets.spreadsheets().values().update(
        spreadsheetId=MAP_SPREADSHEET_ID,
        range=f"'{cmr_tab}'!{cmr_col}2",
        valueInputOption="USER_ENTERED",
        body={"values": [[formula]]},
    ).execute()
    print(f"  [OK] CMR '{cmr_tab}' {cmr_col}2 <- IMPORTRANGE from ISR MR!{mr_col}")


def main():
    start = time.time()
    print("=" * 70)
    print("  UNENROLL WORKFLOW SETUP")
    print("=" * 70)

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEY, scopes=SCOPES
    )
    sheets = build("sheets", "v4", credentials=creds)

    # Step 1+2: for each ISR, set up SR + MR Unenroll columns
    for cmr_tab, conf in ISR_CONFIG.items():
        print(f"\n--- {cmr_tab} ---")
        isr_id = conf["isr_id"]
        sr_col = conf["sr_unenroll_col"]
        mr_col = conf["mr_unenroll_col"]
        # Use a conservative row count of 1200 (matches MR default; SR will cover its own size too)
        row_count = 1200
        try:
            ensure_sr_unenroll(sheets, isr_id, sr_col, row_count)
            print(
                f"  [OK] SR col {col_letter(sr_col)} = 'Unenroll' with checkbox validation"
            )
        except Exception as e:
            print(f"  [ERR] SR setup failed: {e}")
            continue

        try:
            ensure_mr_unenroll(sheets, isr_id, sr_col, mr_col, row_count)
            print(
                f"  [OK] MR col {col_letter(mr_col)} = 'Unenroll' mirror formula + checkbox"
            )
        except Exception as e:
            print(f"  [ERR] MR setup failed: {e}")
            continue

    # Step 3: CMR IMPORTRANGE for each campus tab
    print("\n" + "=" * 70)
    print("  CMR: Wire up IMPORTRANGE on all 9 Unenroll columns")
    print("=" * 70)
    for cmr_tab, conf in ISR_CONFIG.items():
        try:
            setup_cmr_importrange(
                sheets, cmr_tab, conf["isr_id"], conf["mr_unenroll_col"]
            )
        except Exception as e:
            print(f"  [ERR] {cmr_tab}: {e}")

    elapsed = time.time() - start
    print(f"\n  Completed in {elapsed:.1f}s")
    print("\n  NEXT STEP:")
    print("  Open the CMR spreadsheet and click 'Allow access' on each")
    print("  IMPORTRANGE permission prompt (one per ISR).")
    print(f"  CMR: https://docs.google.com/spreadsheets/d/{MAP_SPREADSHEET_ID}")


if __name__ == "__main__":
    main()
