"""setup_summer_school_columns.py - Provision Summer School columns (v2.9.0).

Idempotent. For the 3 summer-school campuses (JHMS, JHES, JRHS) it adds 5
columns to each layer of the roster chain, appended AFTER the existing
last-used column so nothing existing shifts:

  Student Roster (SR, ISR):  Summer School (checkbox) + Teacher Email + Teacher
                             + Grade + Subjects   [typed source; loader fills]
  MAP Roster   (MR, ISR):    same 5 as =ARRAYFORMULA mirrors of the SR columns
  Combined MAP Roster (CMR): same 5 as =IMPORTRANGE from the ISR MR columns

Mirrors setup_unenroll_columns.py (the Unenroll precedent: SR -> MR -> CMR).
The per-student VALUES (which students are in summer school, their teacher /
grade) are written separately by the one-time data loader, NOT here, so this
module is PII-free and safe to commit + re-run.

Run:  python setup_summer_school_columns.py
"""

import time

from google.oauth2 import service_account
from googleapiclient.discovery import build

from config import SERVICE_ACCOUNT_KEY, SCOPES, MAP_SPREADSHEET_ID, ISR_CONFIG
from retry_helper import retry_api

# The 5 Summer School columns, in display order. The first is a boolean checkbox.
SUMMER_HEADERS = [
    "Summer School",
    "Summer School Teacher Email",
    "Summer School Teacher",
    "Summer School Grade",
    "Summer School Subjects",
]
FLAG_HEADER = SUMMER_HEADERS[0]

# Campuses running summer school (CMR tab name -> ISR id via ISR_CONFIG).
SUMMER_TABS = [
    "Hardeeville Junior & Senior High School (Dash)",
    "Hardeeville Elementary School (Dash)",
    "Ridgeland Secondary Academy of Excellence (Dash)",
    "Allendale Fairfax Middle School (Dash)",  # v2.9.1
]

SR_TAB = "Student Roster"
MR_TAB = "MAP Roster"

# Combined view tab (v2.9.1): every Summer School = TRUE row from all SUMMER_TABS.
ROSTER_TAB = "Summer School Roster"
# Core A:N columns are identical across all CMR campus tabs (the IMPORTRANGE core).
CORE_HEADERS = [
    "Student ID",
    "Student Email",
    "Campus",
    "NWEA Account",
    "Last Name",
    "First Name",
    "Grade",
    "Level",
    "DOB",
    "Gender",
    "Accommodations",
    "Subjects",
    "Start Date",
    "Notes",
]


def col_letter(i):
    """0-based index -> A1 column letter (handles A..ZZ)."""
    if i < 26:
        return chr(65 + i)
    return chr(65 + i // 26 - 1) + chr(65 + i % 26)


def get_sheet_props(sheets, ssid, tab):
    """Return (sheetId, rowCount, colCount) for the named tab."""
    resp = retry_api(
        lambda: sheets.spreadsheets()
        .get(spreadsheetId=ssid, fields="sheets.properties")
        .execute(),
        label=f"get props '{tab}'",
    )
    for s in resp.get("sheets", []):
        if s["properties"]["title"] == tab:
            p = s["properties"]
            g = p.get("gridProperties", {})
            return p["sheetId"], g.get("rowCount", 1000), g.get("columnCount", 26)
    raise ValueError(f"Tab '{tab}' not found in {ssid}")


def ensure_grid_cols(sheets, ssid, sheet_id, current_cols, min_cols):
    """Append columns if the sheet has fewer than min_cols."""
    if current_cols >= min_cols:
        return
    retry_api(
        lambda: sheets.spreadsheets()
        .batchUpdate(
            spreadsheetId=ssid,
            body={
                "requests": [
                    {
                        "appendDimension": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "length": min_cols - current_cols,
                        }
                    }
                ]
            },
        )
        .execute(),
        label="append cols",
    )


def read_header_row(sheets, ssid, tab):
    resp = retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .get(spreadsheetId=ssid, range=f"'{tab}'!1:1")
        .execute(),
        label=f"read header '{tab}'",
    )
    vals = resp.get("values", [])
    return vals[0] if vals else []


def assign_positions(header_row):
    """Find-or-assign a 0-based col index for each SUMMER_HEADERS entry.

    Existing headers are reused (idempotent). Missing ones are appended
    contiguously after the current last-used column. Returns {header: idx}.
    """
    existing = {
        str(h).strip().lower(): i for i, h in enumerate(header_row) if str(h).strip()
    }
    cursor = (max(existing.values()) if existing else -1) + 1
    positions = {}
    for h in SUMMER_HEADERS:
        key = h.lower()
        if key in existing:
            positions[h] = existing[key]
        else:
            positions[h] = cursor
            cursor += 1
    return positions


def write_headers(sheets, ssid, tab, positions):
    data = [
        {"range": f"'{tab}'!{col_letter(c)}1", "values": [[h]]}
        for h, c in positions.items()
    ]
    retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .batchUpdate(spreadsheetId=ssid, body={"valueInputOption": "RAW", "data": data})
        .execute(),
        label=f"write headers '{tab}'",
    )


def set_checkbox(sheets, ssid, sheet_id, col_idx, row_count):
    """Apply BOOLEAN (checkbox) data validation to a column's data rows."""
    reqs = [
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": row_count,
                    "startColumnIndex": col_idx,
                    "endColumnIndex": col_idx + 1,
                },
                "rule": {"condition": {"type": "BOOLEAN"}, "showCustomUi": True},
            }
        }
    ]
    try:
        retry_api(
            lambda: sheets.spreadsheets()
            .batchUpdate(spreadsheetId=ssid, body={"requests": reqs})
            .execute(),
            label="checkbox validation",
        )
    except Exception as e:
        if "typed columns" in str(e):
            print("    (skip checkbox - col inside a Table)")
        else:
            raise


def set_plain_number(sheets, ssid, sheet_id, col_idx, row_count):
    """Force a column's data rows to plain integer format. Grades were
    rendering as dates (e.g. 01/07/1900) because the appended cells inherited a
    date number-format; the underlying value is correct, this fixes display."""
    reqs = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": row_count,
                    "startColumnIndex": col_idx,
                    "endColumnIndex": col_idx + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "NUMBER", "pattern": "0"}
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        }
    ]
    retry_api(
        lambda: sheets.spreadsheets()
        .batchUpdate(spreadsheetId=ssid, body={"requests": reqs})
        .execute(),
        label="plain number format (grade)",
    )


def provision_sr(sheets, isr_id):
    """SR: append the 5 headers + checkbox validation on the flag col. Typed
    student values are written later by the loader. Returns {header: idx}."""
    gid, rows, cols = get_sheet_props(sheets, isr_id, SR_TAB)
    pos = assign_positions(read_header_row(sheets, isr_id, SR_TAB))
    ensure_grid_cols(sheets, isr_id, gid, cols, max(pos.values()) + 1)
    write_headers(sheets, isr_id, SR_TAB, pos)
    set_checkbox(sheets, isr_id, gid, pos[FLAG_HEADER], max(rows, 1200))
    set_plain_number(sheets, isr_id, gid, pos[SUMMER_HEADERS[3]], max(rows, 1200))
    print(
        f"  [OK] SR cols {col_letter(min(pos.values()))}..{col_letter(max(pos.values()))}"
        f" (flag={col_letter(pos[FLAG_HEADER])})"
    )
    return pos


def provision_mr(sheets, isr_id, sr_pos):
    """MR: append the 5 headers, each row-2 =ARRAYFORMULA mirror of the matching
    SR column (mirrors the existing MR column style). Returns {header: idx}."""
    gid, rows, cols = get_sheet_props(sheets, isr_id, MR_TAB)
    pos = assign_positions(read_header_row(sheets, isr_id, MR_TAB))
    ensure_grid_cols(sheets, isr_id, gid, cols, max(pos.values()) + 1)
    write_headers(sheets, isr_id, MR_TAB, pos)
    data = []
    for h in SUMMER_HEADERS:
        sr_col = col_letter(sr_pos[h])
        mr_col = col_letter(pos[h])
        data.append(
            {
                "range": f"'{MR_TAB}'!{mr_col}2",
                "values": [[f"=ARRAYFORMULA('{SR_TAB}'!{sr_col}2:{sr_col})"]],
            }
        )
    retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .batchUpdate(
            spreadsheetId=isr_id,
            body={"valueInputOption": "USER_ENTERED", "data": data},
        )
        .execute(),
        label="write MR arrayformulas",
    )
    set_checkbox(sheets, isr_id, gid, pos[FLAG_HEADER], max(rows, 1200))
    set_plain_number(sheets, isr_id, gid, pos[SUMMER_HEADERS[3]], max(rows, 1200))
    print(
        f"  [OK] MR cols {col_letter(min(pos.values()))}..{col_letter(max(pos.values()))}"
        f" = ARRAYFORMULA mirrors"
    )
    return pos


def provision_cmr(sheets, cmr_tab, isr_id, mr_pos):
    """CMR: append the 5 headers, each row-2 =IMPORTRANGE of the matching MR
    column (mirrors the Unenroll precedent)."""
    gid, rows, cols = get_sheet_props(sheets, MAP_SPREADSHEET_ID, cmr_tab)
    pos = assign_positions(read_header_row(sheets, MAP_SPREADSHEET_ID, cmr_tab))
    ensure_grid_cols(sheets, MAP_SPREADSHEET_ID, gid, cols, max(pos.values()) + 1)
    write_headers(sheets, MAP_SPREADSHEET_ID, cmr_tab, pos)
    data = []
    for h in SUMMER_HEADERS:
        mr_col = col_letter(mr_pos[h])
        cmr_col = col_letter(pos[h])
        formula = (
            f'=IMPORTRANGE("https://docs.google.com/spreadsheets/d/{isr_id}",'
            f'"{MR_TAB}!{mr_col}2:{mr_col}")'
        )
        data.append({"range": f"'{cmr_tab}'!{cmr_col}2", "values": [[formula]]})
    retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .batchUpdate(
            spreadsheetId=MAP_SPREADSHEET_ID,
            body={"valueInputOption": "USER_ENTERED", "data": data},
        )
        .execute(),
        label=f"write CMR importranges '{cmr_tab}'",
    )
    set_plain_number(
        sheets, MAP_SPREADSHEET_ID, gid, pos[SUMMER_HEADERS[3]], max(rows, 1200)
    )
    print(
        f"  [OK] CMR '{cmr_tab}' cols {col_letter(min(pos.values()))}.."
        f"{col_letter(max(pos.values()))} = IMPORTRANGE"
    )
    return pos


def read_summer_positions(sheets, ssid, tab):
    """Helper for the data loader: return {header: 0-based col idx} for the
    Summer School columns currently in `tab` (must be provisioned first)."""
    existing = {
        str(h).strip().lower(): i
        for i, h in enumerate(read_header_row(sheets, ssid, tab))
        if str(h).strip()
    }
    return {h: existing[h.lower()] for h in SUMMER_HEADERS if h.lower() in existing}


def build_sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEY, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)


def build_summer_roster_tab(sheets):
    """(Re)build the combined ROSTER_TAB on the CMR: every Summer School = TRUE
    row from all SUMMER_TABS, normalized to core A:N + the 5 summer columns.

    Robust to per-school layout differences: each school's summer columns sit at
    different absolute positions (e.g. AE..AI on the 39-col Jasper tabs, AD..AH
    on the 29-col AFMS tab), so we bring each school's summer block to a fixed
    output position via a per-school horizontal join {core, summer}. The summer
    flag is then always output column 15, which one QUERY filters on.
    """
    blocks = []
    for tab in SUMMER_TABS:
        pos = read_summer_positions(sheets, MAP_SPREADSHEET_ID, tab)
        if len(pos) < len(SUMMER_HEADERS):
            print(f"  [SKIP roster] {tab}: summer cols not provisioned")
            continue
        flag = min(pos.values())
        s0 = col_letter(flag)
        s1 = col_letter(flag + len(SUMMER_HEADERS) - 1)
        blocks.append(f"{{'{tab}'!A2:N, '{tab}'!{s0}2:{s1}}}")
    query = (
        "=QUERY({"
        + "; ".join(blocks)
        + '}, "where Col15 = true order by Col3, Col5", 0)'
    )

    meta = (
        sheets.spreadsheets()
        .get(
            spreadsheetId=MAP_SPREADSHEET_ID, fields="sheets.properties(title,sheetId)"
        )
        .execute()
    )
    existing = {
        s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]
    }
    if ROSTER_TAB in existing:
        gid = existing[ROSTER_TAB]
        sheets.spreadsheets().values().clear(
            spreadsheetId=MAP_SPREADSHEET_ID, range=f"'{ROSTER_TAB}'!A:AZ"
        ).execute()
    else:
        resp = (
            sheets.spreadsheets()
            .batchUpdate(
                spreadsheetId=MAP_SPREADSHEET_ID,
                body={
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": ROSTER_TAB,
                                    "index": 1,
                                    "gridProperties": {
                                        "rowCount": 2000,
                                        "columnCount": 24,
                                        "frozenRowCount": 1,
                                    },
                                }
                            }
                        }
                    ]
                },
            )
            .execute()
        )
        gid = resp["replies"][0]["addSheet"]["properties"]["sheetId"]

    header = CORE_HEADERS + SUMMER_HEADERS
    sheets.spreadsheets().values().batchUpdate(
        spreadsheetId=MAP_SPREADSHEET_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": f"'{ROSTER_TAB}'!A1", "values": [header]},
                {"range": f"'{ROSTER_TAB}'!A2", "values": [[query]]},
            ],
        },
    ).execute()
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=MAP_SPREADSHEET_ID,
        body={
            "requests": [
                {
                    "repeatCell": {
                        "range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1},
                        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                        "fields": "userEnteredFormat.textFormat.bold",
                    }
                }
            ]
        },
    ).execute()
    print(
        f"  [OK] '{ROSTER_TAB}' rebuilt: {len(blocks)} school(s), core A:N + 5 summer cols"
    )


def main():
    start = time.time()
    print("=" * 70)
    print("  SUMMER SCHOOL COLUMN PROVISIONING (v2.9.1)")
    print("=" * 70)
    sheets = build_sheets_service()
    for tab in SUMMER_TABS:
        isr_id = ISR_CONFIG[tab]["isr_id"]
        print(f"\n--- {tab}  (ISR {isr_id[:10]}...) ---")
        sr_pos = provision_sr(sheets, isr_id)
        mr_pos = provision_mr(sheets, isr_id, sr_pos)
        provision_cmr(sheets, tab, isr_id, mr_pos)
    print("\n--- Combined Summer School Roster tab ---")
    build_summer_roster_tab(sheets)
    print(f"\n  Completed in {time.time() - start:.1f}s")
    print("  If the CMR shows #REF! on the new cols, open the CMR once and click")
    print("  'Allow access' per ISR (IMPORTRANGE auth; already granted for these")
    print("  ISR->CMR pairs via the existing A1 + Unenroll imports, so usually none).")


if __name__ == "__main__":
    main()
