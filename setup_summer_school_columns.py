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

# v2.9.5: per-ISR hidden helper tab. The summer flag is keyed to EMAIL here
# (stable AND universal -- every student has an email, including "email-only"
# students whose Student ID is still blank). The MR looks students up in it by
# email, so reordering the (sortable) Student Roster can never misalign the flag.
SUMMER_LIST_TAB = "_SummerList"
SUMMER_LIST_HEADERS = ["email", "grade", "subjects", "teacher_email", "teacher"]

# Campuses running summer school (CMR tab name -> ISR id via ISR_CONFIG).
SUMMER_TABS = [
    "Hardeeville Junior & Senior High School (Dash)",
    "Hardeeville Elementary School (Dash)",
    "Ridgeland Secondary Academy of Excellence (Dash)",
    "Allendale Fairfax Middle School (Dash)",  # v2.9.1
    "Ridgeland Elementary School (Dash)",  # v2.9.2
    "Allendale Fairfax Elementary School (Dash)",  # v2.9.3
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

# v2.9.6: per-CMR helper tab driving the combined-roster highlight + float-to-top.
# Support edits HIGHLIGHT_TAB col A (emails) to control which summer students get
# painted light red AND floated to the TOP of the Summer School Roster. PII-free in
# code: the emails live in the sheet (written operationally), never committed.
HIGHLIGHT_TAB = "_Highlight"
HIGHLIGHT_HEADERS = ["email", "note"]
HIGHLIGHT_COLOR = {"red": 0.9569, "green": 0.8, "blue": 0.8}  # #F4CCCC, light red
# Hidden helper column on the Summer School Roster tab: per-row "is this email in
# _Highlight?" flag the conditional-format rule keys on (CF custom formulas cannot
# reference another sheet, so the lookup is mirrored same-sheet here).
HIGHLIGHT_HELPER_COL = 20  # 0-based -> col U, safely right of the A:S output


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


def provision_summer_list(sheets, isr_id):
    """Ensure the hidden `_SummerList` helper tab exists with headers. This tab
    (student_id, grade, subjects, teacher_email, teacher) is the durable, sort-proof
    source for the summer flag: the MR looks students up here by email, so
    reordering the Student Roster can never misalign the flag. Populated per school
    by the one-time reconcile loader."""
    meta = retry_api(
        lambda: sheets.spreadsheets()
        .get(spreadsheetId=isr_id, fields="sheets.properties(title,sheetId)")
        .execute(),
        label="get props for _SummerList",
    )
    existing = {s["properties"]["title"] for s in meta["sheets"]}
    if SUMMER_LIST_TAB not in existing:
        retry_api(
            lambda: sheets.spreadsheets()
            .batchUpdate(
                spreadsheetId=isr_id,
                body={
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": SUMMER_LIST_TAB,
                                    "hidden": True,
                                    "gridProperties": {
                                        "rowCount": 1000,
                                        "columnCount": 6,
                                    },
                                }
                            }
                        }
                    ]
                },
            )
            .execute(),
            label="add _SummerList tab",
        )
    retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .update(
            spreadsheetId=isr_id,
            range=f"'{SUMMER_LIST_TAB}'!A1",
            valueInputOption="RAW",
            body={"values": [SUMMER_LIST_HEADERS]},
        )
        .execute(),
        label="write _SummerList headers",
    )
    print(f"  [OK] {SUMMER_LIST_TAB} ready (student_id-keyed summer source)")


def provision_sr(sheets, isr_id):
    """SR: ensure the 5 summer headers exist, then CLEAR their data rows. v2.9.5:
    the Student Roster is a sortable tab and is NO LONGER the summer source (static
    flags here detached when the roster re-sorted). The flag now lives in
    `_SummerList` + MR lookups. Returns {header: idx}."""
    gid, rows, cols = get_sheet_props(sheets, isr_id, SR_TAB)
    pos = assign_positions(read_header_row(sheets, isr_id, SR_TAB))
    ensure_grid_cols(sheets, isr_id, gid, cols, max(pos.values()) + 1)
    write_headers(sheets, isr_id, SR_TAB, pos)
    c0 = col_letter(min(pos.values()))
    c1 = col_letter(max(pos.values()))
    retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .clear(spreadsheetId=isr_id, range=f"'{SR_TAB}'!{c0}2:{c1}")
        .execute(),
        label="clear stale SR summer data",
    )
    print(f"  [OK] SR summer cols {c0}..{c1} cleared (decoupled; source=_SummerList)")
    return pos


def provision_mr(sheets, isr_id):
    """MR: ensure the 5 summer headers exist; write each as a sort-proof
    ARRAYFORMULA that looks the student up in `_SummerList` by EMAIL (MR col B),
    NOT a mirror of the sortable Student Roster. Email-keyed so it also covers
    students whose Student ID is still blank. Returns {header: idx}.

    _SummerList layout: A=email, B=grade, C=subjects, D=teacher_email, E=teacher.
    """
    gid, rows, cols = get_sheet_props(sheets, isr_id, MR_TAB)
    pos = assign_positions(read_header_row(sheets, isr_id, MR_TAB))
    ensure_grid_cols(sheets, isr_id, gid, cols, max(pos.values()) + 1)
    write_headers(sheets, isr_id, MR_TAB, pos)
    # v2.9.5: clear the whole summer-column data region first. Stale static values
    # left in these columns (e.g. a block of "False" from the old mirror) would
    # block the new ARRAYFORMULA spill with #REF! ("result would overwrite data").
    mc0 = col_letter(min(pos.values()))
    mc1 = col_letter(max(pos.values()))
    retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .clear(spreadsheetId=isr_id, range=f"'{MR_TAB}'!{mc0}2:{mc1}")
        .execute(),
        label="clear MR summer region (unblock arrayformula spill)",
    )
    lk = f"'{SUMMER_LIST_TAB}'!$A$2:$E"
    ids = f"'{SUMMER_LIST_TAB}'!$A$2:$A"
    formula_for = {
        SUMMER_HEADERS[
            0
        ]: f'=ARRAYFORMULA(IF(B2:B="","",IF(ISNUMBER(MATCH(B2:B,{ids},0)),TRUE,FALSE)))',
        SUMMER_HEADERS[
            1
        ]: f'=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,{lk},4,FALSE),"")))',
        SUMMER_HEADERS[
            2
        ]: f'=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,{lk},5,FALSE),"")))',
        SUMMER_HEADERS[
            3
        ]: f'=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,{lk},2,FALSE),"")))',
        SUMMER_HEADERS[
            4
        ]: f'=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,{lk},3,FALSE),"")))',
    }
    data = [
        {"range": f"'{MR_TAB}'!{col_letter(pos[h])}2", "values": [[formula_for[h]]]}
        for h in SUMMER_HEADERS
    ]
    retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .batchUpdate(
            spreadsheetId=isr_id,
            body={"valueInputOption": "USER_ENTERED", "data": data},
        )
        .execute(),
        label="write MR summer lookups",
    )
    set_plain_number(sheets, isr_id, gid, pos[SUMMER_HEADERS[3]], max(rows, 1200))
    print("  [OK] MR summer cols = _SummerList lookups keyed on email (sort-proof)")
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


def provision_highlight_tab(sheets):
    """Ensure the CMR has the hidden _Highlight helper tab. Col A = student emails
    to paint light red AND float to the top of the Summer School Roster; col B =
    free-text note for support. Idempotent: creates + (re)writes headers, NEVER
    clears existing rows (this is support-managed data)."""
    meta = retry_api(
        lambda: sheets.spreadsheets()
        .get(
            spreadsheetId=MAP_SPREADSHEET_ID, fields="sheets.properties(title,sheetId)"
        )
        .execute(),
        label="get CMR props (_Highlight)",
    )
    existing = {
        s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]
    }
    if HIGHLIGHT_TAB in existing:
        gid = existing[HIGHLIGHT_TAB]
    else:
        resp = retry_api(
            lambda: sheets.spreadsheets()
            .batchUpdate(
                spreadsheetId=MAP_SPREADSHEET_ID,
                body={
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": HIGHLIGHT_TAB,
                                    "hidden": True,
                                    "gridProperties": {
                                        "rowCount": 200,
                                        "columnCount": 2,
                                    },
                                }
                            }
                        }
                    ]
                },
            )
            .execute(),
            label="add _Highlight tab",
        )
        gid = resp["replies"][0]["addSheet"]["properties"]["sheetId"]
    retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .update(
            spreadsheetId=MAP_SPREADSHEET_ID,
            range=f"'{HIGHLIGHT_TAB}'!A1:B1",
            valueInputOption="RAW",
            body={"values": [HIGHLIGHT_HEADERS]},
        )
        .execute(),
        label="write _Highlight headers",
    )
    print(
        f"  [OK] '{HIGHLIGHT_TAB}' helper ready (hidden; col A = highlight/top emails)"
    )
    return gid


def build_summer_roster_tab(sheets):
    """(Re)build the combined ROSTER_TAB on the CMR: every Summer School = TRUE
    row from all SUMMER_TABS, normalized to core A:N + the 5 summer columns.

    Robust to per-school layout differences: each school's summer columns sit at
    different absolute positions (e.g. AE..AI on the 39-col Jasper tabs, AD..AH
    on the 29-col AFMS tab), so we bring each school's summer block to a fixed
    output position via a per-school horizontal join {core, summer}. The summer
    flag is then always output column 15, which one QUERY filters on.
    """
    hl_gid = provision_highlight_tab(sheets)
    ncol = len(CORE_HEADERS) + len(SUMMER_HEADERS)  # 19 output cols (A:S)
    sort_col = ncol + 1  # Col20 = per-row float-to-top key (0 = in _Highlight)
    blocks = []
    for tab in SUMMER_TABS:
        pos = read_summer_positions(sheets, MAP_SPREADSHEET_ID, tab)
        if len(pos) < len(SUMMER_HEADERS):
            print(f"  [SKIP roster] {tab}: summer cols not provisioned")
            continue
        flag = min(pos.values())
        s0 = col_letter(flag)
        s1 = col_letter(flag + len(SUMMER_HEADERS) - 1)
        # 3rd sub-array = float-to-top key: 0 when this row's email (col B) is in
        # _Highlight, else 1. Keyed on email so it survives any re-sort.
        topkey = (
            f'ARRAYFORMULA(IF(\'{tab}\'!B2:B="","",'
            f"IF(COUNTIF('{HIGHLIGHT_TAB}'!$A$2:$A,'{tab}'!B2:B)>0,0,1)))"
        )
        blocks.append(f"{{'{tab}'!A2:N, '{tab}'!{s0}2:{s1}, {topkey}}}")
    sel = ", ".join(
        f"Col{i}" for i in range(1, ncol + 1)
    )  # Col1..Col19 (drop sort key)
    query = (
        "=QUERY({"
        + "; ".join(blocks)
        + '}, "select '
        + sel
        + " where Col15 = true order by Col"
        + str(sort_col)
        + ', Col3, Col5", 0)'
    )

    meta = (
        sheets.spreadsheets()
        .get(
            spreadsheetId=MAP_SPREADSHEET_ID,
            fields="sheets.properties(title,sheetId,gridProperties.rowCount)",
        )
        .execute()
    )
    existing = {
        s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]
    }
    rowcounts = {
        s["properties"]["title"]: s["properties"]
        .get("gridProperties", {})
        .get("rowCount", 0)
        for s in meta["sheets"]
    }
    if ROSTER_TAB in existing:
        gid = existing[ROSTER_TAB]
        sheets.spreadsheets().values().clear(
            spreadsheetId=MAP_SPREADSHEET_ID, range=f"'{ROSTER_TAB}'!A:AZ"
        ).execute()
        # Grow the grid so the QUERY spill + CF have headroom: the CF range is
        # clamped to the grid row count, so a small grid would cap where the
        # highlight paints. Only ever grows (never shrinks -> no data loss).
        if rowcounts.get(ROSTER_TAB, 0) < 2000:
            sheets.spreadsheets().batchUpdate(
                spreadsheetId=MAP_SPREADSHEET_ID,
                body={
                    "requests": [
                        {
                            "updateSheetProperties": {
                                "properties": {
                                    "sheetId": gid,
                                    "gridProperties": {"rowCount": 2000},
                                },
                                "fields": "gridProperties.rowCount",
                            }
                        }
                    ]
                },
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
    hcol = col_letter(HIGHLIGHT_HELPER_COL)
    # Same-sheet helper: per output row, TRUE when its email (col B) is in _Highlight.
    # The conditional-format rule keys on this (CF formulas can't cross sheets).
    helper = (
        '=ARRAYFORMULA(IF($B2:$B="","",'
        f"COUNTIF('{HIGHLIGHT_TAB}'!$A$2:$A,$B2:$B)>0))"
    )
    sheets.spreadsheets().values().batchUpdate(
        spreadsheetId=MAP_SPREADSHEET_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": f"'{ROSTER_TAB}'!A1", "values": [header]},
                {"range": f"'{ROSTER_TAB}'!A2", "values": [[query]]},
                {"range": f"'{ROSTER_TAB}'!{hcol}1", "values": [["_highlight_flag"]]},
                {"range": f"'{ROSTER_TAB}'!{hcol}2", "values": [[helper]]},
            ],
        },
    ).execute()
    # Idempotent CF: delete any existing rules on the tab, then add ours once.
    cf_meta = retry_api(
        lambda: sheets.spreadsheets()
        .get(
            spreadsheetId=MAP_SPREADSHEET_ID,
            ranges=[ROSTER_TAB],
            fields="sheets(properties.sheetId,conditionalFormats)",
        )
        .execute(),
        label="read roster CF rules",
    )
    n_cf = 0
    for s in cf_meta.get("sheets", []):
        if s["properties"]["sheetId"] == gid:
            n_cf = len(s.get("conditionalFormats", []))
    requests = [
        {
            "repeatCell": {
                "range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1},
                "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                "fields": "userEnteredFormat.textFormat.bold",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": gid,
                    "dimension": "COLUMNS",
                    "startIndex": HIGHLIGHT_HELPER_COL,
                    "endIndex": HIGHLIGHT_HELPER_COL + 1,
                },
                "properties": {"hiddenByUser": True},
                "fields": "hiddenByUser",
            }
        },
    ]
    requests += [
        {"deleteConditionalFormatRule": {"sheetId": gid, "index": 0}}
        for _ in range(n_cf)
    ]
    requests.append(
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": gid,
                            "startRowIndex": 1,
                            "endRowIndex": 2000,
                            "startColumnIndex": 0,
                            "endColumnIndex": ncol,
                        }
                    ],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": f"=${hcol}2=TRUE"}],
                        },
                        "format": {"backgroundColor": HIGHLIGHT_COLOR},
                    },
                },
                "index": 0,
            }
        }
    )
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=MAP_SPREADSHEET_ID, body={"requests": requests}
    ).execute()
    print(
        f"  [OK] '{ROSTER_TAB}' rebuilt: {len(blocks)} school(s); _Highlight members "
        f"floated to top + painted light red"
    )


def main():
    start = time.time()
    print("=" * 70)
    print("  SUMMER SCHOOL PROVISIONING (v2.9.5: _SummerList + MR student_id lookups)")
    print("=" * 70)
    sheets = build_sheets_service()
    for tab in SUMMER_TABS:
        isr_id = ISR_CONFIG[tab]["isr_id"]
        print(f"\n--- {tab}  (ISR {isr_id[:10]}...) ---")
        provision_summer_list(sheets, isr_id)
        provision_sr(sheets, isr_id)
        mr_pos = provision_mr(sheets, isr_id)
        provision_cmr(sheets, tab, isr_id, mr_pos)
    print("\n--- Combined Summer School Roster tab ---")
    build_summer_roster_tab(sheets)
    print(f"\n  Completed in {time.time() - start:.1f}s")
    print("  Next: run the reconcile loader per school to populate each _SummerList.")
    print("  (If the CMR shows #REF! on the summer cols, open the CMR once and click")
    print("   'Allow access' per ISR -- usually already granted via existing imports.)")


if __name__ == "__main__":
    main()
