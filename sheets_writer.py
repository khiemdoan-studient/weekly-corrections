"""Google Sheets API writer for the Weekly Corrections tool.

Layout (all 3 visible sheets):
  Row 1: Title (merged, navy dark, 20pt bold white)
  Row 2: Caption with User Guide hyperlink (merged, navy med, italic 12pt grey)
  Row 3: Spacer (5px, dark)
  Row 4: Filter labels + Sort By label (merged pairs, dark navy)
  Row 5: Dropdown values + Sort By dropdown (merged pairs, teal bg, data validation)
  Row 6: Column headers (navy, bold white)
  Row 7+: SORT(QUERY(...)) formula output (filtered + sorted from hidden data tabs)

Filtering and sorting both done by formulas — NOT by Apps Script.
"""

import time

from config import (
    OUTPUT_SPREADSHEET_ID,
    TAB_CORRECTED,
    TAB_SIS,
    TAB_APPROVED,
    OUTPUT_FIELDS,
)

# ── Colour constants (matching dashboard pipeline) ─────────────────────────


def _rgb(h):
    h = h.lstrip("#")
    return {
        "red": int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue": int(h[4:6], 16) / 255,
    }


NAVY_DARK = _rgb("0F1B33")
NAVY_MED = _rgb("263D66")
FILTER_BG = _rgb("1E3A5F")
WHITE = _rgb("FFFFFF")
GREY_LABEL = _rgb("94A3B8")
GREY_BORDER = _rgb("CBD5E1")
LINK_BLUE = _rgb("93C5FD")
ALT_ROW = _rgb("EDF2F7")
DROPDOWN_BG = _rgb("2D4A7A")  # lighter blue for dropdown row (row 5)
RED_HDR = _rgb("7F1D1D")  # dark red for Mismatch Summary header
RED_LIGHT = _rgb("FEE2E2")  # light red for Mismatch Summary data cells

GUIDE_URL = (
    "https://docs.google.com/document/d/1O1WEAHSttdNVRUa_CoQ3T6w4QEFPyLz5FDdM2IMHEu4"
)

CAPTION_TEXT = (
    "Implementation Managers are responsible for checking off the corrections "
    "to make for their own schools  |  Ticket will be automatically submitted "
    "once a week with the checked off boxes  |  User Guide"
)

# Sort options — must match QUERY output column order for MATCH() to work
SORT_OPTS_SHEET1 = [  # _CorrData 13 cols
    "Campus",
    "Grade",
    "Level",
    "First Name",
    "Last Name",
    "Email",
    "Student Group",
    "Guide First Name",
    "Guide Last Name",
    "Guide Email",
    "Student_ID",
    "External Student ID",
    "Mismatch Summary",
]
SORT_OPTS_SHEET2 = [  # _SISData 12 cols
    "Campus",
    "Grade",
    "Level",
    "First Name",
    "Last Name",
    "Email",
    "Student Group",
    "Guide First Name",
    "Guide Last Name",
    "Guide Email",
    "Student_ID",
    "External Student ID",
]
SORT_OPTS_SHEET3 = [  # _ApprovedData 13 cols
    "Date Approved",
    "Campus",
    "Grade",
    "Level",
    "First Name",
    "Last Name",
    "Email",
    "Student Group",
    "Guide First Name",
    "Guide Last Name",
    "Guide Email",
    "Student_ID",
    "External Student ID",
]

# ── API helpers ────────────────────────────────────────────────────────────


def _retry_api(fn, max_retries=3, delay=5):
    for attempt in range(max_retries):
        try:
            return fn()
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            print(f"     Warning: API call failed (attempt {attempt+1}): {e}")
            time.sleep(delay * (attempt + 1))


def _ensure_tab_exists(sheets_service, spreadsheet_id, tab_name):
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


def _clear_banding(sheets_service, spreadsheet_id, sheet_ids):
    resp = _retry_api(
        lambda: sheets_service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties.sheetId,bandedRanges)",
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


# ── Sheets API request builders ───────────────────────────────────────────


def _mg(sid, r1, r2, c1, c2):
    return {
        "mergeCells": {
            "range": {
                "sheetId": sid,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": c1,
                "endColumnIndex": c2,
            },
            "mergeType": "MERGE_ALL",
        }
    }


def _rc(sid, r1, r2, c1, c2, cf):
    return {
        "repeatCell": {
            "range": {
                "sheetId": sid,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": c1,
                "endColumnIndex": c2,
            },
            "cell": {"userEnteredFormat": cf},
            "fields": "userEnteredFormat",
        }
    }


def _cw(sid, c, w):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sid,
                "dimension": "COLUMNS",
                "startIndex": c,
                "endIndex": c + 1,
            },
            "properties": {"pixelSize": w},
            "fields": "pixelSize",
        }
    }


def _rh(sid, r, h):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sid,
                "dimension": "ROWS",
                "startIndex": r,
                "endIndex": r + 1,
            },
            "properties": {"pixelSize": h},
            "fields": "pixelSize",
        }
    }


def _dd(sid, r, c, lr):
    v = f"={lr}" if "!" in lr else f"='{lr}'"
    return {
        "setDataValidation": {
            "range": {
                "sheetId": sid,
                "startRowIndex": r,
                "endRowIndex": r + 1,
                "startColumnIndex": c,
                "endColumnIndex": c + 1,
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_RANGE",
                    "values": [{"userEnteredValue": v}],
                },
                "showCustomUi": True,
                "strict": False,
            },
        }
    }


def _cf(bg=None, fg=None, bold=False, italic=False, sz=10, ha="LEFT", va="MIDDLE"):
    f = {
        "textFormat": {
            "fontFamily": "Arial",
            "fontSize": sz,
            "bold": bold,
            "italic": italic,
        },
        "horizontalAlignment": ha,
        "verticalAlignment": va,
    }
    if fg:
        f["textFormat"]["foregroundColorStyle"] = {"rgbColor": fg}
    if bg:
        f["backgroundColorStyle"] = {"rgbColor": bg}
    return f


def _subtitle_with_link(sid, row, text, nc):
    """Create subtitle cell with clickable 'User Guide' hyperlink."""
    link_start = text.find("User Guide")
    link_end = link_start + len("User Guide")
    runs = [
        {"startIndex": 0},
        {
            "startIndex": link_start,
            "format": {
                "link": {"uri": GUIDE_URL},
                "foregroundColorStyle": {"rgbColor": LINK_BLUE},
                "underline": True,
            },
        },
    ]
    if link_end < len(text):
        runs.append({"startIndex": link_end})
    return {
        "updateCells": {
            "range": {
                "sheetId": sid,
                "startRowIndex": row,
                "endRowIndex": row + 1,
                "startColumnIndex": 0,
                "endColumnIndex": 1,
            },
            "rows": [
                {
                    "values": [
                        {
                            "userEnteredValue": {"stringValue": text},
                            "textFormatRuns": runs,
                        }
                    ]
                }
            ],
            "fields": "userEnteredValue,textFormatRuns",
        }
    }


# ── Compute unique filter values ──────────────────────────────────────────


def _compute_unique_values(corrections_map, corrections_sis):
    unique = {
        "campus": set(),
        "grade": set(),
        "level": set(),
        "student_group": set(),
        "guide_email": set(),
    }
    for rec in corrections_map + corrections_sis:
        for key, field in [
            ("campus", "Campus"),
            ("grade", "Grade"),
            ("level", "Level"),
            ("student_group", "Student Group"),
            ("guide_email", "Guide Email"),
        ]:
            val = rec.get(field, "")
            if val and val != "NOT FOUND IN SIS":
                unique[key].add(str(val).strip())
    return {k: ["All"] + sorted(v) for k, v in unique.items()}


# ── QUERY formula builders (with SORT wrapper) ───────────────────────────
# All filter values wrapped with SUBSTITUTE(cell,"'","''") to escape single
# quotes in student/campus names (e.g. O'Brien) that would break QUERY syntax.


def _sq(cell):
    """Wrap a cell reference with single-quote escaping for QUERY strings."""
    return f'SUBSTITUTE({cell},"\'","\'\'")'


def _build_sorted_query_sheet1(sort_list_range):
    """SORT(QUERY()) for Sheet 1. Filters in B5,D5,F5,H5,J5. Sort By in L5."""
    src = "'_CorrData'!A:M"
    # _CorrData cols: 1=Campus 2=Grade 3=Level 4=FirstName 5=LastName
    #   6=Email 7=StudentGroup 8=GuideFirst 9=GuideLast 10=GuideEmail
    #   11=StudentID 12=ExtStudentID 13=MismatchSummary
    q = (
        f"=IFERROR(SORT(QUERY({src}, "
        '"SELECT * WHERE 1=1"'
        f'& IF($B$5="All", "", " AND Col1=\'" & {_sq("$B$5")} & "\'")'  # Campus
        f'& IF($D$5="All", "", " AND Col2=\'" & {_sq("$D$5")} & "\'")'  # Grade
        f'& IF($F$5="All", "", " AND Col3=\'" & {_sq("$F$5")} & "\'")'  # Level
        f'& IF($H$5="All", "", " AND Col7=\'" & {_sq("$H$5")} & "\'")'  # Student Group
        f'& IF($J$5="All", "", " AND Col10=\'" & {_sq("$J$5")} & "\'")'  # Guide Email
        f", 0), MATCH($L$5, {sort_list_range}, 0), "
        'IF(OR($L$5="Grade"), FALSE, TRUE)), "")'
    )
    return q


def _build_sorted_query_sheet2(sort_list_range):
    """SORT(QUERY()) for Sheet 2. Filters in A5,C5,E5,G5,I5. Sort By in K5."""
    src = "'_SISData'!A:L"
    # _SISData cols: same as _CorrData minus MismatchSummary (12 cols)
    q = (
        f"=IFERROR(SORT(QUERY({src}, "
        '"SELECT * WHERE 1=1"'
        f'& IF($A$5="All", "", " AND Col1=\'" & {_sq("$A$5")} & "\'")'  # Campus
        f'& IF($C$5="All", "", " AND Col2=\'" & {_sq("$C$5")} & "\'")'  # Grade
        f'& IF($E$5="All", "", " AND Col3=\'" & {_sq("$E$5")} & "\'")'  # Level
        f'& IF($G$5="All", "", " AND Col7=\'" & {_sq("$G$5")} & "\'")'  # Student Group
        f'& IF($I$5="All", "", " AND Col10=\'" & {_sq("$I$5")} & "\'")'  # Guide Email
        f", 0), MATCH($K$5, {sort_list_range}, 0), "
        'IF(OR($K$5="Grade"), FALSE, TRUE)), "")'
    )
    return q


def _build_sorted_query_sheet3(sort_list_range):
    """SORT(QUERY()) for Sheet 3. Filters in A5,C5,E5,G5,I5. Sort By in K5.
    _ApprovedData cols: 1=Date 2=Campus 3=Grade 4=Level 5=FirstName 6=LastName
      7=Email 8=StudentGroup 9=GuideFirst 10=GuideLast 11=GuideEmail
      12=StudentID 13=ExtStudentID
    """
    src = "'_ApprovedData'!A:M"
    q = (
        f"=IFERROR(SORT(QUERY({src}, "
        '"SELECT * WHERE 1=1"'
        f'& IF($A$5="All", "", " AND Col2=\'" & {_sq("$A$5")} & "\'")'  # Campus=Col2
        f'& IF($C$5="All", "", " AND Col3=\'" & {_sq("$C$5")} & "\'")'  # Grade=Col3
        f'& IF($E$5="All", "", " AND Col4=\'" & {_sq("$E$5")} & "\'")'  # Level=Col4
        f'& IF($G$5="All", "", " AND Col8=\'" & {_sq("$G$5")} & "\'")'  # StudentGroup=Col8
        f'& IF($I$5="All", "", " AND Col11=\'" & {_sq("$I$5")} & "\'")'  # GuideEmail=Col11
        f", 0), MATCH($K$5, {sort_list_range}, 0), "
        'IF(OR($K$5="Date Approved"), FALSE, TRUE)), "")'
    )
    return q


# ══════════════════════════════════════════════════════════════════════════
# MAIN WRITE FUNCTION
# ══════════════════════════════════════════════════════════════════════════


def write_corrections(sheets_service, corrections_map, corrections_sis):
    sid = OUTPUT_SPREADSHEET_ID
    print(f"\n  Writing {len(corrections_map)} correction rows...")

    # ── Ensure all tabs exist ─────────────────────────────────────────
    sheet1_id = _ensure_tab_exists(sheets_service, sid, TAB_CORRECTED)
    sheet2_id = _ensure_tab_exists(sheets_service, sid, TAB_SIS)
    sheet3_id = _ensure_tab_exists(sheets_service, sid, TAB_APPROVED)
    corr_data_id = _ensure_tab_exists(sheets_service, sid, "_CorrData")
    sis_data_id = _ensure_tab_exists(sheets_service, sid, "_SISData")
    approved_data_id = _ensure_tab_exists(sheets_service, sid, "_ApprovedData")
    lists_id = _ensure_tab_exists(sheets_service, sid, "_Lists")

    # ── Clear banding + data ──────────────────────────────────────────
    _clear_banding(sheets_service, sid, [sheet1_id, sheet2_id, sheet3_id])
    for tab in [
        TAB_CORRECTED,
        TAB_SIS,
        TAB_APPROVED,
        "_CorrData",
        "_SISData",
        "_Lists",
    ]:
        _retry_api(
            lambda t=tab: sheets_service.spreadsheets()
            .values()
            .clear(spreadsheetId=sid, range=f"'{t}'!A:Z")
            .execute()
        )
    # NOTE: _ApprovedData is NOT cleared — it's cumulative (managed by Apps Script)

    if not corrections_map:
        print("  No mismatches found — sheets cleared.")
        _retry_api(
            lambda: sheets_service.spreadsheets()
            .values()
            .update(
                spreadsheetId=sid,
                range=f"'{TAB_CORRECTED}'!A1",
                valueInputOption="RAW",
                body={"values": [["No mismatches found."]]},
            )
            .execute()
        )
        return

    # ── Compute unique values ─────────────────────────────────────────
    unique = _compute_unique_values(corrections_map, corrections_sis)

    # ══════════════════════════════════════════════════════════════════
    # WRITE HIDDEN TABS
    # ══════════════════════════════════════════════════════════════════
    print("  Writing hidden data tabs...")

    # _CorrData: 13 cols (Campus..MismatchSummary), NO header
    corr_rows = []
    for rec in corrections_map:
        row = [rec.get(f, "") for f in OUTPUT_FIELDS]
        row.append(rec.get("mismatch_summary", ""))
        corr_rows.append(row)
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range="'_CorrData'!A1",
            valueInputOption="RAW",
            body={"values": corr_rows},
        )
        .execute()
    )

    # _SISData: 12 cols, NO header
    sis_rows = []
    for rec in corrections_sis:
        sis_rows.append([rec.get(f, "") for f in OUTPUT_FIELDS])
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range="'_SISData'!A1",
            valueInputOption="RAW",
            body={"values": sis_rows},
        )
        .execute()
    )

    # _Lists: 8 columns — 5 filter values (A-E) + 3 sort options (F-H)
    n = {k: len(v) for k, v in unique.items()}
    s1 = SORT_OPTS_SHEET1
    s2 = SORT_OPTS_SHEET2
    s3 = SORT_OPTS_SHEET3
    max_len = max(max(n.values()), len(s1), len(s2), len(s3))
    lists_header = [
        "campus",
        "grade",
        "level",
        "student_group",
        "guide_email",
        "sort_sheet1",
        "sort_sheet2",
        "sort_sheet3",
    ]
    lists_rows = [lists_header]
    filter_keys = ["campus", "grade", "level", "student_group", "guide_email"]
    for i in range(max_len):
        row = []
        for key in filter_keys:
            vals = unique[key]
            row.append(vals[i] if i < len(vals) else "")
        row.append(s1[i] if i < len(s1) else "")
        row.append(s2[i] if i < len(s2) else "")
        row.append(s3[i] if i < len(s3) else "")
        lists_rows.append(row)

    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range="'_Lists'!A1",
            valueInputOption="RAW",
            body={"values": lists_rows},
        )
        .execute()
    )

    list_ranges = {
        "campus": f"_Lists!A2:A{n['campus'] + 1}",
        "grade": f"_Lists!B2:B{n['grade'] + 1}",
        "level": f"_Lists!C2:C{n['level'] + 1}",
        "student_group": f"_Lists!D2:D{n['student_group'] + 1}",
        "guide_email": f"_Lists!E2:E{n['guide_email'] + 1}",
        "sort1": f"_Lists!F2:F{len(s1) + 1}",
        "sort2": f"_Lists!G2:G{len(s2) + 1}",
        "sort3": f"_Lists!H2:H{len(s3) + 1}",
    }

    # ══════════════════════════════════════════════════════════════════
    # WRITE VISIBLE SHEETS
    # ══════════════════════════════════════════════════════════════════
    print("  Writing visible sheets...")
    vals_raw = {}
    vals_entered = {}

    NC1 = 14  # checkbox + 12 fields + mismatch summary
    NC2 = 12  # 12 fields
    NC3 = 13  # date + 12 fields

    # ── Sheet 1: Corrected Roster Info ────────────────────────────────
    vals_raw[f"'{TAB_CORRECTED}'!A1"] = [["Corrected Roster Info"]]
    # Row 4: filter labels (B,D,F,H,J) + Sort By (L)
    s1_r4 = [""] * NC1
    s1_r4[1] = "Campus"
    s1_r4[3] = "Grade"
    s1_r4[5] = "Level"
    s1_r4[7] = "Student Group"
    s1_r4[9] = "Guide Email"
    s1_r4[11] = "SORT BY"
    vals_raw[f"'{TAB_CORRECTED}'!A4"] = [s1_r4]
    # Row 5: dropdown defaults
    s1_r5 = [""] * NC1
    s1_r5[1] = "All"
    s1_r5[3] = "All"
    s1_r5[5] = "All"
    s1_r5[7] = "All"
    s1_r5[9] = "All"
    s1_r5[11] = "Campus"
    vals_raw[f"'{TAB_CORRECTED}'!A5"] = [s1_r5]
    # Row 6: headers
    vals_raw[f"'{TAB_CORRECTED}'!A6"] = [
        ["\u2713"] + OUTPUT_FIELDS + ["Mismatch Summary"]
    ]
    # Row 7: SORT(QUERY()) formula
    vals_entered[f"'{TAB_CORRECTED}'!B7"] = [
        [_build_sorted_query_sheet1(list_ranges["sort1"])]
    ]

    # ── Sheet 2: Current Roster Info in SIS ───────────────────────────
    vals_raw[f"'{TAB_SIS}'!A1"] = [["Current Roster Info in SIS"]]
    s2_r4 = [""] * NC2
    s2_r4[0] = "Campus"
    s2_r4[2] = "Grade"
    s2_r4[4] = "Level"
    s2_r4[6] = "Student Group"
    s2_r4[8] = "Guide Email"
    s2_r4[10] = "SORT BY"
    vals_raw[f"'{TAB_SIS}'!A4"] = [s2_r4]
    s2_r5 = [""] * NC2
    s2_r5[0] = "All"
    s2_r5[2] = "All"
    s2_r5[4] = "All"
    s2_r5[6] = "All"
    s2_r5[8] = "All"
    s2_r5[10] = "Campus"
    vals_raw[f"'{TAB_SIS}'!A5"] = [s2_r5]
    vals_raw[f"'{TAB_SIS}'!A6"] = [OUTPUT_FIELDS[:]]
    vals_entered[f"'{TAB_SIS}'!A7"] = [
        [_build_sorted_query_sheet2(list_ranges["sort2"])]
    ]

    # ── Sheet 3: Automated Correction List ────────────────────────────
    vals_raw[f"'{TAB_APPROVED}'!A1"] = [["Automated Correction List"]]
    s3_r4 = [""] * NC3
    s3_r4[0] = "Campus"
    s3_r4[2] = "Grade"
    s3_r4[4] = "Level"
    s3_r4[6] = "Student Group"
    s3_r4[8] = "Guide Email"
    s3_r4[10] = "SORT BY"
    vals_raw[f"'{TAB_APPROVED}'!A4"] = [s3_r4]
    s3_r5 = [""] * NC3
    s3_r5[0] = "All"
    s3_r5[2] = "All"
    s3_r5[4] = "All"
    s3_r5[6] = "All"
    s3_r5[8] = "All"
    s3_r5[10] = "Date Approved"
    vals_raw[f"'{TAB_APPROVED}'!A5"] = [s3_r5]
    vals_raw[f"'{TAB_APPROVED}'!A6"] = [["Date Approved"] + OUTPUT_FIELDS]
    vals_entered[f"'{TAB_APPROVED}'!A7"] = [
        [_build_sorted_query_sheet3(list_ranges["sort3"])]
    ]

    # ── Write values (batched — 2 API calls instead of 15+) ─────────
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .batchUpdate(
            spreadsheetId=sid,
            body={
                "valueInputOption": "RAW",
                "data": [{"range": r, "values": v} for r, v in vals_raw.items()],
            },
        )
        .execute()
    )
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .batchUpdate(
            spreadsheetId=sid,
            body={
                "valueInputOption": "USER_ENTERED",
                "data": [{"range": r, "values": v} for r, v in vals_entered.items()],
            },
        )
        .execute()
    )

    # ══════════════════════════════════════════════════════════════════
    # FORMATTING
    # ══════════════════════════════════════════════════════════════════
    print("  Applying formatting...")
    fmt = []

    fmt.extend(
        _format_visible_sheet(
            sheet1_id,
            NC1,
            list_ranges,
            has_checkbox=True,
            sort_col_start=11,
            sort_list_key="sort1",
            num_data_rows=len(corr_rows),
        )
    )
    fmt.extend(
        _format_visible_sheet(
            sheet2_id,
            NC2,
            list_ranges,
            has_checkbox=False,
            sort_col_start=10,
            sort_list_key="sort2",
            num_data_rows=len(sis_rows),
        )
    )
    fmt.extend(
        _format_visible_sheet(
            sheet3_id,
            NC3,
            list_ranges,
            has_checkbox=False,
            sort_col_start=10,
            sort_list_key="sort3",
            num_data_rows=0,
        )
    )

    # Hide data tabs
    for hid in [corr_data_id, sis_data_id, approved_data_id, lists_id]:
        fmt.append(
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": hid, "hidden": True},
                    "fields": "hidden",
                }
            }
        )

    # Date format on Sheet 3 col A (QUERY strips number formatting from _ApprovedData)
    fmt.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet3_id,
                    "startRowIndex": 6,
                    "endRowIndex": 1006,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "DATE_TIME",
                            "pattern": "yyyy-MM-dd HH:mm:ss",
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        }
    )

    # Checkbox data validation on Sheet 1 col A, rows 7+
    fmt.append(
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet1_id,
                    "startRowIndex": 6,
                    "endRowIndex": 6 + len(corr_rows) + 10,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "rule": {"condition": {"type": "BOOLEAN"}, "showCustomUi": True},
            }
        }
    )

    # Execute in batches
    for i in range(0, len(fmt), 200):
        batch = fmt[i : i + 200]
        _retry_api(
            lambda b=batch: sheets_service.spreadsheets()
            .batchUpdate(spreadsheetId=sid, body={"requests": b})
            .execute()
        )

    print(f"  Done — {len(corrections_map)} students written to corrections sheet.")


# ══════════════════════════════════════════════════════════════════════════
# FORMATTING HELPERS
# ══════════════════════════════════════════════════════════════════════════


def _format_visible_sheet(
    sheet_id,
    nc,
    list_ranges,
    has_checkbox,
    sort_col_start,
    sort_list_key,
    num_data_rows,
):
    fmt = []
    offset = 1 if has_checkbox else 0

    # ── Row 0 (title): merged, navy dark, 20pt bold white ─────────────
    fmt.append(_mg(sheet_id, 0, 1, 0, nc))
    fmt.append(_rh(sheet_id, 0, 55))
    fmt.append(
        _rc(
            sheet_id,
            0,
            1,
            0,
            nc,
            _cf(bg=NAVY_DARK, fg=WHITE, bold=True, sz=20, ha="CENTER", va="MIDDLE"),
        )
    )

    # ── Row 1 (caption): merged, navy med, italic 12pt grey + link ────
    fmt.append(_mg(sheet_id, 1, 2, 0, nc))
    fmt.append(_rh(sheet_id, 1, 36))
    fmt.append(
        _rc(
            sheet_id,
            1,
            2,
            0,
            nc,
            _cf(bg=NAVY_MED, fg=GREY_LABEL, italic=True, sz=12, ha="CENTER"),
        )
    )
    fmt.append(_subtitle_with_link(sheet_id, 1, CAPTION_TEXT, nc))

    # ── Row 2 (spacer) ───────────────────────────────────────────────
    fmt.append(_rh(sheet_id, 2, 5))
    fmt.append(_rc(sheet_id, 2, 3, 0, nc, _cf(bg=NAVY_DARK)))

    # ── Rows 3-4 (filter area): dark background ─────────────────────
    fmt.append(_rh(sheet_id, 3, 18))
    fmt.append(_rh(sheet_id, 4, 34))
    fmt.append(_rc(sheet_id, 3, 5, 0, nc, _cf(bg=NAVY_DARK)))

    # ── 5 filter dropdowns (merged pairs) ─────────────────────────────
    filter_keys = ["campus", "grade", "level", "student_group", "guide_email"]
    for i, key in enumerate(filter_keys):
        c1 = offset + i * 2
        c2 = c1 + 2
        if c2 > nc:
            break
        # Merge + format label (row 3)
        fmt.append(_mg(sheet_id, 3, 4, c1, c2))
        fmt.append(
            _rc(
                sheet_id,
                3,
                4,
                c1,
                c2,
                _cf(
                    bg=NAVY_DARK,
                    fg=GREY_LABEL,
                    bold=True,
                    sz=10,
                    ha="CENTER",
                    va="BOTTOM",
                ),
            )
        )
        # Merge + format dropdown (row 4) — lighter blue to stand out
        fmt.append(_mg(sheet_id, 4, 5, c1, c2))
        fmt.append(
            _rc(
                sheet_id,
                4,
                5,
                c1,
                c2,
                _cf(
                    bg=DROPDOWN_BG, fg=WHITE, bold=True, sz=11, ha="CENTER", va="MIDDLE"
                ),
            )
        )
        if key in list_ranges:
            fmt.append(_dd(sheet_id, 4, c1, list_ranges[key]))

    # ── Sort By dropdown (merged, cols sort_col_start to sort_col_start+2) ──
    sc = sort_col_start
    se = min(sc + 3, nc)  # 3 cols wide (L-N for Sheet 1)
    # Label (row 3)
    fmt.append(_mg(sheet_id, 3, 4, sc, se))
    fmt.append(
        _rc(
            sheet_id,
            3,
            4,
            sc,
            se,
            _cf(
                bg=NAVY_DARK, fg=GREY_LABEL, bold=True, sz=10, ha="CENTER", va="BOTTOM"
            ),
        )
    )
    # Dropdown (row 4) — lighter blue to stand out
    fmt.append(_mg(sheet_id, 4, 5, sc, se))
    fmt.append(
        _rc(
            sheet_id,
            4,
            5,
            sc,
            se,
            _cf(bg=DROPDOWN_BG, fg=WHITE, bold=True, sz=11, ha="CENTER", va="MIDDLE"),
        )
    )
    if sort_list_key in list_ranges:
        fmt.append(_dd(sheet_id, 4, sc, list_ranges[sort_list_key]))

    # ── Row 5 (headers): navy bold white ──────────────────────────────
    fmt.append(_rh(sheet_id, 5, 30))
    fmt.append(
        _rc(
            sheet_id,
            5,
            6,
            0,
            nc,
            _cf(bg=FILTER_BG, fg=WHITE, bold=True, sz=10, ha="CENTER", va="MIDDLE"),
        )
    )

    # ── Freeze rows 1-6 ──────────────────────────────────────────────
    fmt.append(
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 6},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        }
    )

    # ── Column widths ─────────────────────────────────────────────────
    if has_checkbox:
        fmt.append(_cw(sheet_id, 0, 30))
    field_widths = [150, 60, 80, 100, 100, 220, 150, 100, 100, 220, 120, 140]
    for i, w in enumerate(field_widths):
        fmt.append(_cw(sheet_id, offset + i, w))
    if has_checkbox:
        fmt.append(_cw(sheet_id, offset + len(field_widths), 200))

    # ── Alternating row colors ────────────────────────────────────────
    end_row = max(6 + num_data_rows + 5, 20)
    fmt.append(
        {
            "addBanding": {
                "bandedRange": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 5,
                        "endRowIndex": end_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": nc,
                    },
                    "rowProperties": {
                        "headerColor": FILTER_BG,
                        "firstBandColor": _rgb("FFFFFF"),
                        "secondBandColor": ALT_ROW,
                    },
                }
            }
        }
    )

    # ── Mismatch Summary column: dark red header, light red data (Sheet 1 only) ──
    if has_checkbox:
        mismatch_col = nc - 1  # last column (index 13 for Sheet 1)
        # Dark red header (row 5)
        fmt.append(
            _rc(
                sheet_id,
                5,
                6,
                mismatch_col,
                mismatch_col + 1,
                _cf(bg=RED_HDR, fg=WHITE, bold=True, sz=10, ha="CENTER", va="MIDDLE"),
            )
        )
        # Light red data cells (row 6 onward)
        fmt.append(
            _rc(
                sheet_id,
                6,
                end_row,
                mismatch_col,
                mismatch_col + 1,
                _cf(bg=RED_LIGHT, sz=10),
            )
        )

    return fmt
