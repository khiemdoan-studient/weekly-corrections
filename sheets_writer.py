"""Google Sheets API writer for the Weekly Corrections tool.

Layout (all 6 visible sheets):
  Row 1: Title (merged, navy dark, 20pt bold white)
  Row 2: Caption with User Guide hyperlink (merged, navy med, italic 12pt grey)
  Row 3: Spacer (5px, dark)
  Row 4: Filter labels + Sort By label (merged pairs, dark navy)
  Row 5: Dropdown values + Sort By dropdown (merged pairs, teal bg, data validation)
  Row 6: Column headers (navy, bold white)
  Row 7+: SORT(QUERY(...)) formula output (filtered + sorted from hidden data tabs)

Visible sheets:
  1. Corrected Roster Info    — accept/reject + QUERY from _CorrData (all mismatch types)
  2. Current Roster Info in SIS — QUERY from _SISData
  3. Automated Correction List — QUERY from _ApprovedData (field mismatches)
  4. Roster Additions          — QUERY from _AdditionsData
  5. Roster Unenrollments      — QUERY from _UnenrollData
  6. Rejected Changes          — QUERY from _RejectedData + Reason for Rejection

Filtering and sorting both done by formulas — NOT by Apps Script.
"""

import time

from config import (
    OUTPUT_SPREADSHEET_ID,
    TAB_CORRECTED,
    TAB_SIS,
    TAB_APPROVED,
    TAB_ADDITIONS,
    TAB_UNENROLL,
    TAB_REJECTED,
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
GREEN_LIGHT = _rgb("D4EDDA")  # light green for Roster Addition
YELLOW_MM = _rgb("FFF3CD")  # yellow for field mismatches
YELLOW_LIGHT = _rgb("FFFDE7")  # light yellow for Unenrolling
ACCEPT_BG = _rgb("D4EDDA")  # light green for Accept Changes column
REJECT_BG = _rgb("FEE2E2")  # light red for Reject Changes column

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
SORT_OPTS_SHEET3 = [  # _ApprovedData 14 cols
    "Date Approved",
    "Mismatch Summary",
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
SORT_OPTS_SHEET4 = list(SORT_OPTS_SHEET3)  # _AdditionsData — same layout
SORT_OPTS_SHEET5 = list(SORT_OPTS_SHEET3)  # _UnenrollData — same layout
SORT_OPTS_SHEET6 = [  # _RejectedData 14 cols
    "Date Rejected",
    "Mismatch Summary",
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


def _ensure_all_tabs(sheets_service, spreadsheet_id, tab_names):
    """Ensure all tabs exist in a single batched API call. Returns dict of name -> sheetId."""
    resp = _retry_api(
        lambda: sheets_service.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
        .execute()
    )
    existing = {}
    for sheet in resp.get("sheets", []):
        props = sheet["properties"]
        existing[props["title"]] = props["sheetId"]

    missing = [name for name in tab_names if name not in existing]
    if missing:
        requests = [{"addSheet": {"properties": {"title": name}}} for name in missing]
        result = _retry_api(
            lambda: sheets_service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests})
            .execute()
        )
        for reply in result.get("replies", []):
            if "addSheet" in reply:
                props = reply["addSheet"]["properties"]
                existing[props["title"]] = props["sheetId"]
                print(
                    f"     Created tab '{props['title']}' (sheetId={props['sheetId']})"
                )

    return {name: existing[name] for name in tab_names}


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
    """SORT(QUERY()) for Sheet 1. Filters in C5,E5,G5,I5,K5. Sort By in M5.
    (Offset by 2 for accept/reject checkbox columns A-B.)"""
    src = "'_CorrData'!A:M"
    # _CorrData cols: 1=Campus 2=Grade 3=Level 4=FirstName 5=LastName
    #   6=Email 7=StudentGroup 8=GuideFirst 9=GuideLast 10=GuideEmail
    #   11=StudentID 12=ExtStudentID 13=MismatchSummary
    q = (
        f"=IFERROR(SORT(QUERY({src}, "
        '"SELECT * WHERE 1=1"'
        f'& IF($C$5="All", "", " AND Col1=\'" & {_sq("$C$5")} & "\'")'  # Campus
        f'& IF($E$5="All", "", " AND Col2=\'" & {_sq("$E$5")} & "\'")'  # Grade
        f'& IF($G$5="All", "", " AND Col3=\'" & {_sq("$G$5")} & "\'")'  # Level
        f'& IF($I$5="All", "", " AND Col7=\'" & {_sq("$I$5")} & "\'")'  # Student Group
        f'& IF($K$5="All", "", " AND Col10=\'" & {_sq("$K$5")} & "\'")'  # Guide Email
        f", 0), MATCH($M$5, {sort_list_range}, 0), "
        'IF(OR($M$5="Grade"), FALSE, TRUE)), "")'
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
    _ApprovedData cols: 1=Date 2=MismatchSummary 3=Campus 4=Grade 5=Level
      6=FirstName 7=LastName 8=Email 9=StudentGroup 10=GuideFirst 11=GuideLast
      12=GuideEmail 13=StudentID 14=ExtStudentID
    """
    src = "'_ApprovedData'!A:N"
    q = (
        f"=IFERROR(SORT(QUERY({src}, "
        '"SELECT * WHERE 1=1"'
        f'& IF($A$5="All", "", " AND Col3=\'" & {_sq("$A$5")} & "\'")'  # Campus=Col3
        f'& IF($C$5="All", "", " AND Col4=\'" & {_sq("$C$5")} & "\'")'  # Grade=Col4
        f'& IF($E$5="All", "", " AND Col5=\'" & {_sq("$E$5")} & "\'")'  # Level=Col5
        f'& IF($G$5="All", "", " AND Col9=\'" & {_sq("$G$5")} & "\'")'  # StudentGroup=Col9
        f'& IF($I$5="All", "", " AND Col12=\'" & {_sq("$I$5")} & "\'")'  # GuideEmail=Col12
        f", 0), MATCH($K$5, {sort_list_range}, 0), "
        'IF(OR($K$5="Date Approved"), FALSE, TRUE)), "")'
    )
    return q


def _build_sorted_query_sheet4(sort_list_range):
    """SORT(QUERY()) for Sheet 4 (Roster Additions). Source: _AdditionsData.
    Same col mapping as _ApprovedData: 1=Date 2=Mismatch ... 14=ExtStudentID
    """
    src = "'_AdditionsData'!A:N"
    q = (
        f"=IFERROR(SORT(QUERY({src}, "
        '"SELECT * WHERE 1=1"'
        f'& IF($A$5="All", "", " AND Col3=\'" & {_sq("$A$5")} & "\'")'  # Campus=Col3
        f'& IF($C$5="All", "", " AND Col4=\'" & {_sq("$C$5")} & "\'")'  # Grade=Col4
        f'& IF($E$5="All", "", " AND Col5=\'" & {_sq("$E$5")} & "\'")'  # Level=Col5
        f'& IF($G$5="All", "", " AND Col9=\'" & {_sq("$G$5")} & "\'")'  # StudentGroup=Col9
        f'& IF($I$5="All", "", " AND Col12=\'" & {_sq("$I$5")} & "\'")'  # GuideEmail=Col12
        f", 0), MATCH($K$5, {sort_list_range}, 0), "
        'IF(OR($K$5="Date Approved"), FALSE, TRUE)), "")'
    )
    return q


def _build_sorted_query_sheet5(sort_list_range):
    """SORT(QUERY()) for Sheet 5 (Roster Unenrollments). Source: _UnenrollData.
    Same col mapping as _ApprovedData: 1=Date 2=Mismatch ... 14=ExtStudentID
    """
    src = "'_UnenrollData'!A:N"
    q = (
        f"=IFERROR(SORT(QUERY({src}, "
        '"SELECT * WHERE 1=1"'
        f'& IF($A$5="All", "", " AND Col3=\'" & {_sq("$A$5")} & "\'")'  # Campus=Col3
        f'& IF($C$5="All", "", " AND Col4=\'" & {_sq("$C$5")} & "\'")'  # Grade=Col4
        f'& IF($E$5="All", "", " AND Col5=\'" & {_sq("$E$5")} & "\'")'  # Level=Col5
        f'& IF($G$5="All", "", " AND Col9=\'" & {_sq("$G$5")} & "\'")'  # StudentGroup=Col9
        f'& IF($I$5="All", "", " AND Col12=\'" & {_sq("$I$5")} & "\'")'  # GuideEmail=Col12
        f", 0), MATCH($K$5, {sort_list_range}, 0), "
        'IF(OR($K$5="Date Approved"), FALSE, TRUE)), "")'
    )
    return q


def _build_sorted_query_sheet6(sort_list_range):
    """SORT(QUERY()) for Sheet 6 (Rejected Changes). Source: _RejectedData.
    Same col mapping as _ApprovedData: 1=Date 2=Mismatch ... 14=ExtStudentID
    """
    src = "'_RejectedData'!A:N"
    q = (
        f"=IFERROR(SORT(QUERY({src}, "
        '"SELECT * WHERE 1=1"'
        f'& IF($A$5="All", "", " AND Col3=\'" & {_sq("$A$5")} & "\'")'  # Campus=Col3
        f'& IF($C$5="All", "", " AND Col4=\'" & {_sq("$C$5")} & "\'")'  # Grade=Col4
        f'& IF($E$5="All", "", " AND Col5=\'" & {_sq("$E$5")} & "\'")'  # Level=Col5
        f'& IF($G$5="All", "", " AND Col9=\'" & {_sq("$G$5")} & "\'")'  # StudentGroup=Col9
        f'& IF($I$5="All", "", " AND Col12=\'" & {_sq("$I$5")} & "\'")'  # GuideEmail=Col12
        f", 0), MATCH($K$5, {sort_list_range}, 0), "
        'IF(OR($K$5="Date Rejected"), FALSE, TRUE)), "")'
    )
    return q


def _clear_conditional_format_rules(sheets_service, spreadsheet_id, sheet_ids):
    """Remove all conditional formatting rules from specified sheets."""
    resp = _retry_api(
        lambda: sheets_service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties.sheetId,conditionalFormats)",
        )
        .execute()
    )
    requests = []
    for sheet in resp.get("sheets", []):
        sid = sheet["properties"]["sheetId"]
        if sid not in sheet_ids:
            continue
        for i in range(len(sheet.get("conditionalFormats", [])) - 1, -1, -1):
            requests.append(
                {"deleteConditionalFormatRule": {"sheetId": sid, "index": i}}
            )
    if requests:
        _retry_api(
            lambda: sheets_service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests})
            .execute()
        )


# ══════════════════════════════════════════════════════════════════════════
# DATA MIGRATION — fix corrupted cumulative tabs from v2.1.0 → v2.2.0
# ══════════════════════════════════════════════════════════════════════════


def _migrate_cumulative_tabs(sheets_service, spreadsheet_id):
    """Content-based migration: realign corrupted rows using email position.

    Target format (14 cols): Date, MismatchSummary, Campus, Grade, Level,
    FirstName, LastName, Email, StudentGroup, GuideFirst, GuideLast,
    GuideEmail, StudentID, ExtStudentID

    Strategy: find the student email (@2hourlearning.com) which must be at
    index 7 in the correct layout. Compute the shift and realign.
    """
    cumulative_tabs = [
        "_ApprovedData",
        "_AdditionsData",
        "_UnenrollData",
        "_RejectedData",
    ]

    for tab_name in cumulative_tabs:
        resp = _retry_api(
            lambda t=tab_name: sheets_service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=f"'{t}'!A:Z")
            .execute()
        )
        rows = resp.get("values", [])
        if not rows:
            continue

        # Quick skip: if every row has email at index 7, no migration needed.
        # (Sheets API strips trailing empty strings, so len<14 is normal post-migration.)
        all_aligned = all(
            len(r) > 7 and isinstance(r[7], str) and "2hourlearning" in r[7].lower()
            for r in rows
        )
        if all_aligned:
            continue

        migrated = []
        fixed = 0
        for row in rows:
            result = _realign_row(row)
            if result != row:
                fixed += 1
            migrated.append(result)

        if fixed == 0:
            continue

        print(f"  Migrating {tab_name}: {fixed}/{len(rows)} rows realigned...")
        _retry_api(
            lambda t=tab_name: sheets_service.spreadsheets()
            .values()
            .clear(spreadsheetId=spreadsheet_id, range=f"'{t}'!A:Z")
            .execute()
        )
        if migrated:
            _retry_api(
                lambda t=tab_name, m=migrated: sheets_service.spreadsheets()
                .values()
                .update(
                    spreadsheetId=spreadsheet_id,
                    range=f"'{t}'!A1",
                    valueInputOption="RAW",
                    body={"values": m},
                )
                .execute()
            )


def _realign_row(row):
    """Realign a single cumulative-tab row to 14-col format.

    Uses student email (@2hourlearning) as anchor — it must be at index 7.
    Falls back to guide email (@) patterns if student email not found.
    """
    TARGET_EMAIL_IDX = 7
    TARGET_COLS = 14

    # Find student email index (contains "2hourlearning")
    email_idx = None
    for i, val in enumerate(row):
        if isinstance(val, str) and "2hourlearning" in val.lower():
            email_idx = i
            break

    if email_idx is None:
        # No student email found — pad/truncate to 14 with blank MismatchSummary
        if len(row) < 2:
            return [""] * TARGET_COLS
        result = [row[0], ""] + list(row[1:])
        while len(result) < TARGET_COLS:
            result.append("")
        return result[:TARGET_COLS]

    shift = email_idx - TARGET_EMAIL_IDX

    if shift > 0:
        # Too many values before email — remove extras after date
        fixed = [row[0]] + list(row[1 + shift :])
    elif shift < 0:
        # Too few values before email — insert blanks after date
        fixed = [row[0]] + [""] * (-shift) + list(row[1:])
    else:
        fixed = list(row)

    # Clean up: replace FALSE/TRUE in MismatchSummary slot with ""
    if len(fixed) > 1 and str(fixed[1]).upper() in ("FALSE", "TRUE"):
        fixed[1] = ""

    # Pad or truncate to 14
    while len(fixed) < TARGET_COLS:
        fixed.append("")
    return fixed[:TARGET_COLS]


def _backfill_mismatch_summary(sheets_service, spreadsheet_id):
    """Fill blank MismatchSummary in cumulative tabs using _CorrData lookup.

    _CorrData has current mismatch types (col 12 = MismatchSummary, col 10 = StudentID).
    For _AdditionsData, type is always "Roster Addition".
    For _UnenrollData, type is always "Unenrolling".
    For _ApprovedData and _RejectedData, look up by StudentID.
    """
    # Backfill each cumulative tab
    backfill_map = {
        "_ApprovedData": None,  # lookup from _CorrData
        "_AdditionsData": "Roster Addition",  # always this type
        "_UnenrollData": "Unenrolling",  # always this type
        "_RejectedData": None,  # lookup from _CorrData
    }

    # First pass: check if any tab needs backfill (skip _CorrData read if all are filled)
    tab_rows = {}
    any_needs_backfill = False
    for tab_name in backfill_map:
        resp = _retry_api(
            lambda t=tab_name: sheets_service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=f"'{t}'!A:N")
            .execute()
        )
        rows = resp.get("values", [])
        tab_rows[tab_name] = rows
        if any(len(r) > 1 and not str(r[1]).strip() for r in rows):
            any_needs_backfill = True

    if not any_needs_backfill:
        return

    # Build lookup: student_id -> mismatch_summary from _CorrData
    resp = _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range="'_CorrData'!A:M")
        .execute()
    )
    corr_rows = resp.get("values", [])
    lookup = {}
    for r in corr_rows:
        if len(r) > 12:
            sid = str(r[10]).strip()
            mismatch = str(r[12]).strip()
            if sid and mismatch:
                lookup[sid] = mismatch

    for tab_name, fixed_type in backfill_map.items():
        rows = tab_rows[tab_name]
        if not rows:
            continue

        filled = 0
        for row in rows:
            while len(row) < 14:
                row.append("")
            if row[1].strip():
                continue  # already has mismatch summary

            if fixed_type:
                row[1] = fixed_type
                filled += 1
            else:
                sid = str(row[12]).strip()
                if sid in lookup:
                    row[1] = lookup[sid]
                    filled += 1

        if filled == 0:
            continue

        print(f"  Backfilling {tab_name}: {filled} rows filled with mismatch type")
        _retry_api(
            lambda t=tab_name: sheets_service.spreadsheets()
            .values()
            .clear(spreadsheetId=spreadsheet_id, range=f"'{t}'!A:Z")
            .execute()
        )
        _retry_api(
            lambda t=tab_name, r=rows: sheets_service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=f"'{t}'!A1",
                valueInputOption="RAW",
                body={"values": r},
            )
            .execute()
        )


# ══════════════════════════════════════════════════════════════════════════
# MAIN WRITE FUNCTION
# ══════════════════════════════════════════════════════════════════════════


def write_corrections(sheets_service, corrections_map, corrections_sis):
    sid = OUTPUT_SPREADSHEET_ID
    print(f"\n  Writing {len(corrections_map)} correction rows...")

    # ── Ensure all tabs exist (1 read + 1 batch create) ────────────────
    all_tab_names = [
        TAB_CORRECTED,
        TAB_SIS,
        TAB_APPROVED,
        "_CorrData",
        "_SISData",
        "_ApprovedData",
        "_AdditionsData",
        "_UnenrollData",
        "_RejectedData",
        "_Lists",
        TAB_ADDITIONS,
        TAB_UNENROLL,
        TAB_REJECTED,
    ]
    tab_ids = _ensure_all_tabs(sheets_service, sid, all_tab_names)
    sheet1_id = tab_ids[TAB_CORRECTED]
    sheet2_id = tab_ids[TAB_SIS]
    sheet3_id = tab_ids[TAB_APPROVED]
    sheet4_id = tab_ids[TAB_ADDITIONS]
    sheet5_id = tab_ids[TAB_UNENROLL]
    sheet6_id = tab_ids[TAB_REJECTED]
    corr_data_id = tab_ids["_CorrData"]
    sis_data_id = tab_ids["_SISData"]
    approved_data_id = tab_ids["_ApprovedData"]
    additions_data_id = tab_ids["_AdditionsData"]
    unenroll_data_id = tab_ids["_UnenrollData"]
    rejected_data_id = tab_ids["_RejectedData"]
    lists_id = tab_ids["_Lists"]

    visible_ids = [sheet1_id, sheet2_id, sheet3_id, sheet4_id, sheet5_id, sheet6_id]

    # ── Clear banding + conditional formatting + data ─────────────────
    _clear_banding(sheets_service, sid, visible_ids)
    _clear_conditional_format_rules(sheets_service, sid, [sheet1_id])

    # Unmerge all cells on visible sheets (1 batched call instead of 6)
    unmerge_requests = [
        {
            "unmergeCells": {
                "range": {
                    "sheetId": uid,
                    "startRowIndex": 0,
                    "endRowIndex": 1000,
                    "startColumnIndex": 0,
                    "endColumnIndex": 50,
                }
            }
        }
        for uid in visible_ids
    ]
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .batchUpdate(spreadsheetId=sid, body={"requests": unmerge_requests})
        .execute()
    )

    # Clear visible + data tabs (1 batched call instead of 9)
    clear_tabs = [
        TAB_CORRECTED,
        TAB_SIS,
        TAB_APPROVED,
        TAB_ADDITIONS,
        TAB_UNENROLL,
        TAB_REJECTED,
        "_CorrData",
        "_SISData",
        "_Lists",
    ]
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .batchClear(
            spreadsheetId=sid,
            body={"ranges": [f"'{t}'!A:Z" for t in clear_tabs]},
        )
        .execute()
    )
    # NOTE: _ApprovedData, _AdditionsData, _UnenrollData, _RejectedData NOT cleared — cumulative

    # ── Migrate cumulative tabs to 14-col format (Date + MismatchSummary + 12 fields) ──
    _migrate_cumulative_tabs(sheets_service, sid)
    _backfill_mismatch_summary(sheets_service, sid)

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

    # _Lists: 11 columns — 5 filter values (A-E) + 6 sort options (F-K)
    n = {k: len(v) for k, v in unique.items()}
    s1 = SORT_OPTS_SHEET1
    s2 = SORT_OPTS_SHEET2
    s3 = SORT_OPTS_SHEET3
    s4 = SORT_OPTS_SHEET4
    s5 = SORT_OPTS_SHEET5
    s6 = SORT_OPTS_SHEET6
    max_len = max(max(n.values()), len(s1), len(s2), len(s3), len(s4), len(s5), len(s6))
    lists_header = [
        "campus",
        "grade",
        "level",
        "student_group",
        "guide_email",
        "sort_sheet1",
        "sort_sheet2",
        "sort_sheet3",
        "sort_sheet4",
        "sort_sheet5",
        "sort_sheet6",
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
        row.append(s4[i] if i < len(s4) else "")
        row.append(s5[i] if i < len(s5) else "")
        row.append(s6[i] if i < len(s6) else "")
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
        "sort4": f"_Lists!I2:I{len(s4) + 1}",
        "sort5": f"_Lists!J2:J{len(s5) + 1}",
        "sort6": f"_Lists!K2:K{len(s6) + 1}",
    }

    # ══════════════════════════════════════════════════════════════════
    # WRITE VISIBLE SHEETS
    # ══════════════════════════════════════════════════════════════════
    print("  Writing visible sheets...")
    vals_raw = {}
    vals_entered = {}

    NC1 = 15  # accept + reject + 12 fields + mismatch summary
    NC2 = 12  # 12 fields
    NC3 = 14  # date + mismatch summary + 12 fields
    NC4 = 14  # date + mismatch summary + 12 fields (Roster Additions)
    NC5 = 14  # date + mismatch summary + 12 fields (Roster Unenrollments)
    NC6 = 15  # date + mismatch summary + 12 fields + reason for rejection

    # ── Sheet 1: Corrected Roster Info ────────────────────────────────
    vals_raw[f"'{TAB_CORRECTED}'!A1"] = [["Corrected Roster Info"]]
    # Row 4: filter labels (C,E,G,I,K) + Sort By (M) — offset by 2 for accept/reject cols
    s1_r4 = [""] * NC1
    s1_r4[2] = "Campus"
    s1_r4[4] = "Grade"
    s1_r4[6] = "Level"
    s1_r4[8] = "Student Group"
    s1_r4[10] = "Guide Email"
    s1_r4[12] = "SORT BY"
    vals_raw[f"'{TAB_CORRECTED}'!A4"] = [s1_r4]
    # Row 5: dropdown defaults
    s1_r5 = [""] * NC1
    s1_r5[2] = "All"
    s1_r5[4] = "All"
    s1_r5[6] = "All"
    s1_r5[8] = "All"
    s1_r5[10] = "All"
    s1_r5[12] = "Campus"
    vals_raw[f"'{TAB_CORRECTED}'!A5"] = [s1_r5]
    # Row 6: headers
    vals_raw[f"'{TAB_CORRECTED}'!A6"] = [
        ["Accept Changes", "Reject Changes"] + OUTPUT_FIELDS + ["Mismatch Summary"]
    ]
    # Row 7: SORT(QUERY()) formula in C7 (offset by 2)
    vals_entered[f"'{TAB_CORRECTED}'!C7"] = [
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
    vals_raw[f"'{TAB_APPROVED}'!A6"] = [
        ["Date Approved", "Mismatch Summary"] + OUTPUT_FIELDS
    ]
    vals_entered[f"'{TAB_APPROVED}'!A7"] = [
        [_build_sorted_query_sheet3(list_ranges["sort3"])]
    ]

    # ── Sheet 4: Roster Additions ─────────────────────────────────────
    vals_raw[f"'{TAB_ADDITIONS}'!A1"] = [["Roster Additions"]]
    s4_r4 = [""] * NC4
    s4_r4[0] = "Campus"
    s4_r4[2] = "Grade"
    s4_r4[4] = "Level"
    s4_r4[6] = "Student Group"
    s4_r4[8] = "Guide Email"
    s4_r4[10] = "SORT BY"
    vals_raw[f"'{TAB_ADDITIONS}'!A4"] = [s4_r4]
    s4_r5 = [""] * NC4
    s4_r5[0] = "All"
    s4_r5[2] = "All"
    s4_r5[4] = "All"
    s4_r5[6] = "All"
    s4_r5[8] = "All"
    s4_r5[10] = "Date Approved"
    vals_raw[f"'{TAB_ADDITIONS}'!A5"] = [s4_r5]
    vals_raw[f"'{TAB_ADDITIONS}'!A6"] = [
        ["Date Approved", "Mismatch Summary"] + OUTPUT_FIELDS
    ]
    vals_entered[f"'{TAB_ADDITIONS}'!A7"] = [
        [_build_sorted_query_sheet4(list_ranges["sort4"])]
    ]

    # ── Sheet 5: Roster Unenrollments ─────────────────────────────────
    vals_raw[f"'{TAB_UNENROLL}'!A1"] = [["Roster Unenrollments"]]
    s5_r4 = [""] * NC5
    s5_r4[0] = "Campus"
    s5_r4[2] = "Grade"
    s5_r4[4] = "Level"
    s5_r4[6] = "Student Group"
    s5_r4[8] = "Guide Email"
    s5_r4[10] = "SORT BY"
    vals_raw[f"'{TAB_UNENROLL}'!A4"] = [s5_r4]
    s5_r5 = [""] * NC5
    s5_r5[0] = "All"
    s5_r5[2] = "All"
    s5_r5[4] = "All"
    s5_r5[6] = "All"
    s5_r5[8] = "All"
    s5_r5[10] = "Date Approved"
    vals_raw[f"'{TAB_UNENROLL}'!A5"] = [s5_r5]
    vals_raw[f"'{TAB_UNENROLL}'!A6"] = [
        ["Date Approved", "Mismatch Summary"] + OUTPUT_FIELDS
    ]
    vals_entered[f"'{TAB_UNENROLL}'!A7"] = [
        [_build_sorted_query_sheet5(list_ranges["sort5"])]
    ]

    # ── Sheet 6: Rejected Changes ─────────────────────────────────────
    vals_raw[f"'{TAB_REJECTED}'!A1"] = [["Rejected Changes"]]
    s6_r4 = [""] * NC6
    s6_r4[0] = "Campus"
    s6_r4[2] = "Grade"
    s6_r4[4] = "Level"
    s6_r4[6] = "Student Group"
    s6_r4[8] = "Guide Email"
    s6_r4[10] = "SORT BY"
    vals_raw[f"'{TAB_REJECTED}'!A4"] = [s6_r4]
    s6_r5 = [""] * NC6
    s6_r5[0] = "All"
    s6_r5[2] = "All"
    s6_r5[4] = "All"
    s6_r5[6] = "All"
    s6_r5[8] = "All"
    s6_r5[10] = "Date Rejected"
    vals_raw[f"'{TAB_REJECTED}'!A5"] = [s6_r5]
    vals_raw[f"'{TAB_REJECTED}'!A6"] = [
        ["Date Rejected", "Mismatch Summary"] + OUTPUT_FIELDS + ["Reason for Rejection"]
    ]
    vals_entered[f"'{TAB_REJECTED}'!A7"] = [
        [_build_sorted_query_sheet6(list_ranges["sort6"])]
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
            sort_col_start=12,
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
    fmt.extend(
        _format_visible_sheet(
            sheet4_id,
            NC4,
            list_ranges,
            has_checkbox=False,
            sort_col_start=10,
            sort_list_key="sort4",
            num_data_rows=0,
        )
    )
    fmt.extend(
        _format_visible_sheet(
            sheet5_id,
            NC5,
            list_ranges,
            has_checkbox=False,
            sort_col_start=10,
            sort_list_key="sort5",
            num_data_rows=0,
        )
    )
    fmt.extend(
        _format_visible_sheet(
            sheet6_id,
            NC6,
            list_ranges,
            has_checkbox=False,
            sort_col_start=10,
            sort_list_key="sort6",
            num_data_rows=0,
        )
    )
    # Reason for Rejection column width on Sheet 6 (last col = index 14)
    fmt.append(_cw(sheet6_id, NC6 - 1, 250))

    # Mismatch Summary col B (index 1): red header + 200px width on Sheets 3-6
    for ms_sheet_id in [sheet3_id, sheet4_id, sheet5_id, sheet6_id]:
        fmt.append(
            _rc(
                ms_sheet_id,
                5,
                6,
                1,
                2,
                _cf(bg=RED_HDR, fg=WHITE, bold=True, sz=10, ha="CENTER", va="MIDDLE"),
            )
        )
        fmt.append(_cw(ms_sheet_id, 1, 200))

    # Hide data tabs
    for hid in [
        corr_data_id,
        sis_data_id,
        approved_data_id,
        additions_data_id,
        unenroll_data_id,
        rejected_data_id,
        lists_id,
    ]:
        fmt.append(
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": hid, "hidden": True},
                    "fields": "hidden",
                }
            }
        )

    # Date format on Sheets 3/4/5/6 col A (QUERY strips number formatting)
    for date_sheet_id in [sheet3_id, sheet4_id, sheet5_id, sheet6_id]:
        fmt.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": date_sheet_id,
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

    # Checkbox data validation on Sheet 1 col A (Accept) and col B (Reject), rows 7+
    cb_end = 6 + len(corr_rows) + 10
    for cb_col in [0, 1]:  # col A = Accept, col B = Reject
        fmt.append(
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet1_id,
                        "startRowIndex": 6,
                        "endRowIndex": cb_end,
                        "startColumnIndex": cb_col,
                        "endColumnIndex": cb_col + 1,
                    },
                    "rule": {"condition": {"type": "BOOLEAN"}, "showCustomUi": True},
                }
            }
        )

    # Accept column (A) light green background, Reject column (B) light red background
    fmt.append(_rc(sheet1_id, 6, cb_end, 0, 1, _cf(bg=ACCEPT_BG, sz=10, ha="CENTER")))
    fmt.append(_rc(sheet1_id, 6, cb_end, 1, 2, _cf(bg=REJECT_BG, sz=10, ha="CENTER")))

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
    offset = 2 if has_checkbox else 0  # 2 cols for accept/reject checkboxes

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
        fmt.append(_cw(sheet_id, 0, 105))  # Accept Changes col
        fmt.append(_cw(sheet_id, 1, 105))  # Reject Changes col
    field_widths = [150, 60, 80, 100, 100, 220, 150, 100, 100, 220, 120, 140]
    for i, w in enumerate(field_widths):
        fmt.append(_cw(sheet_id, offset + i, w))
    if has_checkbox:
        fmt.append(_cw(sheet_id, offset + len(field_widths), 200))  # Mismatch Summary

    # ── Alternating row colors ────────────────────────────────────────
    end_row = max(6 + num_data_rows + 5, 206)
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

    # ── Mismatch Summary column: dark red header + conditional color by type (Sheet 1 only) ──
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
        # Conditional formatting rules (priority order: 0=highest)
        mm_range = {
            "sheetId": sheet_id,
            "startRowIndex": 6,
            "endRowIndex": end_row,
            "startColumnIndex": mismatch_col,
            "endColumnIndex": mismatch_col + 1,
        }
        # Rule 0: "Roster Addition" → green
        fmt.append(
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [mm_range],
                        "booleanRule": {
                            "condition": {
                                "type": "TEXT_CONTAINS",
                                "values": [{"userEnteredValue": "Roster Addition"}],
                            },
                            "format": {"backgroundColor": GREEN_LIGHT},
                        },
                    },
                    "index": 0,
                }
            }
        )
        # Rule 1: "Unenrolling" → light yellow
        fmt.append(
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [mm_range],
                        "booleanRule": {
                            "condition": {
                                "type": "TEXT_CONTAINS",
                                "values": [{"userEnteredValue": "Unenrolling"}],
                            },
                            "format": {"backgroundColor": YELLOW_LIGHT},
                        },
                    },
                    "index": 1,
                }
            }
        )
        # Rule 2: NOT_BLANK → yellow (catches field mismatches)
        fmt.append(
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [mm_range],
                        "booleanRule": {
                            "condition": {"type": "NOT_BLANK"},
                            "format": {"backgroundColor": YELLOW_MM},
                        },
                    },
                    "index": 2,
                }
            }
        )

    return fmt
