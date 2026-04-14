"""Google Sheets API writer for the Weekly Corrections tool.

Layout (all 3 visible sheets):
  Row 1: Title (merged, navy dark, 20pt bold white)
  Row 2: Caption with User Guide hyperlink (merged, navy med, italic grey)
  Row 3: Spacer (5px, dark)
  Row 4: Filter labels (merged pairs, dark navy, small grey text)
  Row 5: Dropdown values (merged pairs, teal bg, data validation from _Lists)
  Row 6: Column headers (navy, bold white)
  Row 7+: QUERY formula output (filtered from hidden _CorrData / _SISData tabs)

Filtering is done by QUERY formulas — NOT by Apps Script.
The QUERY references dropdown cells in row 5 and auto-recalculates.
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

GUIDE_URL = (
    "https://docs.google.com/document/d/1O1WEAHSttdNVRUa_CoQ3T6w4QEFPyLz5FDdM2IMHEu4"
)

CAPTION_TEXT = (
    "Implementation Managers are responsible for checking off the corrections "
    "to make for their own schools  |  Ticket will be automatically submitted "
    "once a week with the checked off boxes  |  User Guide"
)

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


# ── Sheets API request builders (matching sheets_builder.py) ──────────────


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


def _no_bdr():
    b = {"style": "NONE"}
    return {"top": b, "bottom": b, "left": b, "right": b}


def _bdr():
    b = {"style": "SOLID", "colorStyle": {"rgbColor": GREY_BORDER}}
    return {"top": b, "bottom": b, "left": b, "right": b}


def _subtitle_with_link(sid, row, text, nc):
    """Create subtitle cell with clickable 'User Guide' hyperlink using textFormatRuns."""
    link_start = text.find("User Guide")
    link_end = link_start + len("User Guide")
    # Build textFormatRuns — trailing run only if link doesn't end at string boundary
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
    """Extract unique values for each filter field from both data sets."""
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


# ── QUERY formula builders ────────────────────────────────────────────────


def _build_query_formula_sheet1(tab_name):
    """Build QUERY formula for Sheet 1 (Corrected Roster Info).

    Data is on hidden _CorrData tab (13 cols A-M: Campus..MismatchSummary).
    Dropdowns are in row 5 at merged pairs starting col B: B5, D5, F5, H5, J5.
    Formula outputs to B7 (col A is for checkboxes).
    """
    src = f"'{tab_name}'!A:M"
    # Col1=Campus, Col2=Grade, Col3=Level, Col7=StudentGroup, Col10=GuideEmail
    q = (
        f"=IFERROR(QUERY({src}, "
        '"SELECT * WHERE 1=1"'
        '& IF($B$5="All", "", " AND Col1=\'" & $B$5 & "\'")'  # Campus
        '& IF($D$5="All", "", " AND Col2=\'" & $D$5 & "\'")'  # Grade
        '& IF($F$5="All", "", " AND Col3=\'" & $F$5 & "\'")'  # Level
        '& IF($H$5="All", "", " AND Col7=\'" & $H$5 & "\'")'  # Student Group
        '& IF($J$5="All", "", " AND Col10=\'" & $J$5 & "\'")'  # Guide Email
        ', 0), "")'
    )
    return q


def _build_query_formula_sheet2(tab_name):
    """Build QUERY formula for Sheet 2 (Current Roster Info in SIS).

    Data is on hidden _SISData tab (12 cols A-L: Campus..ExtStudentID).
    Dropdowns are in row 5 at merged pairs: A5, C5, E5, G5, I5.
    Formula outputs to A7.
    """
    src = f"'{tab_name}'!A:L"
    # Col1=Campus, Col2=Grade, Col3=Level, Col7=StudentGroup, Col10=GuideEmail
    q = (
        f"=IFERROR(QUERY({src}, "
        '"SELECT * WHERE 1=1"'
        '& IF($A$5="All", "", " AND Col1=\'" & $A$5 & "\'")'  # Campus
        '& IF($C$5="All", "", " AND Col2=\'" & $C$5 & "\'")'  # Grade
        '& IF($E$5="All", "", " AND Col3=\'" & $E$5 & "\'")'  # Level
        '& IF($G$5="All", "", " AND Col7=\'" & $G$5 & "\'")'  # Student Group
        '& IF($I$5="All", "", " AND Col10=\'" & $I$5 & "\'")'  # Guide Email
        ', 0), "")'
    )
    return q


# ══════════════════════════════════════════════════════════════════════════
# MAIN WRITE FUNCTION
# ══════════════════════════════════════════════════════════════════════════


def write_corrections(sheets_service, corrections_map, corrections_sis):
    """Write comparison results with QUERY-based dropdown filtering."""
    sid = OUTPUT_SPREADSHEET_ID
    print(f"\n  Writing {len(corrections_map)} correction rows...")

    # ── Ensure all tabs exist ─────────────────────────────────────────
    sheet1_id = _ensure_tab_exists(sheets_service, sid, TAB_CORRECTED)
    sheet2_id = _ensure_tab_exists(sheets_service, sid, TAB_SIS)
    sheet3_id = _ensure_tab_exists(sheets_service, sid, TAB_APPROVED)
    corr_data_id = _ensure_tab_exists(sheets_service, sid, "_CorrData")
    sis_data_id = _ensure_tab_exists(sheets_service, sid, "_SISData")
    lists_id = _ensure_tab_exists(sheets_service, sid, "_Lists")

    # ── Clear everything (banding + data) ─────────────────────────────
    _clear_banding(sheets_service, sid, [sheet1_id, sheet2_id, sheet3_id])
    for tab in [TAB_CORRECTED, TAB_SIS, "_CorrData", "_SISData", "_Lists"]:
        _retry_api(
            lambda t=tab: sheets_service.spreadsheets()
            .values()
            .clear(spreadsheetId=sid, range=f"'{t}'!A:Z")
            .execute()
        )

    if not corrections_map:
        print("  No mismatches found — sheets cleared.")
        _retry_api(
            lambda: sheets_service.spreadsheets()
            .values()
            .update(
                spreadsheetId=sid,
                range=f"'{TAB_CORRECTED}'!A1",
                valueInputOption="RAW",
                body={
                    "values": [["No mismatches found between MAP roster and SIS data."]]
                },
            )
            .execute()
        )
        return

    # ── Compute unique values for dropdowns ───────────────────────────
    unique = _compute_unique_values(corrections_map, corrections_sis)

    # ══════════════════════════════════════════════════════════════════
    # WRITE HIDDEN TABS
    # ══════════════════════════════════════════════════════════════════
    print("  Writing hidden data tabs...")

    # _CorrData: 13 columns (Campus through MismatchSummary), NO header
    corr_rows = []
    for rec in corrections_map:
        row = [rec.get(field, "") for field in OUTPUT_FIELDS]
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

    # _SISData: 12 columns (Campus through ExtStudentID), NO header
    sis_rows = []
    for rec in corrections_sis:
        row = [rec.get(field, "") for field in OUTPUT_FIELDS]
        sis_rows.append(row)

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

    # _Lists: unique filter values (5 columns)
    max_len = max(len(v) for v in unique.values())
    lists_header = ["campus", "grade", "level", "student_group", "guide_email"]
    lists_rows = [lists_header]
    for i in range(max_len):
        row = []
        for key in lists_header:
            vals = unique[key]
            row.append(vals[i] if i < len(vals) else "")
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

    n_campus = len(unique["campus"])
    n_grade = len(unique["grade"])
    n_level = len(unique["level"])
    n_group = len(unique["student_group"])
    n_email = len(unique["guide_email"])
    list_ranges = {
        "campus": f"_Lists!A2:A{n_campus + 1}",
        "grade": f"_Lists!B2:B{n_grade + 1}",
        "level": f"_Lists!C2:C{n_level + 1}",
        "student_group": f"_Lists!D2:D{n_group + 1}",
        "guide_email": f"_Lists!E2:E{n_email + 1}",
    }

    # ══════════════════════════════════════════════════════════════════
    # BUILD VISIBLE SHEETS (RAW values for labels/headers/defaults)
    # ══════════════════════════════════════════════════════════════════
    print("  Writing visible sheets...")

    # Number of columns
    NC1 = 14  # Sheet 1: checkbox + 12 fields + mismatch summary
    NC2 = 12  # Sheet 2: 12 fields
    NC3 = 13  # Sheet 3: date + 12 fields

    # ── Sheet 1: Corrected Roster Info ────────────────────────────────
    # Row 1: Title
    vals_raw = {}
    vals_entered = {}  # USER_ENTERED for formulas

    vals_raw[f"'{TAB_CORRECTED}'!A1"] = [["Corrected Roster Info"]]
    # Row 2: Caption (written via updateCells for link formatting)
    # Row 4: Filter labels (cols B,D,F,H,J — merged pairs)
    s1_r4 = [""] * NC1
    s1_r4[1] = "Campus"
    s1_r4[3] = "Grade"
    s1_r4[5] = "Level"
    s1_r4[7] = "Student Group"
    s1_r4[9] = "Guide Email"
    vals_raw[f"'{TAB_CORRECTED}'!A4"] = [s1_r4]

    # Row 5: Dropdown defaults
    s1_r5 = [""] * NC1
    s1_r5[1] = "All"
    s1_r5[3] = "All"
    s1_r5[5] = "All"
    s1_r5[7] = "All"
    s1_r5[9] = "All"
    vals_raw[f"'{TAB_CORRECTED}'!A5"] = [s1_r5]

    # Row 6: Headers
    header1 = ["\u2713"] + OUTPUT_FIELDS + ["Mismatch Summary"]
    vals_raw[f"'{TAB_CORRECTED}'!A6"] = [header1]

    # Row 7: QUERY formula in B7 (col A left empty for checkboxes)
    query1 = _build_query_formula_sheet1("_CorrData")
    vals_entered[f"'{TAB_CORRECTED}'!B7"] = [[query1]]

    # ── Sheet 2: Current Roster Info in SIS ───────────────────────────
    vals_raw[f"'{TAB_SIS}'!A1"] = [["Current Roster Info in SIS"]]
    # Row 4: Filter labels (cols A,C,E,G,I — merged pairs)
    s2_r4 = [""] * NC2
    s2_r4[0] = "Campus"
    s2_r4[2] = "Grade"
    s2_r4[4] = "Level"
    s2_r4[6] = "Student Group"
    s2_r4[8] = "Guide Email"
    vals_raw[f"'{TAB_SIS}'!A4"] = [s2_r4]

    # Row 5: Dropdown defaults
    s2_r5 = [""] * NC2
    s2_r5[0] = "All"
    s2_r5[2] = "All"
    s2_r5[4] = "All"
    s2_r5[6] = "All"
    s2_r5[8] = "All"
    vals_raw[f"'{TAB_SIS}'!A5"] = [s2_r5]

    # Row 6: Headers
    header2 = OUTPUT_FIELDS[:]
    vals_raw[f"'{TAB_SIS}'!A6"] = [header2]

    # Row 7: QUERY formula in A7
    query2 = _build_query_formula_sheet2("_SISData")
    vals_entered[f"'{TAB_SIS}'!A7"] = [[query2]]

    # ── Sheet 3: Automated Correction List ────────────────────────────
    vals_raw[f"'{TAB_APPROVED}'!A1"] = [["Automated Correction List"]]
    header3 = ["Date Approved"] + OUTPUT_FIELDS
    vals_raw[f"'{TAB_APPROVED}'!A6"] = [header3]

    # ── Write RAW values ──────────────────────────────────────────────
    for rng, val in vals_raw.items():
        _retry_api(
            lambda r=rng, v=val: sheets_service.spreadsheets()
            .values()
            .update(
                spreadsheetId=sid, range=r, valueInputOption="RAW", body={"values": v}
            )
            .execute()
        )

    # ── Write USER_ENTERED values (formulas) ──────────────────────────
    for rng, val in vals_entered.items():
        _retry_api(
            lambda r=rng, v=val: sheets_service.spreadsheets()
            .values()
            .update(
                spreadsheetId=sid,
                range=r,
                valueInputOption="USER_ENTERED",
                body={"values": v},
            )
            .execute()
        )

    # ══════════════════════════════════════════════════════════════════
    # FORMATTING (all via batchUpdate)
    # ══════════════════════════════════════════════════════════════════
    print("  Applying formatting...")
    fmt = []

    # ── Format Sheet 1 ────────────────────────────────────────────────
    fmt.extend(
        _format_visible_sheet(
            sheet1_id,
            NC1,
            list_ranges,
            has_checkbox=True,
            title="Corrected Roster Info",
            num_data_rows=len(corr_rows),
        )
    )

    # ── Format Sheet 2 ────────────────────────────────────────────────
    fmt.extend(
        _format_visible_sheet(
            sheet2_id,
            NC2,
            list_ranges,
            has_checkbox=False,
            title="Current Roster Info in SIS",
            num_data_rows=len(sis_rows),
        )
    )

    # ── Format Sheet 3 (no dropdowns, no QUERY) ───────────────────────
    fmt.extend(_format_sheet3(sheet3_id, NC3))

    # ── Hide data tabs ────────────────────────────────────────────────
    for hid in [corr_data_id, sis_data_id, lists_id]:
        fmt.append(
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": hid, "hidden": True},
                    "fields": "hidden",
                }
            }
        )

    # ── Checkbox data validation on Sheet 1 col A, rows 7+ ───────────
    # Set enough rows for all possible data (QUERY expands dynamically)
    fmt.append(
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet1_id,
                    "startRowIndex": 6,
                    "endRowIndex": 6 + len(corr_rows) + 10,  # extra buffer
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "rule": {"condition": {"type": "BOOLEAN"}, "showCustomUi": True},
            }
        }
    )

    # ── Execute all formatting ────────────────────────────────────────
    # Send in batches of 200 to avoid payload limits
    for i in range(0, len(fmt), 200):
        batch = fmt[i : i + 200]
        _retry_api(
            lambda b=batch: sheets_service.spreadsheets()
            .batchUpdate(spreadsheetId=sid, body={"requests": b})
            .execute()
        )

    print(f"  Done — {len(corrections_map)} students written to corrections sheet.")


# ══════════════════════════════════════════════════════════════════════════
# SHEET FORMATTING HELPERS
# ══════════════════════════════════════════════════════════════════════════


def _format_visible_sheet(
    sheet_id, nc, list_ranges, has_checkbox, title, num_data_rows
):
    """Build all formatting requests for a visible sheet (Sheet 1 or 2)."""
    fmt = []
    offset = 1 if has_checkbox else 0  # col A offset for checkbox sheets

    # ── Row 0 (title): merged, navy dark, 20pt bold white centered ────
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

    # ── Row 1 (caption): merged, navy med, italic grey + User Guide link ──
    fmt.append(_mg(sheet_id, 1, 2, 0, nc))
    fmt.append(_rh(sheet_id, 1, 28))
    fmt.append(
        _rc(
            sheet_id,
            1,
            2,
            0,
            nc,
            _cf(bg=NAVY_MED, fg=GREY_LABEL, italic=True, sz=10, ha="CENTER"),
        )
    )
    fmt.append(_subtitle_with_link(sheet_id, 1, CAPTION_TEXT, nc))

    # ── Row 2 (spacer): thin, dark ────────────────────────────────────
    fmt.append(_rh(sheet_id, 2, 5))
    fmt.append(_rc(sheet_id, 2, 3, 0, nc, _cf(bg=NAVY_DARK)))

    # ── Rows 3-4 (filter area): dark background ──────────────────────
    fmt.append(_rh(sheet_id, 3, 18))
    fmt.append(_rh(sheet_id, 4, 34))
    fmt.append(_rc(sheet_id, 3, 5, 0, nc, _cf(bg=NAVY_DARK)))

    # ── Filter dropdowns (5 pairs of merged cols) ─────────────────────
    filter_keys = ["campus", "grade", "level", "student_group", "guide_email"]
    filter_labels = ["Campus", "Grade", "Level", "Student Group", "Guide Email"]

    for i, (label, key) in enumerate(zip(filter_labels, filter_keys)):
        c1 = offset + i * 2
        c2 = c1 + 2
        if c2 > nc:
            break

        # Merge label (row 3) and dropdown (row 4)
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
        fmt.append(_mg(sheet_id, 4, 5, c1, c2))
        fmt.append(
            _rc(
                sheet_id,
                4,
                5,
                c1,
                c2,
                _cf(bg=FILTER_BG, fg=WHITE, bold=True, sz=11, ha="CENTER", va="MIDDLE"),
            )
        )

        # Data validation dropdown
        if key in list_ranges:
            fmt.append(_dd(sheet_id, 4, c1, list_ranges[key]))

    # ── Row 5 (headers): navy, bold white ─────────────────────────────
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
    data_start = offset
    if has_checkbox:
        fmt.append(_cw(sheet_id, 0, 30))  # checkbox column

    field_widths = [150, 60, 80, 100, 100, 220, 150, 100, 100, 220, 120, 140]
    for i, w in enumerate(field_widths):
        fmt.append(_cw(sheet_id, data_start + i, w))

    if has_checkbox:
        fmt.append(
            _cw(sheet_id, data_start + len(field_widths), 200)
        )  # mismatch summary

    # ── Alternating row colors (rows 6+) ──────────────────────────────
    end_row = 6 + num_data_rows + 5  # buffer
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

    return fmt


def _format_sheet3(sheet_id, nc):
    """Format Sheet 3 (Automated Correction List) — title/caption + headers only."""
    fmt = []

    # Row 0: Title
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

    # Row 1: Caption
    fmt.append(_mg(sheet_id, 1, 2, 0, nc))
    fmt.append(_rh(sheet_id, 1, 28))
    fmt.append(
        _rc(
            sheet_id,
            1,
            2,
            0,
            nc,
            _cf(bg=NAVY_MED, fg=GREY_LABEL, italic=True, sz=10, ha="CENTER"),
        )
    )
    fmt.append(_subtitle_with_link(sheet_id, 1, CAPTION_TEXT, nc))

    # Rows 2-4: Spacer (dark)
    for r in range(2, 5):
        fmt.append(_rh(sheet_id, r, 5))
        fmt.append(_rc(sheet_id, r, r + 1, 0, nc, _cf(bg=NAVY_DARK)))

    # Row 5: Headers
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

    # Freeze rows 1-6
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

    # Column widths
    fmt.append(_cw(sheet_id, 0, 160))  # Date Approved
    field_widths = [150, 60, 80, 100, 100, 220, 150, 100, 100, 220, 120, 140]
    for i, w in enumerate(field_widths):
        fmt.append(_cw(sheet_id, 1 + i, w))

    return fmt
