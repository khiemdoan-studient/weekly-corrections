"""Google Sheets API writer for the Weekly Corrections tool.

Layout (all 3 sheets):
  Row 1: Filter labels (dark navy background)
  Row 2: Dropdown values (teal background, data validation from _Lists)
  Row 3: Column headers (navy, bold, white)
  Row 4+: Data rows (alternating white/grey)
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
FILTER_BG = _rgb("1E3A5F")
WHITE = _rgb("FFFFFF")
GREY_LABEL = _rgb("94A3B8")
ALT_ROW = _rgb("EDF2F7")

# Filter definitions: (label, column_index_in_data)
# These map to columns in the output data (0-based relative to data start)
FILTER_DEFS_SHEET1 = [
    ("Campus", 1),  # col B
    ("Grade", 2),  # col C
    ("Level", 3),  # col D
    ("Student Group", 7),  # col H
    ("Guide Email", 10),  # col K
]
FILTER_DEFS_SHEET2 = [
    ("Campus", 0),  # col A
    ("Grade", 1),  # col B
    ("Level", 2),  # col C
    ("Student Group", 6),  # col G
    ("Guide Email", 9),  # col J
]
FILTER_DEFS_SHEET3 = [
    ("Campus", 1),  # col B (after Date Approved)
    ("Grade", 2),  # col C
    ("Level", 3),  # col D
    ("Student Group", 7),  # col H
    ("Guide Email", 10),  # col K
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


# ── Compute unique filter values ───────────────────────────────────────────


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

    # Sort and prepend "All"
    return {k: ["All"] + sorted(v) for k, v in unique.items()}


# ── Main write function ───────────────────────────────────────────────────


def write_corrections(sheets_service, corrections_map, corrections_sis):
    """Write comparison results with dashboard-style dropdown filters."""
    sid = OUTPUT_SPREADSHEET_ID
    print(f"\n  Writing {len(corrections_map)} correction rows...")

    # Ensure all tabs exist
    sheet1_id = _ensure_tab_exists(sheets_service, sid, TAB_CORRECTED)
    sheet2_id = _ensure_tab_exists(sheets_service, sid, TAB_SIS)
    sheet3_id = _ensure_tab_exists(sheets_service, sid, TAB_APPROVED)
    lists_id = _ensure_tab_exists(sheets_service, sid, "_Lists")

    # Clear banding + data on sheets 1, 2 (NOT sheet 3)
    _clear_banding(sheets_service, sid, [sheet1_id, sheet2_id])
    for tab in [TAB_CORRECTED, TAB_SIS]:
        _retry_api(
            lambda t=tab: sheets_service.spreadsheets()
            .values()
            .clear(spreadsheetId=sid, range=f"'{t}'!A:Z")
            .execute()
        )
    # Clear _Lists
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .clear(spreadsheetId=sid, range="'_Lists'!A:Z")
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

    # ── Write _Lists (hidden tab with unique values) ──────────────────
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

    # ── Build Sheet 1 rows ────────────────────────────────────────────
    # Row 1: filter labels, Row 2: dropdown values, Row 3: headers, Row 4+: data
    header1 = ["\u2713"] + OUTPUT_FIELDS + ["Mismatch Summary"]
    data_rows1 = []
    for rec in corrections_map:
        row = [False]
        for field in OUTPUT_FIELDS:
            row.append(rec.get(field, ""))
        row.append(rec.get("mismatch_summary", ""))
        data_rows1.append(row)

    # ── Build Sheet 2 rows ────────────────────────────────────────────
    header2 = OUTPUT_FIELDS[:]
    data_rows2 = []
    for rec in corrections_sis:
        row = [rec.get(field, "") for field in OUTPUT_FIELDS]
        data_rows2.append(row)

    # ── Write data (rows 3+ = header + data, rows 1-2 written via vals) ──
    # Sheet 1
    sheet1_vals = [header1] + data_rows1
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range=f"'{TAB_CORRECTED}'!A3",
            valueInputOption="RAW",
            body={"values": sheet1_vals},
        )
        .execute()
    )
    # Sheet 2
    sheet2_vals = [header2] + data_rows2
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range=f"'{TAB_SIS}'!A3",
            valueInputOption="RAW",
            body={"values": sheet2_vals},
        )
        .execute()
    )

    # ── Write filter labels + defaults (rows 1-2) ────────────────────
    filter_keys = ["campus", "grade", "level", "student_group", "guide_email"]
    filter_labels = ["Campus", "Grade", "Level", "Student Group", "Guide Email"]

    # Sheet 1: filters start at col B (col A reserved for checkbox header area)
    s1_label_row = [""] + _build_filter_label_row(filter_labels, len(header1) - 1)
    s1_value_row = [""] + _build_filter_value_row(len(header1) - 1)
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range=f"'{TAB_CORRECTED}'!A1",
            valueInputOption="RAW",
            body={"values": [s1_label_row, s1_value_row]},
        )
        .execute()
    )

    # Sheet 2: filters start at col A
    s2_label_row = _build_filter_label_row(filter_labels, len(header2))
    s2_value_row = _build_filter_value_row(len(header2))
    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range=f"'{TAB_SIS}'!A1",
            valueInputOption="RAW",
            body={"values": [s2_label_row, s2_value_row]},
        )
        .execute()
    )

    # ── Format everything ─────────────────────────────────────────────
    fmt = []
    num_rows1 = len(sheet1_vals) + 2  # +2 for filter rows
    num_rows2 = len(sheet2_vals) + 2
    nc1 = len(header1)
    nc2 = len(header2)

    fmt.extend(
        _format_filter_rows(
            sheet1_id,
            nc1,
            FILTER_DEFS_SHEET1,
            filter_keys,
            list_ranges,
            has_checkbox=True,
        )
    )
    fmt.extend(
        _format_filter_rows(
            sheet2_id,
            nc2,
            FILTER_DEFS_SHEET2,
            filter_keys,
            list_ranges,
            has_checkbox=False,
        )
    )
    fmt.extend(_format_data_area(sheet1_id, num_rows1, nc1, has_checkbox=True))
    fmt.extend(_format_data_area(sheet2_id, num_rows2, nc2, has_checkbox=False))

    # Checkbox data validation (Sheet 1, col A, rows 4+)
    if data_rows1:
        fmt.append(
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet1_id,
                        "startRowIndex": 3,
                        "endRowIndex": 3 + len(data_rows1),
                        "startColumnIndex": 0,
                        "endColumnIndex": 1,
                    },
                    "rule": {"condition": {"type": "BOOLEAN"}, "showCustomUi": True},
                }
            }
        )

    # Hide _Lists tab
    fmt.append(
        {
            "updateSheetProperties": {
                "properties": {"sheetId": lists_id, "hidden": True},
                "fields": "hidden",
            }
        }
    )

    _retry_api(
        lambda: sheets_service.spreadsheets()
        .batchUpdate(spreadsheetId=sid, body={"requests": fmt})
        .execute()
    )

    # ── Ensure Sheet 3 has headers + filters ──────────────────────────
    _ensure_approved_headers(sheets_service, sid, sheet3_id, list_ranges, filter_keys)

    print(f"  Done — {len(corrections_map)} students written to corrections sheet.")


def _build_filter_label_row(labels, total_cols):
    """Build row 1: filter label names spread across columns, rest empty."""
    row = [""] * total_cols
    # Place 5 labels in positions 0, 2, 4, 6, 8 (every 2 columns)
    for i, label in enumerate(labels):
        pos = i * 2
        if pos < total_cols:
            row[pos] = label
    return row


def _build_filter_value_row(total_cols):
    """Build row 2: 'All' defaults for dropdown positions, rest empty."""
    row = [""] * total_cols
    for i in range(5):
        pos = i * 2
        if pos < total_cols:
            row[pos] = "All"
    return row


def _format_filter_rows(
    sheet_id, num_cols, filter_defs, filter_keys, list_ranges, has_checkbox
):
    """Format rows 1-2 as dashboard-style filter area."""
    fmt = []
    offset = 1 if has_checkbox else 0

    # Row 1 (index 0): dark navy background, full width
    fmt.append(
        _rc(
            sheet_id,
            0,
            1,
            0,
            num_cols,
            {
                "backgroundColorStyle": {"rgbColor": NAVY_DARK},
                "textFormat": {
                    "fontFamily": "Arial",
                    "fontSize": 9,
                    "bold": True,
                    "foregroundColorStyle": {"rgbColor": GREY_LABEL},
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "BOTTOM",
            },
        )
    )
    fmt.append(_rh(sheet_id, 0, 22))

    # Row 2 (index 1): filter background, full width
    fmt.append(
        _rc(
            sheet_id,
            1,
            2,
            0,
            num_cols,
            {
                "backgroundColorStyle": {"rgbColor": FILTER_BG},
                "textFormat": {
                    "fontFamily": "Arial",
                    "fontSize": 11,
                    "bold": True,
                    "foregroundColorStyle": {"rgbColor": WHITE},
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            },
        )
    )
    fmt.append(_rh(sheet_id, 1, 34))

    # Add merged cells + data validation for each filter dropdown
    filter_key_map = {
        "Campus": "campus",
        "Grade": "grade",
        "Level": "level",
        "Student Group": "student_group",
        "Guide Email": "guide_email",
    }

    for i, (label, _data_col) in enumerate(filter_defs):
        c1 = offset + i * 2
        c2 = c1 + 2
        if c2 > num_cols:
            break

        # Merge label cells (row 1)
        fmt.append(_mg(sheet_id, 0, 1, c1, c2))
        # Merge dropdown cells (row 2)
        fmt.append(_mg(sheet_id, 1, 2, c1, c2))

        # Data validation dropdown
        key = filter_key_map.get(label, "")
        if key and key in list_ranges:
            fmt.append(_dd(sheet_id, 1, c1, list_ranges[key]))

    return fmt


def _format_data_area(sheet_id, num_rows, num_cols, has_checkbox):
    """Format row 3 (header) and rows 4+ (data)."""
    fmt = []

    # Row 3 (index 2): column headers — navy background, white bold
    fmt.append(
        _rc(
            sheet_id,
            2,
            3,
            0,
            num_cols,
            {
                "backgroundColorStyle": {"rgbColor": FILTER_BG},
                "textFormat": {
                    "fontFamily": "Arial",
                    "fontSize": 10,
                    "bold": True,
                    "foregroundColorStyle": {"rgbColor": WHITE},
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            },
        )
    )
    fmt.append(_rh(sheet_id, 2, 30))

    # Freeze rows 1-3
    fmt.append(
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 3},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        }
    )

    # Alternating row colors (rows 3+)
    if num_rows > 3:
        fmt.append(
            {
                "addBanding": {
                    "bandedRange": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 2,
                            "endRowIndex": num_rows,
                            "startColumnIndex": 0,
                            "endColumnIndex": num_cols,
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

    # Column widths
    data_start = 1 if has_checkbox else 0
    if has_checkbox:
        fmt.append(_cw(sheet_id, 0, 30))  # checkbox column

    field_widths = [150, 60, 80, 100, 100, 220, 150, 100, 100, 220, 120, 140]
    for i, w in enumerate(field_widths):
        fmt.append(_cw(sheet_id, data_start + i, w))

    if has_checkbox:
        fmt.append(
            _cw(sheet_id, data_start + len(field_widths), 200)
        )  # mismatch summary

    return fmt


def _ensure_approved_headers(sheets_service, sid, sheet3_id, list_ranges, filter_keys):
    """Write headers + filter rows to Sheet 3 if empty."""
    resp = _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range=f"'{TAB_APPROVED}'!A3:A3")
        .execute()
    )
    if resp.get("values"):
        return  # already has content

    header = ["Date Approved"] + OUTPUT_FIELDS

    # Row 1: filter labels, Row 2: dropdowns, Row 3: header
    filter_labels = ["Campus", "Grade", "Level", "Student Group", "Guide Email"]
    label_row = _build_filter_label_row(filter_labels, len(header))
    value_row = _build_filter_value_row(len(header))

    _retry_api(
        lambda: sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sid,
            range=f"'{TAB_APPROVED}'!A1",
            valueInputOption="RAW",
            body={"values": [label_row, value_row, header]},
        )
        .execute()
    )

    # Format
    fmt = []
    nc = len(header)
    filter_key_map = {
        "Campus": "campus",
        "Grade": "grade",
        "Level": "level",
        "Student Group": "student_group",
        "Guide Email": "guide_email",
    }

    fmt.extend(
        _format_filter_rows(
            sheet3_id,
            nc,
            FILTER_DEFS_SHEET3,
            filter_keys,
            list_ranges,
            has_checkbox=False,
        )
    )

    # Header row format
    fmt.append(
        _rc(
            sheet3_id,
            2,
            3,
            0,
            nc,
            {
                "backgroundColorStyle": {"rgbColor": FILTER_BG},
                "textFormat": {
                    "fontFamily": "Arial",
                    "fontSize": 10,
                    "bold": True,
                    "foregroundColorStyle": {"rgbColor": WHITE},
                },
                "horizontalAlignment": "CENTER",
            },
        )
    )
    fmt.append(
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet3_id,
                    "gridProperties": {"frozenRowCount": 3},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        }
    )

    _retry_api(
        lambda: sheets_service.spreadsheets()
        .batchUpdate(spreadsheetId=sid, body={"requests": fmt})
        .execute()
    )
