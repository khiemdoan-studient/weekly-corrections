# AI Instructions — Weekly Corrections

## Project Overview

This tool compares student enrollment data between two sources:
1. **MAP Roster** (Google Sheet) — the source of truth for student enrollment
2. **SIS Pipeline** (BigQuery `alpha_roster` table) — deduped export from Athena `alpha_student`

Students with mismatched data appear in a corrections spreadsheet for manager review.

## Critical File Map

| File | Lines | Purpose |
|------|-------|---------|
| `generate_corrections.py` | ~350 | Main orchestrator: auth, read MAP, query BQ, compare, write |
| `config.py` | ~90 | Constants, header mappings, campus list, sheet IDs |
| `queries.py` | ~30 | Single BQ query function for alpha_roster |
| `sheets_writer.py` | ~850 | Sheets API: hidden tabs, QUERY formulas, title/caption, filters, format |
| `apps_script/Code.gs` | ~105 | Apps Script onEdit: accept/reject checkboxes → route by type, clear on filter change |
| `write_user_guide.py` | ~215 | Google Docs API: write formatted user guide |
| `run_export.ps1` | ~70 | One-time: Athena CTAS → S3 → GCS → BQ for alpha_roster |
| `alpha_roster_ctas.sql` | ~30 | Athena SQL with dedup, handles reserved word "group" |
| `setup_unenroll_columns.py` | ~150 | One-time: provision Unenroll columns on all 9 ISRs + CMR. Idempotent — safe to re-run |

## Spreadsheet Architecture

### Hidden Tabs (written by Python, never shown to users)
- `_CorrData` — Raw MAP roster data for mismatched students (13 cols, no headers)
- `_SISData` — Raw SIS data for same students (12 cols, no headers)
- `_Lists` — Unique values for filter dropdowns (cols A-E) + sort options (cols F-K, one per sheet)
- `_ApprovedData` — Cumulative approved field-mismatch corrections (14 cols: Date, MismatchSummary, 12 fields)
- `_AdditionsData` — Cumulative approved roster additions (14 cols, same layout)
- `_UnenrollData` — Cumulative approved unenrollments (14 cols, same layout)
- `_RejectedData` — Cumulative rejected changes (14 cols, same layout)

### Sheet 1: "Corrected Roster Info" (MAP data + accept/reject checkboxes)
| Row | Content |
|-----|---------|
| 1 | Title: "Corrected Roster Info" (merged, navy dark, 20pt bold white) |
| 2 | Caption with clickable "User Guide" hyperlink (merged, navy med, italic grey) |
| 3 | Spacer (5px, dark) |
| 4 | Filter labels + "SORT BY" label (merged pairs, dark) |
| 5 | Dropdown values + Sort By dropdown (merged pairs, teal bg, data validation from _Lists) |
| 6 | Column headers: Accept Changes, Reject Changes, Campus, Grade, ... Mismatch Summary (navy, bold white) |
| 7+ | SORT(QUERY()) formula in C7 (filtered + sorted from _CorrData), checkboxes in A7+ (green) and B7+ (red) |

### Sheet 2: "Current Roster Info in SIS" (same layout, no checkboxes)
Same title/caption/filter/sort rows. SORT(QUERY()) formula in A7 pulls from `_SISData`.

### Sheets 3-5: Approval sheets (14 cols: Date + Mismatch Summary + 12 fields)
Same title/caption/filter/sort rows. SORT(QUERY()) formula in A7 pulls from hidden cumulative tab (A:N, 14 cols). Mismatch Summary is column B with red header. QUERY col refs: Campus=Col3, Grade=Col4, Level=Col5, StudentGroup=Col9, GuideEmail=Col12.
- Sheet 3 "Automated Correction List" — reads `_ApprovedData` (field mismatches)
- Sheet 4 "Roster Additions" — reads `_AdditionsData` ("Roster Addition" type)
- Sheet 5 "Roster Unenrollments" — reads `_UnenrollData` ("Unenrolling" type)

### Sheet 6: "Rejected Changes" (15 cols: Date + Mismatch Summary + 12 fields + Reason)
Same as Sheets 3-5 but with extra "Reason for Rejection" column (col O, blank for manual entry). QUERY reads `_RejectedData` A:N (14 cols); Reason is outside QUERY output.

## ISR (Individual Student Roster) Architecture

Each of 9 campuses has a dedicated Google Sheet called an ISR (Individual Student Roster). The ISRs feed the CMR (Consolidated MAP Roster), which in turn feeds the pipeline.

### ISR Tab Layout
- **`Student Roster` (SR)** — IMs edit here directly. Has an Unenroll checkbox column.
- **`MAP Roster` (MR)** — formula-derived view, not manually maintained. Values pull from:
  - `SR` via cell references (e.g. `='Student Roster'!H2`)
  - `Auto Synced from SIS` via VLOOKUP (Notes col O = SIS admission status)
  - Some plain values
  - MR is the tab that feeds the CMR via IMPORTRANGE.
- **`Auto Synced from SIS`** — SIS admission status dump (not directly maintained by IMs).
- **`School Info`** — Campus name to ID mapping for student_id derivation.

### Data Flow
```
SR (IM edits) ──formulas──> MR ──IMPORTRANGE──> CMR ──Python Sheets API──> generate_corrections.py
```

### ISR_CONFIG in config.py
Per-campus column positions are hard-coded in `config.py::ISR_CONFIG` with keys:
- `isr_id` — Google Sheet ID of the ISR
- `mr_gid` — gid of the `MAP Roster` tab
- `sr_unenroll_col` — 0-indexed column position of Unenroll on the SR tab
- `mr_unenroll_col` — 0-indexed column position of Unenroll on the MR tab

All 9 campuses are listed, each with their own SR/MR Unenroll column indices (positions vary because column layouts differ across campus sheets).

## Unenroll Workflow (option-C precedence)

Every student record in `read_map_roster` gets a `_unenroll_flag` boolean read from the CMR Unenroll column. In `compare_students`, for each ENROLLED MAP student, the logic is:

1. **FIRST check (option-C precedence)**: if `_unenroll_flag=TRUE` AND SIS `admissionstatus=Enrolled` → flag as `"Unenrolling"` and skip all other checks.
2. **Fall through**: otherwise, run existing checks — Roster Addition, Field mismatch, Notes-based Unenrolling.

**Option-C precedence** means the IM-driven Unenroll checkbox trumps field mismatches — we'd rather process the unenrollment than flag a grade/email change on a student who is about to be unenrolled anyway.

The pipeline prints a breakdown to stdout:
```
Unenrolling: N (IM-flagged: X, Notes-based: Y)
```

## Filtering & Sorting Mechanism

**Dropdowns filter AND sort via SORT(QUERY()) formulas — NOT Apps Script.**

The formula pattern for Sheet 1 (offset by 2 for accept/reject columns):
```
=IFERROR(SORT(QUERY('_CorrData'!A:M, 
  "SELECT * WHERE 1=1"
  & IF($C$5="All", "", " AND Col1='" & $C$5 & "'")
  & IF($E$5="All", "", " AND Col2='" & $E$5 & "'")
  ..., 0),
  MATCH($M$5, _Lists!F$2:F$14, 0),
  IF(OR($M$5="Grade"), FALSE, TRUE)), "")
```

- QUERY handles filtering (WHERE clauses reference dropdown cells)
- SORT wraps QUERY output, using MATCH to find column index from Sort By dropdown
- Sort direction: ascending (TRUE) for all text fields, descending (FALSE) for Grade and Date Approved
- Sort options stored in `_Lists` columns F-K (one per visible sheet)

**Checkbox handling**: When a filter or Sort By dropdown changes (row 5), Apps Script clears both Accept (col A) and Reject (col B) checkboxes because the QUERY output shifts.

**Accept/Reject workflow on Sheet 1**:
- **Col A (Accept Changes)** — light green (#D4EDDA) background. When checked, Apps Script reads col O (Mismatch Summary) and routes to:
  - `"Roster Addition"` → `_AdditionsData` → Sheet 4 auto-updates via QUERY
  - `"Unenrolling"` → `_UnenrollData` → Sheet 5 auto-updates via QUERY
  - field mismatches → `_ApprovedData` → Sheet 3 auto-updates via QUERY
- **Col B (Reject Changes)** — light red (#FEE2E2) background. When checked, all rejected rows → `_RejectedData` → Sheet 6 auto-updates via QUERY
- **Mutual exclusion**: Checking Accept unchecks Reject and vice versa
- Data columns read from C:N (12 cols), Mismatch Summary in col O (column 15)

## Mismatch Types

| Type | Mismatch Summary | Condition | Color |
|------|-----------------|-----------|-------|
| Roster Addition | "Roster Addition" | Enrolled in MAP, student_id not in SIS | Light green (#D4EDDA) |
| Field Mismatch | "Grade, Email" etc. | Enrolled in both, fields differ | Yellow (#FFF3CD) |
| Unenrolling | "Unenrolling" | Notes != "Enrolled" in MAP, admissionstatus = "Enrolled" in SIS | Light yellow (#FFFDE7) |

Colors are applied via conditional formatting rules on the Mismatch Summary column (Sheet 1 only), in priority order: Roster Addition → Unenrolling → NOT_BLANK (field mismatches).

## Data Sources

### MAP Roster (Google Sheet)
- ID: `1scEay0a8OR6vU3uJuxbHKWCEx_RVgSsRXF9naJh3XYw`
- 9 campus sheets with "(Dash)" suffix
- **Column layout varies by campus** — header auto-detection via `MAP_HEADER_MAP`
- Reading CCSD has extra "Full Name" column (Notes at col O instead of N)
- Corrupt header fallback: if col A isn't "Student ID" but 5+ other cols detected, assumes col A

### BigQuery alpha_roster
- Table: `studient-flat-exports-doan.studient_analytics.alpha_roster`
- ~8,200 students after ROW_NUMBER dedup from `alpha_student`
- All admission statuses included (not just Enrolled)

## Column Mapping (MAP to SIS)

| Field | MAP Header | SIS Column | QUERY Col |
|-------|-----------|------------|-----------|
| Campus | Campus | campus | Col1 |
| Grade | Grade | gradelevel | Col2 |
| Level | Level | alphalevellong | Col3 |
| First Name | First Name | firstname | Col4 |
| Last Name | Last Name | lastname | Col5 |
| Email | Student Email | email | Col6 |
| Student Group | School, if separate from Campus | student_group | Col7 |
| Guide First Name | Teacher 1 First Name | SPLIT(advisor)[0] | Col8 |
| Guide Last Name | Teacher 1 Last Name | SPLIT(advisor)[1:] | Col9 |
| Guide Email | Teacher 1 Email | advisoremail | Col10 |
| Student_ID | Student ID | fullid | Col11 |
| External Student ID | SUNS Number | externalstudentid | Col12 |

## Key Design Decisions

1. **QUERY formulas for filtering** — Apps Script `hideRows()` approach failed because it couldn't map dropdown positions to data columns. QUERY formulas auto-recalculate when dropdown cells change.
2. **Hidden data tabs** — Raw data on `_CorrData`/`_SISData`, QUERY on visible tabs. Same pattern as the dashboard pipeline's `_Data` tab.
3. **textFormatRuns for User Guide link** — The caption uses `updateCells` with `textFormatRuns` to create a clickable "User Guide" hyperlink mid-text.
4. **Checkbox stale-clearing** — When a dropdown changes, Apps Script clears both col A and col B checkboxes because QUERY output rows shift.
5. **Header auto-detection** — Handles schema differences across campus sheets.
6. **Batched API pre-writes** — Tab existence (1 read + 1 batch create), unmergeCells (1 batched call for all 6 sheets), clear values (`batchClear` for 9 tabs). ~5 pre-write API calls total instead of ~28.
7. **Three-level unenroll chain** — IM checks SR checkbox → formula mirrors to MR → IMPORTRANGE pushes to CMR → pipeline reads CMR Unenroll via `MAP_HEADER_MAP` auto-detection. No Apps Script needed.

## Common Bugs & Fixes

| Symptom | Cause | Fix |
|---------|-------|-----|
| QUERY returns error when name has apostrophe | Single quotes in values (O'Brien) break QUERY string | `_sq()` wraps cell refs with `SUBSTITUTE(cell,"'","''")` |
| Sheet 3 filters show wrong students | `_ApprovedData` has Date in Col1, shifting all cols by 1 | Fixed: Grade=Col3, Level=Col4, StudentGroup=Col8, GuideEmail=Col11 |
| Date Approved shows serial number (46126.5) | QUERY strips number formatting from source | `repeatCell` with `numberFormat` DATE_TIME on Sheet 3 col A |
| textFormatRun index error | "User Guide" at end of string | Only add trailing run if `link_end < len(text)` |
| addBanding error on re-run | Existing banding not cleared | `_clear_banding()` runs before formatting |
| CTAS fails with "mismatched input 'group'" | Reserved word quotes stripped by shell | Read SQL from file, pass via `--cli-input-json` |
| Campus shows 0 enrolled | Notes column at different index | Header auto-detection via `MAP_HEADER_MAP` |
| Dropdown doesn't filter | Using Apps Script hideRows instead of QUERY | Rewrote to use QUERY formulas referencing dropdown cells |
| Slow writes (~35s) | 15+ individual API calls for value ranges | Batched into 2 calls via `values().batchUpdate()` (~16s) |
| Slow pre-writes (~15s) | 28 individual API calls for tab checks, unmerge, clear | Batched into ~5 calls via `_ensure_all_tabs`, batch `unmergeCells`, `batchClear` |
| Duplicate student IDs silently overwritten | Same ID in multiple campus sheets | Warning logged; last-write-wins preserved |
| mergeCells error on re-run | Old merged cells conflict with new layout | `unmergeCells` runs before formatting on all visible sheets |
| "Invalid requests[0].setDataValidation: This operation is not allowed on cells in typed columns" | Column inside a Google Sheets Table (typed column feature) | `setup_unenroll_columns.py` catches this and skips `setDataValidation` (Table already provides checkbox rendering) |
| "Range exceeds grid limits. Max rows: X, max columns: 27" | SR tabs only had 27 cols, can't write to col 28 | Script auto-expands grid via `appendDimension` COLUMNS |

## Known Limitations

- **Grade sorts lexicographically** — "10" sorts before "2" because QUERY treats grades as text. Numeric sorting would require a helper column.
- **Cumulative hidden tabs grow indefinitely** — `_ApprovedData`, `_AdditionsData`, `_UnenrollData`, `_RejectedData` are never cleared. Manual cleanup needed periodically.
- **Banding covers 200 data rows** — Alternating row colors extend to row 206. If cumulative sheets exceed 200 approved rows, increase the `end_row` floor in `_format_visible_sheet()`.
- **Reason for Rejection column not in QUERY output** — Sheet 6 QUERY reads from `_RejectedData` A:M (13 cols). The "Reason for Rejection" header is in col N (col 14) but the column content is manually entered by IMs, not populated by QUERY.
