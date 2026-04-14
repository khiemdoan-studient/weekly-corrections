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
| `sheets_writer.py` | ~680 | Sheets API: hidden tabs, QUERY formulas, title/caption, filters, format |
| `apps_script/Code.gs` | ~80 | Apps Script onEdit: checkbox → Sheet 3, clear checkboxes on filter change |
| `write_user_guide.py` | ~215 | Google Docs API: write formatted user guide |
| `run_export.ps1` | ~70 | One-time: Athena CTAS → S3 → GCS → BQ for alpha_roster |
| `alpha_roster_ctas.sql` | ~30 | Athena SQL with dedup, handles reserved word "group" |

## Spreadsheet Architecture

### Hidden Tabs (written by Python, never shown to users)
- `_CorrData` — Raw MAP roster data for mismatched students (13 cols, no headers)
- `_SISData` — Raw SIS data for same students (12 cols, no headers)
- `_Lists` — Unique values for dropdown data validation (5 cols: campus, grade, level, student_group, guide_email)

### Sheet 1: "Corrected Roster Info" (MAP data + checkboxes)
| Row | Content |
|-----|---------|
| 1 | Title: "Corrected Roster Info" (merged, navy dark, 20pt bold white) |
| 2 | Caption with clickable "User Guide" hyperlink (merged, navy med, italic grey) |
| 3 | Spacer (5px, dark) |
| 4 | Filter labels: Campus, Grade, Level, Student Group, Guide Email (merged pairs, dark) |
| 5 | Dropdown values: All (merged pairs, teal bg, data validation from _Lists) |
| 6 | Column headers: check, Campus, Grade, ... Mismatch Summary (navy, bold white) |
| 7+ | QUERY formula in B7 (filtered from _CorrData), checkboxes in A7+ |

### Sheet 2: "Current Roster Info in SIS" (same layout, no checkboxes)
Same title/caption/filter rows. QUERY formula in A7 pulls from `_SISData`.

### Sheet 3: "Automated Correction List" (cumulative, managed by Apps Script)
Title/caption rows, headers in row 6, data from row 7+ (appended by Apps Script).

## Filtering Mechanism

**Dropdowns filter via QUERY formulas — NOT Apps Script.**

The QUERY formula pattern (from `sheets_builder.py`):
```
=IFERROR(QUERY('_CorrData'!A:M, 
  "SELECT * WHERE 1=1"
  & IF($B$5="All", "", " AND Col1='" & $B$5 & "'")
  & IF($D$5="All", "", " AND Col2='" & $D$5 & "'")
  ..., 0), "")
```

When a dropdown in row 5 changes, the QUERY auto-recalculates and shows only matching rows. "All" means no filter for that field.

**Checkbox handling**: When a filter dropdown changes (row 5), the Apps Script `onEdit` clears all checkboxes in column A because the QUERY output shifts and checkboxes become stale.

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
4. **Checkbox stale-clearing** — When a dropdown changes, Apps Script clears all col A checkboxes because QUERY output rows shift.
5. **Header auto-detection** — Handles schema differences across campus sheets.

## Common Bugs & Fixes

| Symptom | Cause | Fix |
|---------|-------|-----|
| textFormatRun index error | "User Guide" at end of string, trailing run index = len(text) | Only add trailing run if `link_end < len(text)` |
| addBanding error on re-run | Existing banding not cleared | `_clear_banding()` runs before formatting |
| CTAS fails with "mismatched input 'group'" | Reserved word quotes stripped by shell | Read SQL from file, pass via `--cli-input-json` |
| Campus shows 0 enrolled | Notes column at different index | Header auto-detection via `MAP_HEADER_MAP` |
| Dropdown doesn't filter | Using Apps Script hideRows instead of QUERY | Rewrote to use QUERY formulas referencing dropdown cells |
