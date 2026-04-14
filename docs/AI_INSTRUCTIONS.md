# AI Instructions — Weekly Corrections

## Project Overview

This tool compares student enrollment data between two sources:
1. **MAP Roster** (Google Sheet) — the source of truth for student enrollment
2. **SIS Pipeline** (BigQuery `alpha_roster` table) — deduped export from Athena `alpha_student`

Students with mismatched data appear in a corrections spreadsheet for manager review.

## Critical File Map

| File | Lines | Purpose |
|------|-------|---------|
| `generate_corrections.py` | ~300 | Main orchestrator: auth, read MAP, query BQ, compare, write |
| `config.py` | ~90 | Constants, header mappings, campus list, sheet IDs |
| `queries.py` | ~30 | Single BQ query function for alpha_roster |
| `sheets_writer.py` | ~280 | Sheets API: clear, write, format, checkboxes, banding |
| `apps_script/Code.gs` | ~55 | Apps Script onEdit: checkbox → Sheet 3 with date |
| `run_export.ps1` | ~70 | One-time: Athena CTAS → S3 → GCS → BQ for alpha_roster |
| `alpha_roster_ctas.sql` | ~30 | Athena SQL with dedup (ROW_NUMBER), handles reserved word "group" |

## Spreadsheet Architecture

### Sheet 1: "Corrected Roster Info" (MAP data for mismatched students)
| Col | Field | Notes |
|-----|-------|-------|
| A | Checkbox | TRUE/FALSE, data validation BOOLEAN |
| B | Campus | |
| C | Grade | |
| D | Level | |
| E | First Name | |
| F | Last Name | |
| G | Email | |
| H | Student Group | = "School, if separate from Campus" in MAP |
| I | Guide First Name | Teacher 1 First Name |
| J | Guide Last Name | Teacher 1 Last Name |
| K | Guide Email | Teacher 1 Email |
| L | Student_ID | Join key (format "088-11901") |
| M | External Student ID | SUNS Number (SC campuses only) |
| N | Mismatch Summary | Comma-separated list of differing fields |

### Sheet 2: "Current Roster Info in SIS" (same students, same order)
Columns A-L mirror Sheet 1 columns B-M (no checkbox, no mismatch summary).

### Sheet 3: "Automated Correction List" (cumulative)
Column A = Date Approved, Columns B-M = same as Sheet 1 B-M.
Managed by Apps Script onEdit trigger. Python never writes to this sheet.

## Data Sources

### MAP Roster (Google Sheet)
- ID: `1scEay0a8OR6vU3uJuxbHKWCEx_RVgSsRXF9naJh3XYw`
- 9 campus sheets, each with "(Dash)" suffix
- **Column layout varies by campus** — Reading CCSD has an extra "Full Name" column that shifts indices
- Header auto-detection via `MAP_HEADER_MAP` in config.py maps field names → possible header strings
- Only students with Notes = "Enrolled" are included
- Corrupt headers handled: if col A header isn't "Student ID" but other columns match, col A is assumed

### BigQuery alpha_roster
- Table: `studient-flat-exports-doan.studient_analytics.alpha_roster`
- Created by CTAS from `studient.alpha_student` with ROW_NUMBER dedup
- Dedup priority: Enrolled status first, then group-populated rows, then fullid ASC
- All admission statuses included (not just Enrolled) to detect enrollment mismatches
- ~8,200 students after dedup

## Column Mapping (MAP → SIS)

| Field | MAP Header | SIS Column |
|-------|-----------|------------|
| Campus | Campus | campus |
| Grade | Grade | gradelevel |
| Level | Level | alphalevellong |
| First Name | First Name | firstname |
| Last Name | Last Name | lastname |
| Email | Student Email | email |
| Student Group | School, if separate from Campus / School_Name | student_group (= alpha_student.group) |
| Guide First Name | Teacher 1 First Name / Teacher_First | SPLIT(advisor)[0] |
| Guide Last Name | Teacher 1 Last Name / Teacher_Last | SPLIT(advisor)[1:] |
| Guide Email | Teacher 1 Email / Teacher_Email | advisoremail |
| Student_ID | Student ID | fullid |
| External Student ID | SUNS Number | externalstudentid |

## Comparison Logic

1. Join on Student_ID (MAP col A = BQ fullid, format "088-11901")
2. Normalize before comparing: `_normalize()` → strip, lowercase, collapse whitespace
3. Guide name: combine MAP first+last → compare to SIS advisor (single field)
4. If ANY field differs → include in both Sheet 1 and Sheet 2
5. Students in MAP but not BQ → Sheet 2 shows "NOT FOUND IN SIS"
6. Mismatch Summary = comma-separated list of differing field names

## Key Design Decisions

1. **Header auto-detection** over hard-coded indices — Reading CCSD has a different column layout (extra "Full Name" column, Notes at col O instead of N)
2. **MAP roster is source of truth** — only MAP-enrolled students are compared; SIS-only students are ignored
3. **Sheet 3 is never touched by Python** — Apps Script manages it; cumulative history
4. **alpha_roster includes all admission statuses** — enables detecting "Enrolled in MAP but Former Student in SIS"
5. **Reserved word "group"** in alpha_student — CTAS uses subquery to alias it, SQL read from file to preserve double-quote escaping through PowerShell → AWS CLI

## Pipeline Integration

Step 11b in `Refresh-Data.ps1` (Studient Excel Automation project) exports alpha_roster:
1. CTAS from `studient.alpha_student` → Parquet on S3
2. S3 → local → GCS → BigQuery load via `bq_load.py`
3. Uses `--cli-input-json` to preserve SQL double-quote escaping

## Common Bugs & Fixes

| Symptom | Cause | Fix |
|---------|-------|-----|
| Campus shows 0 enrolled | Notes column at different index | Add header variants to `MAP_HEADER_MAP` |
| `addBanding` error on re-run | Existing banding not cleared | `_clear_banding()` runs before formatting |
| CTAS fails with "mismatched input 'group'" | Reserved word double-quotes stripped by shell | Read SQL from file, pass via `--cli-input-json` |
| Col A header corrupt (e.g. "4") | Someone edited the header cell | Fallback: assume col A = student_id if 5+ other cols detected |
