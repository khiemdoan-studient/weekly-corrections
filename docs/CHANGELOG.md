# Changelog

## [v2.4.5] — 2026-04-23

### Fixed
- **`sheets_writer.py:1270` — hardcoded 1006-row ceiling on cumulative-tab date formatting** — The date-format `repeatCell` on Sheets 3/4/5/6 (col A) stopped at `endRowIndex: 1006`. Once any cumulative tab (`_ApprovedData`, `_AdditionsData`, `_UnenrollData`, `_RejectedData`) grew past ~1000 rows, new entries would render as raw numbers (e.g. `45768.0557`) instead of `yyyy-MM-dd HH:mm:ss`, breaking chronological sort on approval sheets. Raised ceiling to `10_000` — ~2 years of headroom at 200 mismatches/week.
- **`sheets_writer.py:1488` — hardcoded 206-row banding floor on Sheets 3-6** — `_format_visible_sheet` was called with `num_data_rows=0` for Sheets 3-6, so `end_row = max(6+0+5, 206) = 206`. Alternating row colors would stop at row 206 once cumulative tabs grew past ~200 rows. Raised floor to `2000`. Banding on empty rows is harmless.

### Why
Found by `/audit` after v2.4.4 shipped. Both are latent issues — current cumulative tabs are well under 1000 rows, so no visible effect yet. Fixes prevent future failures without requiring any user action.

### Verified
- Code-level only (constant bumps, no logic change). Next hourly GitHub Actions run will apply the new ranges to the live sheet automatically.

### User Action Required
- None. Fix takes effect on next hourly pipeline run.

## [v2.4.4] — 2026-04-22

### Added
- **Row-hiding on Sheet 1 for recently-handled students** — Accepted/Rejected rows now disappear from "Corrected Roster Info" on the next hourly pipeline run and stay hidden for 7 days. Eliminates the long-standing confusion where a student's checkbox was cleared on pipeline rebuild but the row kept reappearing until the data team updated SIS. Behavior after 7 days:
  - If the mismatch still exists in MAP vs SIS → student reappears on Sheet 1 (signal that the correction is stale and hasn't been processed)
  - If SIS was updated → student naturally stays absent (no mismatch to flag anyway)
- **`HIDE_HANDLED_DAYS = 7` in `config.py`** — Tunable window. Set to 0 to disable the filter entirely and restore the old always-show-everything behavior.
- **`generate_corrections.py:read_handled_student_ids(sheets_service, days_back)`** — Walks all 4 cumulative tabs (`_ApprovedData`, `_AdditionsData`, `_UnenrollData`, `_RejectedData`); parses canonical `yyyy-MM-dd HH:mm:ss` timestamps from col A; returns the set of `student_id`s handled within the cutoff window.
- **`generate_corrections.py:_hide_recently_handled(corrections_map, corrections_sis, handled_ids)`** — Filters the parallel lists, dropping any student whose ID is in `handled_ids`.

### Changed
- **`generate_corrections.py:main()`** — Now calls `read_handled_student_ids` + `_hide_recently_handled` between `compare_students` and `write_corrections`. Log line added: `Hidden N recently-handled students (within last 7 days). M corrections remain on Sheet 1.`

### Fixed (documentation / expectation reset)
- **Long-standing "rows stay visible forever" behavior** — Earlier docs (including my prior CHANGELOG entry in v2.4.3) said rows only disappear when SIS is updated and the pipeline re-runs. That matched the code but didn't match what users expected. Re-checked git history: no prior version ever hid handled rows (`hideRows()` in early versions was for dropdown filtering only). This release makes the product behave the way users already assumed it did.

### Verified
- Live pipeline run: 198 total corrections → 9 hidden (the test students Accept/Rejected earlier in this session) → **189 visible on Sheet 1**. Log output confirmed exact counts.
- Accept/Reject column colors on Sheet 1: `userEnteredFormat.backgroundColorStyle` verified via Sheets API after pipeline run:
  - Col A row 10: `#D4EDDA` (ACCEPT_BG light green) ✓
  - Col B row 10: `#FEE2E2` (REJECT_BG light red) ✓
  These colors persist because the v2.4.3 Code.gs only touches cols C:O on check/uncheck.

### User Action Required
- **None for v2.4.4 itself** — just let the next hourly GitHub Actions run pick it up. (Or trigger manually: https://github.com/khiemdoan-studient/weekly-corrections/actions)
- **Still outstanding from v2.4.3**: re-paste `apps_script/Code.gs` into Extensions > Apps Script if you haven't yet. Without it, checkbox clicks still wipe the Accept/Reject column colors and create stale cumulative-tab rows.

## [v2.4.3] — 2026-04-22

### Fixed
- **onOpen trigger deleted error** — User pasted v2.4.2 Code.gs (which had no `onOpen`) into an ISR that had a pre-existing Student Cards `onOpen` menu trigger. Google's trigger system kept firing but the function was gone. Resolved by merging the Student Cards generator (`onOpen` + `generateStudentCardsFromTemplate` + helpers) into the same Code.gs file. The merged Code.gs is safe to paste in either the corrections spreadsheet OR any ISR — each feature activates only when its sheet tab is present.
- **Duplicate rows across cumulative tabs when toggling** — User clicked Reject → unchecked → Accept. Student (Allyssa Fortiz-Santos 083-11509) ended up in BOTH `_RejectedData` and `_UnenrollData`. Root cause: `appendRow` always added, never removed. Fix: new `removeStudentFromCumulativeTabs_()` runs before every append, deleting any existing rows for that `student_id` across all 4 cumulative tabs. Idempotent — toggle as many times as you want, only the latest choice persists. Cleaned the Allyssa duplicate manually.
- **Accept/Reject column colors wiped on uncheck** — Previous version called `sheet.getRange(row, 1, 1, 15).setBackground(null)` which reset the permanent green/red on cols A/B. Fix: only touch cols C:O (cols 3–15) for backgrounds; cols A and B keep their column-level `ACCEPT_BG` / `REJECT_BG` applied by `sheets_writer.py`.

### Changed
- **Code.gs structure** — Now a single file with both features clearly delineated:
  - Feature 1: onEdit (accept/reject) — only activates on "Corrected Roster Info" sheet
  - Feature 2: onOpen (Student Cards menu) — only shows up on spreadsheets with a "Copy of MAP Roster" tab

### User Action Required
- Re-paste the latest `apps_script/Code.gs` into Extensions > Apps Script in your spreadsheet (wherever the `onOpen` error was firing). The paste is idempotent.

### Notes
- **How long do accepted/rejected rows take to disappear from Corrected Roster Info?** They don't disappear instantly. When you check Accept or Reject, the row is greyed out on Sheet 1 and the data is copied to the cumulative tab. The row only leaves Sheet 1 when (a) the data team updates SIS to match MAP and (b) the next hourly pipeline run re-reads the CMR + BQ and no longer sees a mismatch. Worst case: data team processes on Friday → row disappears from Sheet 1 within the hour. Best case: if SIS updates overnight, the row is gone by next morning's pipeline run.

## [v2.4.2] — 2026-04-22

### Fixed
- **Inconsistent date formats in cumulative tabs** — The old Code.gs pattern
  `appendRow([new Date(), ...])` + `setNumberFormat(getLastRow(), 1, "yyyy-MM-dd HH:mm:ss")`
  had a race: when multiple onEdit triggers fired nearly simultaneously (e.g. user
  checking 3 checkboxes in quick succession), `getLastRow()` returned different
  values across the concurrent triggers, so `setNumberFormat` sometimes applied
  to the wrong row. Result: some rows displayed as `4/23/2026 1:37:44` (locale
  default) while others got the intended `2026-04-23 01:37:44`, breaking the
  chronological sort on Automated Correction List / Roster Additions / Roster
  Unenrollments / Rejected Changes.

- **`apps_script/Code.gs` hardened** — Pre-formats the timestamp to an ISO
  string via `Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")`
  BEFORE `appendRow`. The row is now appended with the already-formatted string,
  eliminating the post-append `setNumberFormat` call and removing the race.

### Added
- **`normalize_dates.py`** — One-time migration that walks each cumulative tab,
  parses column A regardless of current format (handles both `M/D/YYYY H:MM:SS`
  and `YYYY-MM-DD HH:MM:SS`), rewrites as canonical `yyyy-MM-dd HH:mm:ss` ISO
  string, and applies TEXT number format to the column for display consistency.
  Idempotent — safe to re-run.

### Verified
- Post-normalization sort on Automated Correction List: `2026-04-23 01:37:47` →
  `2026-04-23 01:37:44` → `2026-04-17 10:54:27` → ... → `2026-04-14 12:02:37`.
  Perfectly chronological.

### User Action Required
- Re-paste the updated `apps_script/Code.gs` into Extensions > Apps Script to
  pick up the race-safe timestamp handling. (Existing rows already normalized
  by the one-time script above.)

## [v2.4.1] — 2026-04-22

### Changed
- **Unenrolling conditional formatting color** — Changed from light yellow (#FFFDE7) to light red (#FEE2E2) on Sheet 1's Mismatch Summary column. Matches the Reject checkbox column background color for visual consistency (red = "needs removal"). New constant `RED_LIGHT` added alongside existing `YELLOW_LIGHT` (now marked legacy). Also applied to the Unenroll Queue (Live) sheet data area background.
- **Docs updated**: `HUMAN_INSTRUCTIONS.md`, `AI_INSTRUCTIONS.md`, and the User Guide Google Doc all reflect the new color.

### Verified
- **First GitHub Actions run succeeded** — manual workflow_dispatch trigger completed in 37s, wrote 198 correction rows (192 existing + 6 IM-flagged Unenrolling from Hardeeville Elementary).

## [v2.4.0] — 2026-04-22

### Added
- **"Unenroll Queue (Live)" visible sheet** — new tab in the corrections spreadsheet (sheetId 1118002361) that shows IM-flagged Unenroll students across all 9 campuses in real-time (~1 min latency via IMPORTRANGE refresh). Each campus gets its own 50-row block with `QUERY(IMPORTRANGE(...), "SELECT ... WHERE Col{N} = TRUE")` formula that auto-expands as IMs check boxes. Does NOT include SIS comparison (the hourly Python pipeline handles that).
- **`build_unenroll_queue.py`** (new) — One-time idempotent setup script. Creates the tab, writes formulas, applies formatting (navy title, italic caption, yellow data bg).
- **Hourly GitHub Actions workflow** (`.github/workflows/hourly-pipeline.yml`) — Runs `generate_corrections.py` every hour at :00 UTC on ubuntu-latest. Loads `GCP_SA_KEY` from repo secret, writes to `keys/sa-main.json`, runs pipeline, cleans up credentials. Features: `workflow_dispatch` for manual runs, `concurrency` group for single-flight, 10-min timeout.

### Fixed
- **Boolean WHERE clause in QUERY** — `WHERE UPPER(Col{N}) = 'TRUE'` returned `#VALUE!` because UPPER can't apply to boolean. Fixed to `WHERE Col{N} = TRUE` (unquoted, direct boolean compare).
- **Reading CCSD grade column offset** — Because Reading CCSD has an extra "Full Name" column at position G, its Grade is at col H (Col8) not Col7. `build_campus_formula` checks the tab name and uses the right column.
- **`read_map_roster` read range expanded** from `A1:AC` to `A1:AE` so the pipeline actually picks up the Unenroll column at col AD (29) for 7 of 9 campuses. Previously the code silently missed 6 IM-flagged Hardeeville Elementary unenrollments.

### Operations
- **GCP_SA_KEY GitHub secret** (already installed) — contains verbatim `keys/sa-main.json`. Used only by the workflow; never committed.
- **One-time user action required**: Open "Unenroll Queue (Live)" tab in the corrections sheet and click 'Allow access' on the IMPORTRANGE prompt. After that, real-time updates work automatically.
- **Repo**: https://github.com/khiemdoan-studient/weekly-corrections

## [v2.3.0] — 2026-04-22

### Added
- **IM-driven Unenroll workflow spanning 3 sheet levels** — ISR (Individual Student Roster, 1 per campus, 9 total) → CMR (Combined MAP Roster, 9 campus tabs) → Pipeline (`generate_corrections.py` reads CMR, compares to BigQuery `alpha_roster`). IMs check an Unenroll checkbox on their campus SR tab; the MR tab mirrors via formula; the CMR pulls via IMPORTRANGE; the pipeline reads the flag and flags the student as "Unenrolling".
- **Per-campus `ISR_CONFIG` in `config.py`** — Maps each campus tab name to its ISR spreadsheet ID, MAP Roster gid, SR Unenroll column index, and MR Unenroll column index. Column mapping established:
  - Reading CCSD: SR col X[23], MR col AE[30]
  - Metro Schools: SR col Y[24], MR col AB[27]
  - Allendale Fairfax Elementary / Middle: SR col Z[25], MR col AB[27]
  - Allendale Aspire Academy: SR col AB[27], MR col AB[27]
  - Hardeeville Elementary / Junior-Senior High, Ridgeland Elementary / Secondary: SR col AB[27] (added), MR col AD[29]
- **`setup_unenroll_columns.py`** — One-time provisioning script: creates SR Unenroll checkbox column, MR mirror formula (`=ArrayFormula(SR!Xn)`), and CMR IMPORTRANGE formula across all 9 ISRs. Idempotent.
- **"unenroll" added to `MAP_HEADER_MAP`** — Recognizes headers "unenroll" and "unenrolled" when reading CMR.
- **`_unenroll_flag` on each student record** — `read_map_roster()` now reads the Unenroll column (TRUE/FALSE) into each student record for downstream comparison.
- **Option-C precedence in `compare_students()`** — If IM-flagged `Unenroll=TRUE` AND SIS `admissionstatus=Enrolled`, student is flagged as "Unenrolling" and this takes precedence over any field mismatches. Also prints a breakdown of IM-flagged vs Notes-based unenrollings at the end of the run.

### Fixed
- **Setup script edge cases** — Handles Google Sheets Tables "typed columns" (skips `setDataValidation` because the Table already provides checkbox rendering) and grid-size expansion (`appendDimension COLUMNS`) for SRs that only had 27 columns before the Unenroll column could be added.

### Notes
- Setup script successfully ran: all 9 ISRs provisioned, all 9 CMR Unenroll columns wired with IMPORTRANGE, rendering live FALSE values. 0 IM-flagged unenrollments so far, as expected since no IM has checked any boxes yet.

## [v2.2.0] — 2026-04-17

### Fixed
- **Column shift bug in approval/rejection sheets** — v2.0.0 Apps Script read from col 2 (Reject checkbox) instead of col 3, inserting `FALSE` as first data value and shifting all fields right by 1. ExtStudentID was lost. Root cause: user had not pasted the updated v2.1.0 Code.gs. Fixed by: migration function + updated Code.gs with correct column offsets.
- **Data migration for corrupted cumulative tabs** — `_migrate_cumulative_tabs()` detects and fixes 3 row formats: corrupted (13 cols with FALSE), old v2.0.0 (13 cols correct), and already-migrated (14 cols). Removes FALSE, inserts blank Mismatch Summary, pads missing ExtStudentID.

### Added
- **Mismatch Summary column on Sheets 3-6** — When a correction is accepted/rejected, the mismatch type (e.g. "Roster Addition", "Guide Name", "Unenrolling") is now stored as column B in all cumulative hidden tabs. Visible on all approval/rejection sheets as the 2nd column with red header formatting.
- **14-column layout for cumulative tabs** — _ApprovedData, _AdditionsData, _UnenrollData, _RejectedData now store: Date, MismatchSummary, Campus, Grade, Level, FirstName, LastName, Email, StudentGroup, GuideFirst, GuideLast, GuideEmail, StudentID, ExtStudentID.

### Changed
- **NC3/NC4/NC5 = 14** (was 13), **NC6 = 15** (was 14) — all visible sheets updated for new column
- **QUERY column references shifted** for Sheets 3-6: Campus=Col3, Grade=Col4, Level=Col5, StudentGroup=Col9, GuideEmail=Col12 (was Col2/3/4/8/11). Data range A:M → A:N.
- **SORT_OPTS for Sheets 3-6** — added "Mismatch Summary" as 2nd sort option
- **Code.gs appendRow** — now writes `[date, mismatchSummary].concat(data)` instead of `[date].concat(data)` for both accept and reject paths

## [v2.1.1] — 2026-04-16

### Improved
- **API call batching — pre-write phase** — Reduced ~28 sequential API calls down to ~5:
  - `_ensure_tab_exists` (13 individual calls) → `_ensure_all_tabs` (1 read + 1 batch create)
  - `unmergeCells` (6 individual calls, one per visible sheet) → single batched `batchUpdate` with all 6
  - `values().clear()` (9 individual calls) → single `values().batchClear()` call
- **Banding coverage extended** — Alternating row colors now cover 200 data rows (was 14 for cumulative sheets). Sheets 3-6 previously had `end_row = max(6 + 0 + 5, 20) = 20` because `num_data_rows=0`; now floors at 206.

### Changed
- **HUMAN_INSTRUCTIONS rewritten for v2.1.0** — Updated to document accept/reject workflow, 6 sheets, mismatch types, and new troubleshooting entries. Previously described v1.0 single-checkbox 3-sheet workflow.
- **AI_INSTRUCTIONS updated** — Added batching design decision (#6), slow pre-writes bug/fix entry, and 2 new known limitations (banding 200-row cap, Reason for Rejection column behavior).

## [v2.1.0] — 2026-04-15

### Added
- **Accept/Reject checkboxes on Sheet 1** — Column A ("Accept Changes", light green #D4EDDA) and Column B ("Reject Changes", light red #FEE2E2) replace the single checkbox column. Mutual exclusion: checking one unchecks the other.
- **"Rejected Changes" visible sheet** (Sheet 6) — Same full layout as other approval sheets (title, caption, filters, sort, QUERY from `_RejectedData`), with an extra "Reason for Rejection" column (blank for manual entry).
- **`_RejectedData` hidden tab** — Cumulative storage for all rejected rows (never cleared by Python). Apps Script appends rejected rows here regardless of mismatch type.
- **`unmergeCells` on re-run** — All visible sheets are unmerged before applying new formatting, preventing `mergeCells` errors when column layout changes between versions.

### Changed
- **Sheet 1 column layout** — NC1 changed from 14 to 15 columns (accept + reject + 12 fields + mismatch summary). QUERY formula output starts in C7 (was B7). Filter dropdown cells shifted by 1 column (C5/E5/G5/I5/K5, was B5/D5/F5/H5/J5). Sort By in M5 (was L5).
- **`_Lists` tab expanded to 11 columns** — 5 filter values (A-E) + 6 sort options (F-K), up from 10 columns.
- **Apps Script reads data from C:N** (12 data cols, was B:M) and Mismatch Summary from col O (was col N).
- **Grey-out range** on accepted/rejected rows: 15 columns (was 14).

## [v2.0.0] — 2026-04-15

### Added
- **Roster Addition mismatch type** — Students enrolled in MAP roster whose `student_id` is not found in SIS are now flagged as "Roster Addition" (previously "NOT IN SIS"). Mismatch Summary cell highlighted light green (#D4EDDA).
- **Unenrolling mismatch type** — Students with Notes != "Enrolled" in MAP but `admissionstatus` = "Enrolled" in SIS are flagged as "Unenrolling". Mismatch Summary cell highlighted light yellow (#FFFDE7).
- **"Roster Additions" visible sheet** (Sheet 4) — Same full layout as Automated Correction List (title, caption, filters, sort, QUERY from `_AdditionsData`).
- **"Roster Unenrollments" visible sheet** (Sheet 5) — Same layout, reads from `_UnenrollData`.
- **Apps Script routing by mismatch type** — Checkbox approvals on Sheet 1 now route to `_AdditionsData`, `_UnenrollData`, or `_ApprovedData` based on the Mismatch Summary column value.
- **Conditional formatting on Mismatch Summary column** — Three rules in priority order: "Roster Addition" → green, "Unenrolling" → light yellow, NOT_BLANK → yellow (field mismatches). Replaces previous static red coloring.
- **`_AdditionsData` and `_UnenrollData` hidden tabs** — Cumulative storage for approved roster additions and unenrollments (never cleared by Python).
- **`read_map_roster()` now returns enrolled AND non-enrolled students** — Two separate dicts enable unenrolling detection without changing the enrolled comparison flow.

### Changed
- **Field mismatch color** — Mismatch Summary data cells changed from light red (#FEE2E2) to yellow (#FFF3CD).
- **`_Lists` tab expanded to 10 columns** — 5 filter values (A-E) + 5 sort options (F-J), up from 8 columns (3 sort options).
- **`compare_students()` accepts 3 dicts** — `(map_enrolled, map_non_enrolled, sis_students)` instead of `(map_students, sis_students)`.

## [v1.3.0] — 2026-04-14

### Fixed (Critical)
- **Single-quote injection in QUERY formulas** — Student or campus names containing apostrophes (e.g., O'Brien, St. Mary's) broke all three SORT(QUERY()) formulas. All filter cell references now wrapped with `SUBSTITUTE(cell,"'","''")` to escape single quotes in QUERY string literals.
- **Sheet 3 QUERY column references off by one** — `_ApprovedData` has Date in Col1, shifting all data columns. Grade was Col4 (should be Col3), Level was Col5 (should be Col4), Student Group was Col9 (should be Col8), Guide Email was Col12 (should be Col11). All corrected.

### Fixed (Medium)
- **Duplicate student ID warnings** — `read_map_roster()` now logs a warning when the same `student_id` appears in multiple campus sheets (6 duplicates found in Hardeeville). Last-write-wins behavior preserved but now visible.
- **Dead code removed** — `COMPARE_FIELDS` was defined in `config.py` and imported but never used in comparison logic. Removed from both files.

### Improved
- **API write batching** — Visible sheet values (titles, labels, headers, formulas) now written in 2 batched API calls via `values().batchUpdate()` instead of 15+ individual `values().update()` calls. Runtime reduced from ~35s to ~16s.
- **External Student ID header detection expanded** — `MAP_HEADER_MAP["ext_student_id"]` now matches "suns number", "external student id", "suns #", and "external id" (was only "suns number").

## [v1.2.2] — 2026-04-14

### Fixed
- **Date Approved column** (Sheet 3) now formatted as `yyyy-MM-dd HH:mm:ss`. QUERY formula strips number formatting from the hidden `_ApprovedData` tab, so a `repeatCell` with `numberFormat` is applied to column A rows 7+ on the visible sheet.

## [v1.2.1] — 2026-04-14

### Changed
- **Dropdown row (row 5)** background changed from dark navy (`1E3A5F`) to lighter blue (`2D4A7A`) so dropdowns visually stand out against the dark filter label row above.
- **Mismatch Summary column** (Sheet 1 only): header cell is now dark red (`7F1D1D`), data cells are light red (`FEE2E2`) to highlight the correction reason at a glance.

## [v1.2.0] — 2026-04-13

### Added
- **Sort By dropdown** on all 3 sheets — QUERY formulas wrapped with `SORT(MATCH())` matching the Student Performance Dashboard pattern. Sort options match output column order.
- **Filter dropdowns + QUERY on Sheet 3** (Automated Correction List) — now has the same Campus/Grade/Level/Student Group/Guide Email filters as Sheets 1 and 2, plus Sort By with "Date Approved" default (descending).
- Hidden `_ApprovedData` tab — Apps Script now appends checked corrections here; visible Sheet 3 reads from it via SORT(QUERY()) formula.
- Sort options stored in `_Lists` columns F-H (one column per sheet).

### Changed
- **Subtitle font increased to size 12** (was 10) — caption row now more readable.
- **User guide rewritten** — restructured with PART 1-6 sections, H2/H3 headings, bold labels ("What it shows:", "What you do:", "Tip:", etc.), matching the Claude Code Setup Guide format.
- Apps Script now appends to hidden `_ApprovedData` instead of directly to the visible "Automated Correction List" sheet.

## [v1.1.0] — 2026-04-13

### Changed
- **Dropdown filtering now works** — switched from Apps Script `hideRows()` (which silently failed) to QUERY formulas that auto-recalculate when dropdown cells change. Same pattern as the Student Performance Dashboard.
- Raw data written to hidden `_CorrData` and `_SISData` tabs; visible sheets use QUERY formulas referencing dropdown cells in row 5.
- Sheet layout updated: Row 1=title, Row 2=caption with clickable User Guide link, Row 3=spacer, Rows 4-5=filter dropdowns, Row 6=headers, Row 7+=QUERY output.
- Apps Script updated: checkbox handling targets row 7+, clears stale checkboxes when filter dropdown changes (row 5).

### Added
- Dashboard-style title and caption rows (matching Student Performance Dashboard visual)
- Clickable "User Guide" hyperlink in caption using `textFormatRuns`
- `write_user_guide.py` — auto-writes formatted user guide to Google Docs via Docs API
- Google Doc user guide: https://docs.google.com/document/d/1O1WEAHSttdNVRUa_CoQ3T6w4QEFPyLz5FDdM2IMHEu4

### Fixed
- `textFormatRun.startIndex` error when "User Guide" is at end of caption string

## [v1.0.0] — 2026-04-13

### Added
- Initial release
- MAP Roster reader with header auto-detection (handles schema differences across campuses)
- 9 campus sheets: Ridgeland SAE, Ridgeland Elem, Hardeeville Jr/Sr, Hardeeville Elem, Allendale Aspire, Allendale Fairfax MS, Allendale Fairfax ES, Metro Schools, Reading CCSD
- BigQuery `alpha_roster` query (deduped from `alpha_student` via ROW_NUMBER)
- Comparison engine: 10 fields, normalized (strip/lowercase/collapse whitespace)
- Guide name comparison: combines MAP first+last vs SIS single advisor field
- Google Sheets output: 3 tabs with navy header formatting, alternating row colors, checkboxes
- Apps Script `onEdit` trigger for checkbox approval workflow (copies to Sheet 3 with date)
- `alpha_roster` CTAS export added to `Refresh-Data.ps1` as step 11b
- Corrupt header fallback: if col A header isn't "Student ID", assumes col A = student_id
- `run_export.ps1` for standalone BQ table creation
- Full documentation: README, AI_INSTRUCTIONS, HUMAN_INSTRUCTIONS
