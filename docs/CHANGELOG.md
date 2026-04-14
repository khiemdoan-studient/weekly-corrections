# Changelog

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
