# Changelog

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
