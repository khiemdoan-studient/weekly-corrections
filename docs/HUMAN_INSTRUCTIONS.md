# Human Instructions — Weekly Corrections

## How to Run

### Prerequisites
1. Python 3.12+ installed
2. Service account key at `keys/sa-main.json`
3. MAP roster shared with `service-account@reading-dashboard-482106.iam.gserviceaccount.com` (Viewer)
4. Output sheet shared with the same service account (Editor)
5. `alpha_roster` table exists in BigQuery (created by the dashboard pipeline)

### Install
```bash
cd weekly-corrections
pip install -r requirements.txt
```

### Run
```bash
python generate_corrections.py
```

Expected output shows enrolled counts per campus, mismatch counts by type, and a link to the output sheet.

## How to Review Corrections

1. Open the [Automated Weekly Corrections](https://docs.google.com/spreadsheets/d/12dqu58KKdsZN9nLre9Fntkk7vSILu3KfcW4WDvo5-Ls) spreadsheet
2. **Sheet 1 ("Corrected Roster Info")** shows MAP roster data for mismatched students, with color-coded Mismatch Summary:
   - **Green** — Roster Addition (student in MAP but not in SIS)
   - **Yellow** — Field mismatch (student in both, specific fields differ)
   - **Light yellow** — Unenrolling (student no longer enrolled in MAP but still in SIS)
3. **Sheet 2 ("Current Roster Info in SIS")** shows the same students' SIS data
4. Compare side by side — the Mismatch Summary column (last column in Sheet 1) tells you exactly which fields differ
5. **Accept or Reject** each correction:
   - **Column A (Accept Changes, green)** — Check if the MAP data should replace the SIS data
   - **Column B (Reject Changes, red)** — Check if the correction should NOT be applied
   - Checking one automatically unchecks the other
6. Accepted students are routed to the appropriate approval sheet:
   - Field mismatches → **Sheet 3 ("Automated Correction List")**
   - Roster additions → **Sheet 4 ("Roster Additions")**
   - Unenrollments → **Sheet 5 ("Roster Unenrollments")**
7. Rejected students appear in **Sheet 6 ("Rejected Changes")** — optionally add a reason in the last column
8. Every Friday, the data team reviews all approval sheets and submits corrections

## How to Install Apps Script

The accept/reject workflow requires a one-time Apps Script setup:

1. Open the Automated Weekly Corrections spreadsheet
2. Go to **Extensions > Apps Script**
3. Delete any existing code in the editor
4. Copy the contents of `apps_script/Code.gs` and paste it in
5. Press **Ctrl+S** to save
6. Close the Apps Script editor — no deployment needed, `onEdit` triggers run automatically

**Important:** After any code update (e.g., new version with accept/reject columns), repeat steps 2-5 to paste the latest `Code.gs`.

## Spreadsheet Sheets

| # | Sheet Name | Purpose |
|---|-----------|---------|
| 1 | Corrected Roster Info | MAP roster data + accept/reject checkboxes for mismatched students |
| 2 | Current Roster Info in SIS | SIS data for the same students (read-only comparison) |
| 3 | Automated Correction List | Running history of approved field-mismatch corrections |
| 4 | Roster Additions | Running history of approved new student enrollments |
| 5 | Roster Unenrollments | Running history of approved student unenrollments |
| 6 | Rejected Changes | Running history of rejected corrections with reason column |

## Dropdown Filters

Each sheet has 5 filter dropdowns (Campus, Grade, Level, Student Group, Guide Email) and a Sort By dropdown. Set all filters before checking any boxes — changing a filter resets all checkboxes.

## Adding a New Campus

1. Add the sheet tab name (with "(Dash)" suffix) to the `CAMPUS_SHEETS` list in `config.py`
2. If the new campus has non-standard column headers, add the header text to the appropriate field in `MAP_HEADER_MAP` in `config.py`
3. Re-run the script

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "The caller does not have permission" | Share the spreadsheet with the service account email |
| Campus shows 0 enrolled students | Check the Notes column header in the MAP roster — it must match one of the entries in `MAP_HEADER_MAP["notes"]` |
| "Table alpha_roster was not found" | Run `run_export.ps1` or the full dashboard pipeline to create the BQ table |
| Checkbox doesn't route to approval sheet | Install/re-install the Apps Script (see above) |
| "addBanding" error | Auto-handled on re-runs; the script clears existing banding first |
| "mergeCells" error | Auto-handled on re-runs; the script unmerges all cells before formatting |
| Checkboxes disappeared after changing a filter | Expected behavior — checkboxes reset when filters change because the visible rows shift |
| I don't see my campus in the dropdown | Your campus may not have any mismatched students this week |
