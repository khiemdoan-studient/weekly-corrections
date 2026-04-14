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

Expected output shows enrolled counts per campus, mismatch count, and a link to the output sheet.

## How to Review Corrections

1. Open the [Automated Weekly Corrections](https://docs.google.com/spreadsheets/d/12dqu58KKdsZN9nLre9Fntkk7vSILu3KfcW4WDvo5-Ls) spreadsheet
2. **Sheet 1 ("Corrected Roster Info")** shows MAP roster data for mismatched students
3. **Sheet 2 ("Current Roster Info in SIS")** shows the same students' SIS data
4. Compare side by side — column N in Sheet 1 lists which fields differ
5. Check the checkbox in column A for students whose MAP data should replace the SIS data
6. Checked students automatically appear in **Sheet 3 ("Automated Correction List")** with a date stamp
7. Every Friday, copy Sheet 3 data to the support team for enrollment corrections

## How to Install Apps Script

The checkbox approval workflow requires a one-time Apps Script setup:

1. Open the Automated Weekly Corrections spreadsheet
2. Go to **Extensions > Apps Script**
3. Delete any existing code in the editor
4. Copy the contents of `apps_script/Code.gs` and paste it in
5. Press **Ctrl+S** to save
6. Close the Apps Script editor — no deployment needed, `onEdit` triggers run automatically

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
| Checkbox doesn't copy to Sheet 3 | Install the Apps Script (see above) |
| "addBanding" error | This is auto-handled on re-runs; the script clears existing banding first |
