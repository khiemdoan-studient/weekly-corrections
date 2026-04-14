# Weekly Corrections

Automated comparison of MAP roster (source of truth) against SIS pipeline data (BigQuery). Identifies mismatched student enrollment information and outputs corrections for manager approval.

## Architecture

```
MAP Roster (Google Sheet)          BigQuery (alpha_roster)
  9 campus sheets                    8,229 deduped students
        |                                    |
        v                                    v
    Python: Sheets API read         Python: BQ client query
        |                                    |
        +----------- Compare by Student_ID --+
                           |
                     125 mismatches
                           |
                           v
            "Automated Weekly Corrections" spreadsheet
              Sheet 1: MAP data + checkboxes
              Sheet 2: SIS data (same students)
              Sheet 3: Approved corrections (Apps Script)
```

## Quick Start

### Prerequisites
- Python 3.12+
- Service account key at `keys/sa-main.json`
- MAP roster shared with SA as Viewer
- Output sheet shared with SA as Editor

### Install & Run
```bash
pip install -r requirements.txt
python generate_corrections.py
```

### Pipeline Integration
The `alpha_roster` BQ table must exist. It is created by step 11b of `Refresh-Data.ps1` in the Studient Excel Automation project. For first-time setup, run `run_export.ps1`.

## File Reference

| File | Purpose |
|------|---------|
| `generate_corrections.py` | Main orchestrator: auth, read MAP, query BQ, compare, write |
| `config.py` | Constants: sheet IDs, BQ config, header mappings, campus list |
| `queries.py` | BigQuery query for alpha_roster table |
| `sheets_writer.py` | Google Sheets API: clear, write, format, checkboxes |
| `apps_script/Code.gs` | Apps Script onEdit trigger for checkbox approval |
| `run_export.ps1` | One-time script to create alpha_roster BQ table |
| `alpha_roster_ctas.sql` | Athena CTAS query for alpha_roster export |

## Configuration

| Constant | Value |
|----------|-------|
| GCP Project | `studient-flat-exports-doan` |
| BQ Dataset | `studient_analytics` |
| BQ Table | `alpha_roster` |
| MAP Roster Sheet | `1scEay0a8OR6vU3uJuxbHKWCEx_RVgSsRXF9naJh3XYw` |
| Output Sheet | `12dqu58KKdsZN9nLre9Fntkk7vSILu3KfcW4WDvo5-Ls` |
| Service Account | `service-account@reading-dashboard-482106.iam.gserviceaccount.com` |

## Output Sheet

- **Sheet 1 "Corrected Roster Info"**: MAP roster data for mismatched students. Column A = checkbox for manager approval. Column N = mismatch summary.
- **Sheet 2 "Current Roster Info in SIS"**: SIS data for the same students, aligned row-by-row.
- **Sheet 3 "Automated Correction List"**: Cumulative list of approved corrections with date stamps. Managed by Apps Script.
