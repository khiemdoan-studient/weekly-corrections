# Weekly Corrections

Automated comparison of MAP roster (source of truth) against SIS pipeline data (BigQuery). Identifies mismatched student enrollment information and routes them through an accept/reject approval workflow managed by Implementation Managers.

## Architecture

```
MAP Roster (Google Sheet)    SIS source-of-truth (per campus)
  11 campus tabs:               9 Dash → BigQuery alpha_roster (~8,700 students)
  - 9 Dash campuses             2 Timeback → OneRoster API live (~76 students)
  - Vita (Timeback)                       (Vita + ScienceSIS, v2.7.0+)
  - ScienceSIS (Timeback)
       |                                |                  |
       v                                v                  v
   read_map_roster()         read_sis_data()    query_timeback_enrolled()
       |                                |                  |
       |                                +-- merge (Timeback wins on overlap)
       |                                          |
       +----- compare_students() by student_id ---+
                          |
                    4 mismatch types:
                    • Roster Addition  (in MAP, not in SIS)
                    • Field mismatch   (in both, fields differ)
                    • Unenrolling      (not in MAP, enrolled in SIS)
                    • IM-flagged Unenroll (Unenroll=TRUE on ISR + still SIS-enrolled)
                          |
                          v
            "Automated Weekly Corrections" spreadsheet
              Sheet 1: MAP data + Accept/Reject checkboxes
              Sheet 2: SIS data (same students)
              Sheet 3: Automated Correction List (field mismatches)
              Sheet 4: Roster Additions
              Sheet 5: Roster Unenrollments
              Sheet 6: Rejected Changes (with Reason column)
```

## Quick Start

### Prerequisites
- Python 3.12+
- Service account key at `keys/sa-main.json`
- MAP roster shared with SA as Viewer
- Output sheet shared with SA as Editor
- `Code.js` deployed to the corrections spreadsheet's Apps Script project. Use `npm run deploy` locally, or rely on the GHA auto-deploy workflow (see "Apps Script Auto-Deploy" below).

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
| `config.py` | Constants: sheet IDs, BQ config, header mappings, campus list, tab names |
| `queries.py` | BigQuery query for alpha_roster table (includes admissionstatus) |
| `sheets_writer.py` | Google Sheets API: tabs, QUERY formulas, format, migration, backfill |
| `Code.js` | Apps Script onEdit: accept/reject routing by mismatch type. Auto-deployed via clasp (see "Apps Script Auto-Deploy"). |
| `appsscript.json` | Apps Script manifest (TZ + V8 runtime), pushed alongside Code.js. |
| `write_user_guide.py` | Generates the User Guide Google Doc (linked from every sheet) |
| `run_export.ps1` | One-time script to create alpha_roster BQ table |
| `alpha_roster_ctas.sql` | Athena CTAS query for alpha_roster export (deduped) |

## Configuration

| Constant | Value |
|----------|-------|
| GCP Project | `studient-flat-exports-doan` |
| BQ Dataset | `studient_analytics` |
| BQ Table | `alpha_roster` |
| MAP Roster Sheet | `1scEay0a8OR6vU3uJuxbHKWCEx_RVgSsRXF9naJh3XYw` |
| Output Sheet | `12dqu58KKdsZN9nLre9Fntkk7vSILu3KfcW4WDvo5-Ls` |
| User Guide Doc | `1O1WEAHSttdNVRUa_CoQ3T6w4QEFPyLz5FDdM2IMHEu4` |
| Service Account | `service-account@reading-dashboard-482106.iam.gserviceaccount.com` |

## Output Sheet Structure

### Visible Sheets (6)
- **Sheet 1 "Corrected Roster Info"** — MAP roster data for mismatched students. Col A = Accept Changes (green checkbox), Col B = Reject Changes (red checkbox), Col O = Mismatch Summary (color-coded: green/yellow/light yellow).
- **Sheet 2 "Current Roster Info in SIS"** — SIS data for the same students, aligned row-by-row.
- **Sheet 3 "Automated Correction List"** — Cumulative list of approved field-mismatch corrections with date stamp and mismatch type.
- **Sheet 4 "Roster Additions"** — Cumulative list of approved new enrollments.
- **Sheet 5 "Roster Unenrollments"** — Cumulative list of approved unenrollments.
- **Sheet 6 "Rejected Changes"** — Cumulative list of rejected corrections with Reason for Rejection column.

### Hidden Tabs (7)
- `_CorrData`, `_SISData`, `_Lists` — Source data for QUERY formulas (rebuilt each run)
- `_ApprovedData`, `_AdditionsData`, `_UnenrollData`, `_RejectedData` — Cumulative history (14-col format: Date, MismatchSummary, 12 fields)

## Workflow

1. Script runs (manually or scheduled), comparing MAP vs SIS
2. Mismatches written to Sheet 1 with Accept/Reject checkboxes
3. IMs filter by campus, review, and check Accept or Reject
4. Apps Script routes checked rows to appropriate hidden tab with mismatch type
5. Data team processes approval sheets every Friday

See `docs/AI_INSTRUCTIONS.md` for architecture details and `docs/HUMAN_INSTRUCTIONS.md` for user-facing workflow.

## Apps Script Auto-Deploy

`Code.js` (the Apps Script for the corrections spreadsheet) is auto-deployed
via [clasp](https://github.com/google/clasp). No manual copy-paste step.

### Per-edit workflow (local)

```bash
npm run deploy
```

Runs `node --check Code.js` (syntax) then `clasp push` to deploy to the
linked Apps Script project. Reload the corrections spreadsheet to use the
new code.

### One-time setup (for GHA auto-deploy on git push)

GitHub Actions auto-deploys on every push to main that touches `Code.js`,
`appsscript.json`, or `.claspignore`. Requires two secrets:

1. **CLASPRC_JSON**: contents of your local `~/.clasprc.json` (or `%USERPROFILE%\.clasprc.json` on Windows). Generated when you run `clasp login`.
2. **CLASP_SCRIPT_ID**: the Apps Script project ID for the corrections spreadsheet
   (currently `16_ypoWiIFRpIZzUEpwJGLP8DvexGoCDAiXaZGKoIhRQFa38H8vcS436_`).

Add both via GitHub repo → Settings → Secrets and variables → Actions → New repository secret.

Until secrets are configured, the workflow runs but skips deploy with a warning. You can still use `npm run deploy` locally.

### Files involved

| File | Purpose |
|---|---|
| `Code.js` | Apps Script source (deployed) |
| `appsscript.json` | Apps Script manifest (TZ, runtime) |
| `.claspignore` | Whitelist: only Code.js + appsscript.json get pushed |
| `.clasp.json` | Local-only (gitignored) — contains scriptId of bound Apps Script |
| `package.json` | npm scripts: deploy, push, pull, open |
| `.github/workflows/deploy-apps-script.yml` | GHA auto-deploy workflow |

## Pipeline Health & Monitoring

### Daily / hourly health
Both workflows run with retry hardening (v2.5.2) and only open a tracking
Issue (label `pipeline-failure`) on 3+ consecutive failures (v2.5.3 smart-
notify). Single transient API blips are absorbed silently.

### Weekly health summary
Every Monday at 12:00 UTC, `weekly-health-report.yml` opens a tracking
Issue (label `health-report`) summarizing the last 30 days: success rate,
failure count, median duration, etc.

### On-demand health check
```bash
python health_report.py --days 30
# or write to a file:
python health_report.py --days 30 --output health.md
```
