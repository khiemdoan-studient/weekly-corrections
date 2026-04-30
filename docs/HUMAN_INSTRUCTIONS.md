# Human Instructions — Weekly Corrections

## How to Run

### Prerequisites
1. Python 3.12+ installed
2. Service account key at `keys/sa-main.json`
3. MAP roster shared with `service-account@reading-dashboard-482106.iam.gserviceaccount.com` (Viewer)
4. Output sheet shared with the same service account (Editor)
5. `alpha_roster` table exists in BigQuery (created by the dashboard pipeline)
6. Run `python setup_unenroll_columns.py` once after cloning to provision Unenroll columns on all 9 ISRs + CMR (already done as part of v2.3.0 deployment; only needed if setting up a new environment)

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

### Automation

The pipeline now runs **automatically every hour** via GitHub Actions — you don't need to ask anyone to kick it off. You can see the run history (and any failures) here: https://github.com/khiemdoan-studient/weekly-corrections/actions

Manual runs via `python generate_corrections.py` still work if you ever need one off-cycle, but they aren't necessary for the normal weekly review.

Each Sheets/Drive API call has up to 5 in-script retries with exponential backoff (about 5 minutes of coverage), and GHA itself retries the whole Python step once on failure. You should rarely see a real failure unless the API is down for a long stretch.

## Pipeline alerts (v2.5.3)

### What changed
The pipeline now only escalates persistent failures, not transient blips.
You should rarely get a failure email anymore.

### How to set up notifications
1. **Mute the default GitHub Actions failure emails** —
   - Visit https://github.com/settings/notifications
   - Under "Actions" → uncheck "Send notifications for failed workflows
     for workflows you triggered" (or similar)
   - Without this step, you'll keep getting an email for every transient
     blip alongside the new smart-notify issues.

2. **Subscribe to the new tracking issues**
   - GitHub Issues with label `pipeline-failure` open ONLY when 3+
     consecutive runs fail. This is your real alert.
   - GitHub Issues with label `health-report` open every Monday with a
     30-day summary. This is your trend visibility.
   - Watch the repo at https://github.com/khiemdoan-studient/weekly-corrections
     and set notification preferences to "Issues" only (or use a label-
     filtered subscription).

### What you'll see
- **Single failure**: nothing — the workflow retries internally and the
  next hourly run absorbs it. You can verify by checking the run log.
- **3+ consecutive failures**: a tracking issue opens on the repo,
  GitHub emails you. Investigate the run log.
- **Recovery after persistent failure**: the tracking issue auto-closes
  with a comment "Pipeline recovered. Auto-closing".
- **Every Monday at 8 AM ET**: a health-report issue with a 30-day
  summary appears.

## Row Hide Timing (cheat sheet)

| Action | Latency |
|--------|---------|
| Check Accept/Reject | Instant — Apps Script routes the row to the cumulative tab, greys cols C–O |
| Row disappears from Sheet 1 | Up to 1 hour (next hourly pipeline run) |
| Row stays hidden | 7 days |
| Row reappears on Sheet 1 | Only if mismatch still exists in MAP vs SIS after 7 days (data team hasn't processed it) |

## How to Review Corrections

1. Open the [Automated Weekly Corrections](https://docs.google.com/spreadsheets/d/12dqu58KKdsZN9nLre9Fntkk7vSILu3KfcW4WDvo5-Ls) spreadsheet
2. **Sheet 1 ("Corrected Roster Info")** shows MAP roster data for mismatched students, with color-coded Mismatch Summary:
   - **Green** — Roster Addition (student in MAP but not in SIS)
   - **Yellow** — Field mismatch (student in both, specific fields differ)
   - **Light red** — Unenrolling (student no longer enrolled in MAP but still in SIS)
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

   **What happens after you check Accept or Reject:**
   - The row gets greyed out in cols C–O (data area). Cols A (Accept) and B (Reject) stay green/red — those are the permanent column colors.
   - Within the hour, the next automatic pipeline run will HIDE that row from "Corrected Roster Info" entirely.
   - The row stays hidden for 7 days. After that, if the correction hasn't been processed by the data team (i.e., SIS still has the wrong value), the student reappears on Sheet 1 — this is a signal that something's overdue.
   - If the correction WAS processed within 7 days, the student's MAP and SIS data now match, so they never reappear.
   - The row in Sheet 3 / 4 / 5 / 6 (approval/rejection sheets) is permanent regardless.
7. Rejected students appear in **Sheet 6 ("Rejected Changes")** — optionally add a reason in the last column
8. Every Friday, the data team reviews all approval sheets and submits corrections

## How to Mark a Student for Unenrollment

If you're an Implementation Manager (IM) and a student needs to be unenrolled, you can flag them directly from your campus's Individual Student Roster (ISR) — no need to email the data team.

1. Open your campus's ISR spreadsheet (links below) and go to the **Student Roster** tab
2. Find the row for the student you want to unenroll
3. Check the **Unenroll** checkbox — the column position depends on your campus:
   - **Column X** — Reading CCSD
   - **Column Y** — Metro
   - **Column Z** — AFES, AFMS
   - **Column AB** — AASP, JHES, JHMS, JRES, JRHS
4. Within about a minute (once IMPORTRANGE refreshes), your checkbox will propagate through to the ISR's **MAP Roster** tab and then into the **CMR Unenroll** column
5. The next time the weekly-corrections pipeline runs (`python generate_corrections.py`), that student will show up in the **Roster Unenrollments** sheet
6. The data team processes Roster Unenrollments every Friday along with all other approved corrections

**How this interacts with other corrections (option-C precedence):** If you check Unenroll for a student AND our SIS still has them enrolled, that student shows up as "Unenrolling" in the correction list — this takes priority over any other field mismatches for that student. In other words, if you're saying "unenroll them," we don't bother you about a mismatched grade or guide email — we just unenroll.

**Important:** The checkbox stays checked permanently as a historical record of who was unenrolled and when. **Do NOT uncheck it after the student is processed.**

### Campus ISR Quick Links

| Campus | Student Roster (ISR) |
|--------|----------------------|
| Reading CCSD | https://docs.google.com/spreadsheets/d/1b28bgPy9mysb31Op01DPL6IS5jhMZy51VKVSsS0feII |
| Metro | https://docs.google.com/spreadsheets/d/1Eri0B_WMmjJxPs6SYszK2F1jJt08rjPFRwgrpuKdakU |
| AFES | https://docs.google.com/spreadsheets/d/1zhWCgoJB9WXA9sDxnHj0uZbHQaWTiMFk3bLP61rUKWo |
| AFMS | https://docs.google.com/spreadsheets/d/1r6o0j8ENz01gt9L5ygJLBZAtwCD-L9H2SrutgZyTfQc |
| AASP | https://docs.google.com/spreadsheets/d/10H5y0Z3_QAH9wYH5V80yLSLuqWPKpDXf5k7wDMz7hww |
| JHES | https://docs.google.com/spreadsheets/d/1waahGamoiMb5DkLF1_IlO5kEhpcc9g7NZr3WIeiAfFw |
| JHMS | https://docs.google.com/spreadsheets/d/1g8KUreiGlBd2NM5huZjSSDA30YdDD8kL0geUJb0Ajww |
| JRES | https://docs.google.com/spreadsheets/d/1IwGsdtThjQJmcfbh_eR5ZrFKQiCj9FL5GR02geRofWQ |
| JRHS | https://docs.google.com/spreadsheets/d/1AT4jEZPbaYdFJUI8OTAVIgHOs4cjCY6zh96I7uMvZZM |

## How to Install Apps Script

The accept/reject workflow requires a one-time Apps Script setup:

1. Open the Automated Weekly Corrections spreadsheet
2. Go to **Extensions > Apps Script**
3. Delete any existing code in the editor
4. Copy the contents of `apps_script/Code.gs` and paste it in
5. Press **Ctrl+S** to save
6. Close the Apps Script editor — no deployment needed, `onEdit` triggers run automatically

**Important:** After any code update (e.g., new version with accept/reject columns), repeat steps 2-5 to paste the latest `Code.gs`.

**v2.4.2 update (race-condition fix)**: If you're upgrading from an earlier
version, you need to re-paste `apps_script/Code.gs` once more. Older versions
had a rare but real race where clicking multiple Accept/Reject checkboxes
within a second could leave some rows with an inconsistent date format
(like `4/23/2026 1:37:44` instead of `2026-04-23 01:37:44`), breaking the
"Date Approved" chronological sort. The fix pre-formats the timestamp before
writing the row, eliminating the race.

## Spreadsheet Sheets

| # | Sheet Name | Purpose |
|---|-----------|---------|
| 1 | Corrected Roster Info | MAP roster data + accept/reject checkboxes for mismatched students |
| 2 | Current Roster Info in SIS | SIS data for the same students (read-only comparison) |
| 3 | Automated Correction List | Running history of approved field-mismatch corrections |
| 4 | Roster Additions | Running history of approved new student enrollments |
| 5 | Roster Unenrollments | Running history of approved student unenrollments |
| 6 | Rejected Changes | Running history of rejected corrections with reason column |
| 7 | Unenroll Queue (Live) | Real-time view of IM-flagged Unenroll students across all 9 campuses. Updates within ~1 minute of checking a box. |

## How the Live Queue Updates vs the Full Pipeline

There are two different things that update when you flag a student for Unenroll, and they happen on different timelines:

- **Live Queue** (the "Unenroll Queue (Live)" tab) shows IM-flagged students IMMEDIATELY, but doesn't check the SIS — it's a visibility tool so you and other IMs can see who's been flagged at a glance.
- **Full pipeline** (which updates the "Roster Unenrollments" sheet and others) runs AUTOMATICALLY every hour via GitHub Actions.

So if you flag a student for Unenroll:
- **Within ~1 min:** they appear on "Unenroll Queue (Live)"
- **Within 1 hour:** the full pipeline picks them up, compares to SIS, and adds them to "Roster Unenrollments" for the data team

You no longer need to ask someone to run `python generate_corrections.py` — it runs automatically.

## Dropdown Filters

Each sheet has 5 filter dropdowns (Campus, Grade, Level, Student Group, Guide Email) and a Sort By dropdown. Set all filters before checking any boxes — changing a filter resets all checkboxes.

## Adding a New Campus

1. Add the sheet tab name (with "(Dash)" suffix) to the `CAMPUS_SHEETS` list in `config.py`
2. If the new campus has non-standard column headers, add the header text to the appropriate field in `MAP_HEADER_MAP` in `config.py`
3. Re-run the script

## How to Send Weekly Corrections to Support

Every Monday at 07:00 ET, a new Google Sheet appears in the Shared Drive "Weekly Corrections Archive" with the week's corrections packaged up for the data team / support contact.

**If no IMs accepted any corrections during the past week, no Monday file is
generated.** The Shared Drive "Weekly Corrections Archive" stays clean. This
is intentional — there's nothing to send to support that week.

To check what happened on a Monday morning, look at GitHub Actions:
https://github.com/khiemdoan-studient/weekly-corrections/actions/workflows/weekly-snapshot.yml
The job log will say either "No corrections to send this week. File not
created." or it will show row counts and a URL to the new file.

- **Shared Drive:** https://drive.google.com/drive/folders/0AFQGIqcKjsyFUk9PVA
- **File name:** `M/D Corrections` based on the Monday of the current week (e.g. `4/20 Corrections`)
- **Contents:** 3 tabs, each with the 14-col layout matching the approval sheets:
  - `Correction List` — field mismatches the data team needs to update
  - `Roster Additions` — new enrollments to add in SIS
  - `Roster Unenrollments` — unenrollments to process in SIS
- Empty tabs are hidden automatically — you only see what's actionable.

**How to send it:**
1. Open the sheet from the Shared Drive
2. Click Share, add your support contact as Viewer or Editor
3. Send them the link via email

The same file URL stays valid all week — no need to resend.

**Re-running mid-week:** If more corrections come in through the week, you can trigger the GitHub Actions `Weekly snapshot (Monday)` workflow manually from the Actions tab. It'll find the existing file and update it in place — the URL you already sent stays valid, support just sees the new rows on reload.

### How the Sent Week column works

Every accepted / unenrolled / added row gets stamped with the current Monday's date (e.g. `2026-04-20`) when it's included in a weekly snapshot. The next Monday's run only picks up rows with a blank Sent Week, so you never get duplicates across weeks.

If you ever need to RE-send a row in a later week (rare): manually clear col O on the `_ApprovedData` / `_AdditionsData` / `_UnenrollData` hidden tab for that row. Next Monday's snapshot will pick it back up.

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
| My Unenroll checkbox isn't showing up in the correction list | (1) Has IMPORTRANGE refreshed? It can take up to a minute. (2) Is SIS actually still showing Enrolled for that student? If SIS already matches MAP, nothing to flag. (3) Has the Python pipeline run since you checked the box? |
| Unenroll Queue (Live) shows #REF! or is empty | This is a one-time auth prompt from IMPORTRANGE. Click 'Allow access' when you see the pop-up in the sheet — the data will populate within seconds. |
| Hourly pipeline hasn't run when expected | Check https://github.com/khiemdoan-studient/weekly-corrections/actions for any failed runs. Click 'Re-run all jobs' on a failed workflow or ask Khiem. |
| GitHub Actions email says a workflow failed, but the next hourly run succeeded | Almost certainly a transient Google Sheets API hiccup. The workflow now retries 5 times in-script with exponential backoff, AND the GHA workflow itself retries the whole Python step once if it exits 1. So a single failure email usually means the API was actually hiccuping for >5 minutes. Check the run log: if you see "[retry] attempt N/5 failed (HttpError 5xx)" lines and then a later run succeeded, it self-recovered. No action needed. |
| Date column in approval sheets has mixed formats (`4/23/2026` mixed with `2026-04-23`) | Older Code.gs race condition. Run `python normalize_dates.py` once to fix historical rows, then re-paste the current `apps_script/Code.gs` to prevent future drift. |
| A student I accepted/rejected last week is back on Sheet 1 | Expected behavior — the 7-day hide window expired. It means the data team hasn't processed the correction yet. Re-check your box to re-hide, or ping Khiem. |
| My Accept (col A) / Reject (col B) columns are white/grey instead of green/red | You're probably running an older Apps Script. Re-paste the current `apps_script/Code.gs` from the repo into Extensions > Apps Script. The v2.4.3+ version only modifies cols C–O on checkbox click, so cols A/B keep their permanent green/red column colors. |
| The weekly sheet didn't generate this Monday | Check GitHub Actions. Go to https://github.com/khiemdoan-studient/weekly-corrections/actions — look for the "Weekly snapshot (Monday)" workflow. If it shows failed/skipped, click it to see logs. To re-run manually: click the Run workflow button on the right. |
| "It's Monday and there's no new file in the Shared Drive" | Either (a) no IMs accepted any corrections during the past week, so the snapshot script intentionally skipped file creation — check Actions log for "No corrections to send this week", or (b) the workflow failed — open the most recent run at https://github.com/khiemdoan-studient/weekly-corrections/actions/workflows/weekly-snapshot.yml and read the error. To force a file even when there are no unsent rows: clear col O on a row in the hidden `_ApprovedData` / `_AdditionsData` / `_UnenrollData` tab, then click "Run workflow" on the Actions page. |
| I see '4/20 Corrections' but it shows last week's data | Expected IF no new corrections were accepted this week. The snapshot includes rows stamped with THIS Monday's date plus any unsent rows. If the data team hasn't processed last week's corrections AND no new ones were accepted, the file stays showing those. Once new rows are accepted, re-trigger the workflow to refresh. |
| I manually created a sheet in the Shared Drive but the script made a separate one | The script matches exactly by filename `M/D Corrections` (e.g. `4/20 Corrections`). If you made something differently named, the script creates a fresh one alongside. Rename your file or delete one of them. |
| I want to see the pipeline's recent run history at a glance | Run `python health_report.py --days 30` from the repo root. Or look at https://github.com/khiemdoan-studient/weekly-corrections/issues?q=label%3Ahealth-report for the auto-generated weekly summaries. |
| There's an open issue with label `pipeline-failure` and I want to know which workflow is broken | The issue title contains the workflow name (e.g. "🚨 Hourly corrections pipeline persistently failing"). The body links to the most recent failed run. Click that link to see the error. |
