# AI Instructions — Weekly Corrections

## Project Overview

This tool compares student enrollment data between two sources:
1. **MAP Roster** (Google Sheet) — the source of truth for student enrollment
2. **SIS** — depends on the campus:
   - 9 Dash campuses → BigQuery `alpha_roster` table (deduped export from Athena `alpha_student`).
   - 2 Timeback campuses (Vita + ScienceSIS) → OneRoster API (`api.alpha-1edtech.ai`) live each run, via `timeback_sis.py`. Added in v2.7.0.

Students with mismatched data appear in a corrections spreadsheet for manager review.

## Critical File Map

| File | Lines | Purpose |
|------|-------|---------|
| `generate_corrections.py` | ~470 | Main orchestrator: auth, read MAP, query BQ + Timeback, compare, filter recently-handled, write. v2.7.0: `read_combined_sis_data` merges alpha_roster + Timeback OneRoster results. |
| `config.py` | ~110 | Constants, header mappings, campus list, sheet IDs. v2.7.0: adds `TIMEBACK_CAMPUSES` + `TIMEBACK_CREDS_PATH`. |
| `queries.py` | ~30 | Single BQ query function for alpha_roster |
| `timeback_sis.py` | ~210 | v2.7.0: OneRoster API bridge. OAuth2 + GET /schools/{id}/students. `query_timeback_enrolled()` returns dict shaped like alpha_roster. Self-contained — does NOT depend on sibling `timeback-data-pipeline/oneroster_client.py` (intentional duplication; document any drift). |
| `sheets_writer.py` | ~850 | Sheets API: hidden tabs, QUERY formulas, title/caption, filters, format |
| `Code.js` | ~370 | Apps Script source at repo root: onEdit accept/reject routing, clear on filter change. Auto-deployed via clasp (was `apps_script/Code.gs` pre-v2.6.0). |
| `appsscript.json` | ~7 | Apps Script manifest (TZ America/New_York, V8 runtime). Pushed alongside Code.js by clasp. |
| `.claspignore` | ~22 | clasp whitelist: ignores everything by default, only Code.js + appsscript.json get pushed. Prevents accidental push of Python tooling, docs, or `.claude/` config. |
| `package.json` | ~16 | npm deploy scripts: `npm run deploy` runs `node --check Code.js` + `clasp push`. Also exposes `pull`, `push`, `open`. |
| `.github/workflows/deploy-apps-script.yml` | ~80 | GHA auto-deploy on push to main when Code.js / appsscript.json / .claspignore changes. Fail-soft: logs warning + skips when CLASPRC_JSON / CLASP_SCRIPT_ID secrets aren't configured. |
| `write_user_guide.py` | ~215 | Google Docs API: write formatted user guide |
| `run_export.ps1` | ~70 | One-time: Athena CTAS → S3 → GCS → BQ for alpha_roster |
| `alpha_roster_ctas.sql` | ~30 | Athena SQL with dedup, handles reserved word "group" |
| `setup_unenroll_columns.py` | ~150 | One-time: provision Unenroll columns on all 9 ISRs + CMR. Idempotent — safe to re-run |
| `build_unenroll_queue.py` | ~250 | One-time: create/refresh "Unenroll Queue (Live)" tab on corrections sheet with per-campus QUERY+IMPORTRANGE formulas. Idempotent. |
| `.github/workflows/hourly-pipeline.yml` | ~40 | GitHub Actions cron: runs `generate_corrections.py` every hour at :00 UTC. Uses `GCP_SA_KEY` secret. |
| `normalize_dates.py` | ~130 | One-time: normalize date column A in cumulative tabs to canonical `yyyy-MM-dd HH:mm:ss`. Idempotent. Handles both `M/D/YYYY H:MM:SS` and ISO inputs. |
| `generate_weekly_snapshot.py` | ~1000 | Weekly orchestrator: compute Monday ET, find/create sheet in Shared Drive, filter by Sent Week col, write 3 tabs, stamp sent rows. v2.6.1: `--all-unsent` flag adds an Instructions tab pinned at index 0 for ad-hoc support packets. |
| `add_sent_week_column.py` | ~100 | Pre-flight sanity check for Sent Week col O on cumulative tabs. Reports row counts + blank/sent/malformed state. Safe to re-run. |
| `.github/workflows/weekly-snapshot.yml` | ~50 | GitHub Actions cron: `0 11 * * 1` (Mon 07:00 ET) + workflow_dispatch. Uses GCP_SA_KEY secret. |
| `retry_helper.py` | ~120 | Shared retry helper for transient Google API errors. Used by sheets_writer, generate_weekly_snapshot, generate_corrections, and all 4 helper scripts. 5 attempts, exponential backoff (1s/2s/4s/8s/16s), 25% jitter, transient-only catch (HttpError 5xx/429/408 + TimeoutError + socket.timeout + ConnectionError). |
| `health_report.py` | ~190 | Pipeline health summary script. Queries last N days of GitHub Actions runs via `gh` CLI for both workflow YAMLs. Outputs Markdown with success rate, failure streak, currently-failing count, median duration. Used by the weekly-health-report.yml cron and runnable locally for ad-hoc trend checks. |
| `.github/workflows/weekly-health-report.yml` | ~50 | Monday 12:00 UTC cron. Runs health_report.py and opens a tracking Issue with the summary, label `health-report`. |

Note: `requirements.txt` now includes `tzdata` for Windows (needed by `zoneinfo.ZoneInfo("America/New_York")`).

Note: `sheets_writer.py` and `generate_weekly_snapshot.py` previously had per-file `_retry_api`/`_retry` helpers — those are now removed in favor of the shared `retry_helper` module (v2.5.2).

## Spreadsheet Architecture

### Hidden Tabs (written by Python, never shown to users)
- `_CorrData` — Raw MAP roster data for mismatched students (13 cols, no headers)
- `_SISData` — Raw SIS data for same students (12 cols, no headers)
- `_Lists` — Unique values for filter dropdowns (cols A-E) + sort options (cols F-K, one per sheet)
- `_ApprovedData` — Cumulative approved field-mismatch corrections (14 cols: Date, MismatchSummary, 12 fields)
- `_AdditionsData` — Cumulative approved roster additions (14 cols, same layout)
- `_UnenrollData` — Cumulative approved unenrollments (14 cols, same layout)
- `_RejectedData` — Cumulative rejected changes (14 cols, same layout)

### Sheet 1: "Corrected Roster Info" (MAP data + accept/reject checkboxes)
| Row | Content |
|-----|---------|
| 1 | Title: "Corrected Roster Info" (merged, navy dark, 20pt bold white) |
| 2 | Caption with clickable "User Guide" hyperlink (merged, navy med, italic grey) |
| 3 | Spacer (5px, dark) |
| 4 | Filter labels + "SORT BY" label (merged pairs, dark) |
| 5 | Dropdown values + Sort By dropdown (merged pairs, teal bg, data validation from _Lists) |
| 6 | Column headers: Accept Changes, Reject Changes, Campus, Grade, ... Mismatch Summary (navy, bold white) |
| 7+ | SORT(QUERY()) formula in C7 (filtered + sorted from _CorrData), checkboxes in A7+ (green) and B7+ (red) |

### Row-Hiding on Sheet 1 (v2.4.4)
After v2.4.4, Sheet 1 ("Corrected Roster Info") does NOT include students whose `student_id` appears in any cumulative tab (`_ApprovedData`, `_AdditionsData`, `_UnenrollData`, `_RejectedData`) with a timestamp within the last `HIDE_HANDLED_DAYS` days (default 7). Pipeline flow:
1. `compare_students` returns ALL current mismatches (unchanged)
2. `main()` reads handled IDs via `read_handled_student_ids(sheets_service, HIDE_HANDLED_DAYS)`
3. `_hide_recently_handled(corrections_map, corrections_sis, handled_ids)` filters parallel lists
4. Filtered lists are passed to `write_corrections` → `_CorrData` only contains unhandled students → Sheet 1 QUERY output shrinks

Hide-then-reappear semantics:
- Accepted/rejected student's timestamp < 7 days old → excluded from `_CorrData` → not on Sheet 1
- After 7 days, if mismatch still exists in MAP vs SIS → student reappears on Sheet 1 (signal: correction is stale, data team hasn't processed)
- After 7 days, if SIS updated → student naturally absent (no mismatch to flag)

Timestamp parsing: `datetime.strptime(row[0].strip(), "%Y-%m-%d %H:%M:%S")`. Rows with unparseable timestamps (e.g. legacy `4/23/2026` format from pre-v2.4.2 races) are silently skipped — treated as unhandled. `normalize_dates.py` should be run if there's a mix.

### Sheet 2: "Current Roster Info in SIS" (same layout, no checkboxes)
Same title/caption/filter/sort rows. SORT(QUERY()) formula in A7 pulls from `_SISData`.

### Sheets 3-5: Approval sheets (14 cols: Date + Mismatch Summary + 12 fields)
Same title/caption/filter/sort rows. SORT(QUERY()) formula in A7 pulls from hidden cumulative tab (A:N, 14 cols). Mismatch Summary is column B with red header. QUERY col refs: Campus=Col3, Grade=Col4, Level=Col5, StudentGroup=Col9, GuideEmail=Col12.
- Sheet 3 "Automated Correction List" — reads `_ApprovedData` (field mismatches)
- Sheet 4 "Roster Additions" — reads `_AdditionsData` ("Roster Addition" type)
- Sheet 5 "Roster Unenrollments" — reads `_UnenrollData` ("Unenrolling" type)

### Sheet 6: "Rejected Changes" (15 cols: Date + Mismatch Summary + 12 fields + Reason)
Same as Sheets 3-5 but with extra "Reason for Rejection" column (col O, blank for manual entry). QUERY reads `_RejectedData` A:N (14 cols); Reason is outside QUERY output.

### Real-Time Unenroll Queue (Live) Sheet
A 7th visible sheet in the corrections spreadsheet that shows IM-flagged students from all 9 campuses in real-time (~1 min latency). Built by `build_unenroll_queue.py`.

- Uses QUERY + IMPORTRANGE for each campus, 50-row block each (stacked vertically per campus)
- Per-campus Grade column differs: Reading CCSD uses Col8, all others use Col7 (due to Reading's Full Name insertion)
- WHERE clause uses `Col{N} = TRUE` (direct boolean), NOT `UPPER(Col{N}) = 'TRUE'` which fails with `#VALUE!` on boolean types
- IMPORTRANGE requires one-time 'Allow access' click by the human user when the tab is first opened
- Complementary to the hourly Python pipeline: Live Queue shows the flag instantly, Python does full SIS comparison hourly (Sheet 5)

### Weekly Snapshot Workflow (v2.5.1)

Separate from the main corrections spreadsheet, a weekly snapshot file is
generated each Monday bundling corrections not yet sent to support.

- Hosted in a Google Shared Drive "Weekly Corrections Archive" (id
  `0AFQGIqcKjsyFUk9PVA`). SA is Content Manager.
- File naming: `M/D Corrections` (e.g. `4/20 Corrections`), Monday of the
  current week in America/New_York timezone. No year, no zero-pad.
- Three tabs: `Correction List` (<- _ApprovedData), `Roster Additions`
  (<- _AdditionsData), `Roster Unenrollments` (<- _UnenrollData). Tabs with
  0 data rows are hidden (not deleted). Default `Sheet1` is deleted.
- 14-col header matches the approval sheets: Date Approved, Mismatch Summary,
  Campus, Grade, Level, First Name, Last Name, Email, Student Group, Guide
  First Name, Guide Last Name, Guide Email, Student_ID, External Student ID.
- `_RejectedData` is NOT included (rejected rows don't go to support).

Pipeline order (v2.5.1):
1. Compute current Monday in America/New_York
2. Read all 3 cumulative tabs and filter for this week
3. If total_rows == 0:
   a. Check if a file already exists for this week
   b. If no -> log "No corrections to send this week. File not created." -> exit
   c. If yes -> log + leave existing file untouched -> exit
4. Else: find-or-create the M/D Corrections file in Shared Drive
5. Write the 3 tabs (hide empty ones, delete default Sheet1)
6. Stamp col O of selected rows in cumulative tabs with current Monday ISO

### Sent Week column (v2.5.0)

Col O (0-indexed 14) on `_ApprovedData`, `_AdditionsData`, `_UnenrollData`
holds the ISO date of the Monday when the row was included in a weekly
snapshot (e.g. `2026-04-20`). Blank = unsent.

Selection rule in `generate_weekly_snapshot.py::filter_for_week`: include
rows where col O is blank OR equals current Monday ISO. This means:
- Rows accepted/approved earlier this week are always included.
- Rows that were included in THIS week's snapshot (already stamped) are
  still shown if you re-run during the same week — useful for refreshing
  the snapshot Wed/Thu after more corrections come in.
- Rows stamped with a PRIOR week's ISO date are excluded.

After selection, `main()` stamps col O of the selected rows with current
Monday ISO, so next week's run naturally excludes them.

Cumulative tabs have no header row. Apps Script `appendRow` writes new rows
starting at row 1 without a header, so col O defaults to blank for newly
accepted corrections.

### Support packet mode (v2.6.1)

`python generate_weekly_snapshot.py --all-unsent` runs the snapshot in
**support-packet mode**, an ad-hoc handoff for the data/SIS team:

- **Filter**: includes EVERY row with a blank `Sent Week` (any week, not
  just current Monday). The default no-flag run still uses the
  blank-OR-current-Monday rule from v2.5.0.
- **Instructions tab**: `WEEKLY_TAB_INSTRUCTIONS` (`"Instructions"`) is
  added and pinned at sheet index 0 so support sees it first on open.
  Content is built by `_build_instructions_rows()` — title, what each tab
  means, how to find a student in SIS, column reference (A-N), and a
  "do not edit this file, it is regenerated" warning. Formatting is
  applied by `_instructions_format_requests()`: navy title bg + white
  bold 16pt, h2 bold 13pt, h3 bold 11pt, col A wrapped at 900px, cols
  B-Z hidden.
- **Stamping**: behaves identically to default mode — every selected row
  gets stamped with the current Monday's ISO date in col O. So the next
  default Monday cron run naturally excludes anything already in the
  packet.
- **File reuse**: the support packet writes to the SAME `M/D Corrections`
  file as the regular weekly snapshot (find-or-create on the current
  Monday's filename). Re-running `--all-unsent` is idempotent — same
  file ID, same shareable link.

Default mode has NO Instructions tab. Adding the tab on the regular
Monday cron would be noisy (3 tabs stay simpler for routine use); it's
opt-in for ad-hoc support handoffs only.

## Timeback SIS bridge (v2.7.0)

Two campuses migrated off Dash/alpha_roster to Timeback's OneRoster API: **Vita High School** and **ScienceSIS**. The CMR tabs `"Vita High School (TimeBack)"` and `"ScienceSIS (TimeBack)"` mix into the same pipeline as the 9 Dash campuses, but the SIS-side cross-reference goes to the OneRoster API instead of `alpha_roster` BQ.

### Architecture
```
CMR Vita / ScienceSIS tab          alpha_roster BQ (Dash)         OneRoster API (Timeback)
   38 + 18 enrolled                ~8,700 deduped students        76 currently-rostered
        |                                  |                              |
        v                                  v                              v
read_map_roster()                   read_sis_data()              query_timeback_enrolled()
        |                                  |                              |
        |                                  +----- merge (Timeback wins) --+
        |                                                |
        +--------- compare_students() ------------------+
                            |
                     same 4 mismatch types as Dash
```

### Key files
- `timeback_sis.py` — self-contained OneRoster client. ~210 lines. Public API: `query_timeback_enrolled(timeback_campuses)`. Returns dict keyed by `legacyDashStudentId` shaped exactly like `query_alpha_roster` output.
- `keys/timeback-creds.json` — Cognito OAuth2 credentials. Gitignored. Format: `{"client_id": "...", "client_secret": "..."}`. In GHA, the workflow writes this from `TIMEBACK_CREDS_JSON` secret.
- `config.TIMEBACK_CAMPUSES` — dict mapping CMR tab name → school sourcedId UUID.

### Identifier bridge
The CMR `Student_ID` column for Vita/ScienceSIS rows holds Timeback's `metadata.legacyDashStudentId` (e.g. `066-6757`, `033-2154`), NOT the Timeback `sourcedId` UUID. So `query_timeback_enrolled` keys its returned dict by `legacyDashStudentId` to match what `compare_students` looks up.

Students whose `legacyDashStudentId` is empty in the OneRoster response are skipped (typically test accounts; ~10 of 86 students at the time of v2.7.0 launch).

### Empty Notes coercion
For Timeback CMR tabs, the Notes column is NOT maintained by IMs (the OneRoster API is the SIS source). `read_map_roster` coerces empty Notes to `"Enrolled"` so Timeback rows enter `map_enrolled` and the IM-checkbox unenroll path can fire. Dash campuses still skip empty-Notes rows (existing behavior).

### Merge precedence
On `student_id` collision (62 students currently exist in both alpha_roster and Timeback during the migration window), **Timeback wins**. This makes Timeback the authoritative source for those 62 students' field comparisons.

### Failure mode
On Timeback API failure (auth, network, rate limit after retries exhausted), `read_combined_sis_data` logs the error and continues with Dash data only. Vita/ScienceSIS students will surface as "Roster Addition" mismatches (because they're absent from the SIS dict) instead of correctly cross-referenced. Better to surface noise than to crash the pipeline. The next hourly run typically recovers.

### Cross-repo drift watch
`timeback_sis.py` duplicates a narrow slice of `timeback-data-pipeline/oneroster_client.py` (OAuth + GET /schools/{id}/students only). If the OneRoster API changes auth flow, endpoint paths, or response schema, BOTH files need updating. Look for `BASE_URL`, `TOKEN_URL`, and `get_students` in both repos.

### Field coverage for Timeback rows (v2.7.1)

`_find_mismatches` accepts an `is_timeback` flag. When True (i.e. when `map_rec["Campus"]` is in `config.TIMEBACK_CAMPUS_NAMES = {"ScienceSIS", "Vita High School"}`), the comparison only inspects fields the Timeback OneRoster API actually exposes:

| Field | Compared on Dash? | Compared on Timeback? |
|-------|-------------------|------------------------|
| Campus | yes | yes |
| Grade | yes | yes |
| First Name | yes | yes |
| Last Name | yes | yes |
| Email | yes | yes |
| Level | yes | **NO** — API doesn't expose |
| Student Group | yes | **NO** |
| Guide Email | yes | **NO** |
| Guide Name (first+last combined) | yes | **NO** |
| External Student ID | yes | **NO** — Timeback uses `sourcedId`, no SUNS-equivalent |

The IM-flagged Unenroll path runs identically for both: if `map_rec._unenroll_flag=True` AND `sis_rec.admissionstatus=="Enrolled"`, mismatch=Unenrolling and field comparison is short-circuited. The is_timeback flag only affects the *fallback* field-comparison path.

`timeback_sis.py` strips `" (TimeBack)"` from `campus_label` when generating each student's `Campus` value, so `"ScienceSIS (TimeBack)"` (CMR tab name) becomes `"ScienceSIS"` (matches what's in the Campus column on the CMR row). Without this, every Timeback row would mismatch on Campus.

### Staleness gotcha (v2.7.1 lesson)

When `setup_unenroll_columns.py` first wires up `IMPORTRANGE` on a new ISR, the formula returns `#REF!` until a human opens the CMR and clicks "Allow access". A pipeline run while `IMPORTRANGE` is still in `#REF!` state reads `_unenroll_flag=False` for every Timeback student, even ones with TRUE on their ISR. They will appear as field-mismatches instead of Unenrolling. Re-run after `IMPORTRANGE` resolves and the routing corrects itself.

This won't happen on hourly runs after the initial setup — `IMPORTRANGE` cache stays warm across runs.

## ISR (Individual Student Roster) Architecture

Each of 9 campuses has a dedicated Google Sheet called an ISR (Individual Student Roster). The ISRs feed the CMR (Consolidated MAP Roster), which in turn feeds the pipeline.

### ISR Tab Layout
- **`Student Roster` (SR)** — IMs edit here directly. Has an Unenroll checkbox column.
- **`MAP Roster` (MR)** — formula-derived view, not manually maintained. Values pull from:
  - `SR` via cell references (e.g. `='Student Roster'!H2`)
  - `Auto Synced from SIS` via VLOOKUP (Notes col O = SIS admission status)
  - Some plain values
  - MR is the tab that feeds the CMR via IMPORTRANGE.
- **`Auto Synced from SIS`** — SIS admission status dump (not directly maintained by IMs).
- **`School Info`** — Campus name to ID mapping for student_id derivation.

### Data Flow
```
SR (IM edits) ──formulas──> MR ──IMPORTRANGE──> CMR ──Python Sheets API──> generate_corrections.py
```

### ISR_CONFIG in config.py
Per-campus column positions are hard-coded in `config.py::ISR_CONFIG` with keys:
- `isr_id` — Google Sheet ID of the ISR
- `mr_gid` — gid of the `MAP Roster` tab
- `sr_unenroll_col` — 0-indexed column position of Unenroll on the SR tab
- `mr_unenroll_col` — 0-indexed column position of Unenroll on the MR tab

All 9 campuses are listed, each with their own SR/MR Unenroll column indices (positions vary because column layouts differ across campus sheets).

## Unenroll Workflow (option-C precedence)

Every student record in `read_map_roster` gets a `_unenroll_flag` boolean read from the CMR Unenroll column. In `compare_students`, for each ENROLLED MAP student, the logic is:

1. **FIRST check (option-C precedence)**: if `_unenroll_flag=TRUE` AND SIS `admissionstatus=Enrolled` → flag as `"Unenrolling"` and skip all other checks.
2. **Fall through**: otherwise, run existing checks — Roster Addition, Field mismatch, Notes-based Unenrolling.

**Option-C precedence** means the IM-driven Unenroll checkbox trumps field mismatches — we'd rather process the unenrollment than flag a grade/email change on a student who is about to be unenrolled anyway.

The pipeline prints a breakdown to stdout:
```
Unenrolling: N (IM-flagged: X, Notes-based: Y)
```

## Automation (Hourly Pipeline)

The pipeline runs automatically every hour via GitHub Actions:
- Workflow: `.github/workflows/hourly-pipeline.yml`
- Cron: `0 * * * *` (top of each hour, UTC)
- Trigger: schedule + workflow_dispatch (manual)
- Secret: `GCP_SA_KEY` contains verbatim `keys/sa-main.json`
- Concurrency group: single-flight (queue, don't cancel)
- Runtime: typically 15-20 seconds

Hybrid architecture:
- Real-time UX: "Unenroll Queue (Live)" sheet updates within ~1 min of IM checkbox change (via IMPORTRANGE)
- Hourly backstop: full pipeline with SIS comparison runs on schedule, updates Sheet 5 (Roster Unenrollments) and other approval sheets

### Weekly Snapshot Workflow
- File: `.github/workflows/weekly-snapshot.yml`
- Cron: `0 11 * * 1` (Mon 11:00 UTC = 07:00 ET standard / 06:00 ET DST)
- Trigger: schedule + workflow_dispatch
- Concurrency group: shares `weekly-corrections-pipeline` with the hourly
  workflow so they can't both write cumulative tabs at the same time
- Runtime: ~5-10s typically

### Notification model (v2.5.3+)

| Signal | Where | When |
|---|---|---|
| Single workflow failure | (silenced — user mutes default GHA emails) | Per failure |
| 3+ consecutive failures | GitHub Issue with label `pipeline-failure` | At threshold |
| Recovery | Comment on the tracking Issue + auto-close | Next success |
| Weekly trend | GitHub Issue with label `health-report` | Mon 12:00 UTC |
| Auto-deploy success/failure | Standard GHA notifications | Per push or workflow_dispatch |

Threshold is configurable via `env: THRESHOLD: '3'` in each workflow's
smart-notify step. Increase to `5` if you want even less noise; decrease
to `2` if you want earlier signal.

## Filtering & Sorting Mechanism

**Dropdowns filter AND sort via SORT(QUERY()) formulas — NOT Apps Script.**

The formula pattern for Sheet 1 (offset by 2 for accept/reject columns):
```
=IFERROR(SORT(QUERY('_CorrData'!A:M, 
  "SELECT * WHERE 1=1"
  & IF($C$5="All", "", " AND Col1='" & $C$5 & "'")
  & IF($E$5="All", "", " AND Col2='" & $E$5 & "'")
  ..., 0),
  MATCH($M$5, _Lists!F$2:F$14, 0),
  IF(OR($M$5="Grade"), FALSE, TRUE)), "")
```

- QUERY handles filtering (WHERE clauses reference dropdown cells)
- SORT wraps QUERY output, using MATCH to find column index from Sort By dropdown
- Sort direction: ascending (TRUE) for all text fields, descending (FALSE) for Grade and Date Approved
- Sort options stored in `_Lists` columns F-K (one per visible sheet)

**Checkbox handling**: When a filter or Sort By dropdown changes (row 5), Apps Script clears both Accept (col A) and Reject (col B) checkboxes because the QUERY output shifts.

**Accept/Reject workflow on Sheet 1**:
- **Col A (Accept Changes)** — light green (#D4EDDA) background. When checked, Apps Script reads col O (Mismatch Summary) and routes to:
  - `"Roster Addition"` → `_AdditionsData` → Sheet 4 auto-updates via QUERY
  - `"Unenrolling"` → `_UnenrollData` → Sheet 5 auto-updates via QUERY
  - field mismatches → `_ApprovedData` → Sheet 3 auto-updates via QUERY
- **Col B (Reject Changes)** — light red (#FEE2E2) background. When checked, all rejected rows → `_RejectedData` → Sheet 6 auto-updates via QUERY
- **Mutual exclusion**: Checking Accept unchecks Reject and vice versa
- Data columns read from C:N (12 cols), Mismatch Summary in col O (column 15)

## Mismatch Types

| Type | Mismatch Summary | Condition | Color |
|------|-----------------|-----------|-------|
| Roster Addition | "Roster Addition" | Enrolled in MAP, student_id not in SIS | Light green (#D4EDDA) |
| Field Mismatch | "Grade, Email" etc. | Enrolled in both, fields differ | Yellow (#FFF3CD) |
| Unenrolling | "Unenrolling" | IM-flagged Unenroll=TRUE OR Notes != "Enrolled" in MAP, admissionstatus = "Enrolled" in SIS | Light red (#FEE2E2) |

Colors are applied via conditional formatting rules on the Mismatch Summary column (Sheet 1 only), in priority order: Roster Addition → Unenrolling → NOT_BLANK (field mismatches).

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

## Configuration Constants

- `HIDE_HANDLED_DAYS = 7` in `config.py` — Tunable window for row-hiding on Sheet 1. Set to 0 to disable and restore always-show-everything behavior.

### Weekly Snapshot Constants (v2.5.0)

- `WEEKLY_SHARED_DRIVE_ID = "0AFQGIqcKjsyFUk9PVA"` — Shared Drive hosting the weekly snapshot files.
- `WEEKLY_SHARED_DRIVE_NAME = "Weekly Corrections Archive"` — Human-readable name.
- `WEEKLY_TIMEZONE = "America/New_York"` — ZoneInfo key used to compute the current Monday.
- `SENT_WEEK_COL = 14` — 0-based col O on cumulative tabs.
- `SENT_WEEK_HEADER = "Sent Week"` — mostly used for logging; cumulative tabs have no header row.
- `WEEKLY_TAB_CORRECTIONS = "Correction List"`, `WEEKLY_TAB_ADDITIONS = "Roster Additions"`, `WEEKLY_TAB_UNENROLLMENTS = "Roster Unenrollments"` — weekly file tab names.
- `WEEKLY_SOURCE_TABS` — dict mapping each weekly tab to its source cumulative tab (`Correction List` -> `_ApprovedData`, `Roster Additions` -> `_AdditionsData`, `Roster Unenrollments` -> `_UnenrollData`).
- `WEEKLY_HEADERS` — 14-col header list (Date Approved, Mismatch Summary, Campus, Grade, Level, First Name, Last Name, Email, Student Group, Guide First Name, Guide Last Name, Guide Email, Student_ID, External Student ID).

### Retry Helper Defaults (v2.5.2)

These are module-level constants in `retry_helper.py`, tunable in-place if
outages get longer:

- `DEFAULT_MAX_ATTEMPTS = 5` — total attempts before raising. With exponential
  backoff (1+2+4+8+16=31s of pure sleep) plus per-attempt API timeouts (~60s),
  total wall-clock coverage is ~5 minutes.
- `DEFAULT_BASE_DELAY = 1.0` — first sleep in seconds. Doubles each attempt.
- `DEFAULT_MAX_DELAY = 30.0` — cap on per-attempt sleep (jitter applied after).
- `TRANSIENT_HTTP_STATUSES = {408, 429, 500, 502, 503, 504}` — HTTP statuses
  treated as transient. Anything else (4xx auth errors, 4xx schema errors)
  raises immediately. Adjust if Google adds a new transient status code.

## Key Design Decisions

1. **QUERY formulas for filtering** — Apps Script `hideRows()` approach failed because it couldn't map dropdown positions to data columns. QUERY formulas auto-recalculate when dropdown cells change.
2. **Hidden data tabs** — Raw data on `_CorrData`/`_SISData`, QUERY on visible tabs. Same pattern as the dashboard pipeline's `_Data` tab.
3. **textFormatRuns for User Guide link** — The caption uses `updateCells` with `textFormatRuns` to create a clickable "User Guide" hyperlink mid-text.
4. **Checkbox stale-clearing** — When a dropdown changes, Apps Script clears both col A and col B checkboxes because QUERY output rows shift.
5. **Header auto-detection** — Handles schema differences across campus sheets.
6. **Batched API pre-writes** — Tab existence (1 read + 1 batch create), unmergeCells (1 batched call for all 6 sheets), clear values (`batchClear` for 9 tabs). ~5 pre-write API calls total instead of ~28.
7. **Three-level unenroll chain** — IM checks SR checkbox → formula mirrors to MR → IMPORTRANGE pushes to CMR → pipeline reads CMR Unenroll via `MAP_HEADER_MAP` auto-detection. No Apps Script needed.
8. **Hybrid real-time + hourly architecture** — Sheets-only live queue for instant IM feedback (bypass BQ); full pipeline on GitHub Actions hourly schedule for SIS comparison. No pure-Sheets solution exists because Sheets can't query BigQuery.
9. **Pre-formatted ISO date strings in Code.js** — Instead of `appendRow([new Date(), ...])` + post-write `setNumberFormat`, we pre-compute the date string via `Utilities.formatDate` and append the raw string. This is race-safe under concurrent onEdit triggers (Apps Script can fire multiple concurrent instances when a user toggles multiple checkboxes quickly). onEdit triggers are simple triggers (not installable), fire synchronously per edit, but Google may invoke multiple in parallel when edits happen within milliseconds. Trade-off: the column now stores strings, not serial dates, so numeric date sort relies on the ISO format being lexicographically equivalent to chronological order (it is).
10. **Hide accepted/rejected rows for 7 days** — Previously, Sheet 1 always showed every current mismatch, including students who had already been Accept'd or Reject'd. After handling, the pipeline rebuild cleared the checkbox but re-pulled the student, which confused IMs ("I already checked this, why is it back?"). v2.4.4 filters out students handled within `HIDE_HANDLED_DAYS`. After the window, if the mismatch still exists, the student reappears — which signals that the data team hasn't processed the correction yet. This behavior aligned with user expectation even though no prior version had implemented it.
11. **Shared Drive for weekly snapshot** — Service accounts have 0 bytes of Drive quota by default. Files created by the SA in its own Drive fail with `storageQuotaExceeded` (HTTP 403). Solution: Shared Drive hosts the weekly files — files there are owned by the drive itself, bypassing user quotas. All Drive API calls that touch the Shared Drive must include `supportsAllDrives=True`; `files.list` additionally needs `driveId=<shared_drive_id>`, `corpora='drive'`, `includeItemsFromAllDrives=True`.

### 12. Centralized retry strategy (v2.5.2)

All Google API calls now use the shared `retry_helper.retry_api(fn, ...)`
helper. The previous per-file helpers (`sheets_writer._retry_api`,
`generate_weekly_snapshot._retry`) had drift: one caught `Exception` (too
broad — masked programming bugs), the other caught only `HttpError`
(missed `TimeoutError`, which was half of the 2026-04-29 incident chain).

The shared helper:
- Catches ONLY transient API errors (HttpError 5xx/429/408, TimeoutError,
  socket.timeout, ConnectionError). Programming bugs raise immediately.
- 5 attempts × exponential backoff (1s, 2s, 4s, 8s, 16s) + 25% jitter.
  Total ~5 min coverage when API timeout is ~60s per attempt.
- Each retry logs the attempt count, exception summary, and wait time.
- Optional `label` parameter for caller-side identification in logs.

GitHub Actions workflows additionally use `nick-fields/retry@v3` to
re-run the entire Python step once if it exits 1, on top of the
in-script retries. Cron failure mode goes from "miss this hour" to
"miss this hour AND next hour" — practically never reached.

### 13. Failure-budget mindset & alert hygiene (v2.5.3)

The pipeline runs against Google APIs that have an SLA, not 100% uptime.
Asking "prevent ALL failures" is the wrong frame — failures will happen;
the question is whether they cause business impact and whether you hear
about the right ones.

The defense is layered:

1. **Absorb transient blips silently** (v2.5.2)
   - 5-attempt exponential retry in-script (~5 min coverage)
   - GHA workflow-level retry on top (~10 min more)
   - Total absorbs ~99% of cloud-API hiccups

2. **Only escalate persistent failures** (v2.5.3 smart-notify)
   - Each workflow's final step counts last-10 consecutive failures
   - Opens a tracking Issue ONLY at threshold (3 consecutive failures =
     ~3 hours of real outage for hourly, 3 weeks for weekly)
   - Auto-closes the Issue on next success
   - User mutes default workflow-failure emails; subscribes to label
     `pipeline-failure` for real signal

3. **Trend visibility** (v2.5.3 health-report)
   - weekly-health-report.yml runs Mon 12:00 UTC, opens Issue with
     30-day summary (success rate, failure streak, median duration)
     labeled `health-report`
   - User scans these weekly to detect slow degradation that wouldn't
     trigger smart-notify (e.g. success rate drifting from 99.5% → 95%
     over a month)

4. **State safety** (v2.5.3 row-stamp fix)
   - Stamping uses student_id lookup re-fetched at stamp time, not
     stored row numbers from the read pass. Eliminates the v2.5.1-known
     race window where Apps Script's removeStudentFromCumulativeTabs_
     could shift rows between snapshot's read and stamp.

What's deliberately NOT done: architectural rebuild (Cloud Scheduler,
DB-backed state, etc.). Reserved for if the layered defense above is
insufficient.

### 14. Apps Script as code (v2.6.0)

Before v2.6.0, Apps Script changes required manual copy-paste of `apps_script/Code.gs`
into the Extensions > Apps Script editor in the corrections spreadsheet. This created
a recurring failure mode: every release that touched Apps Script (v2.4.3, v2.4.4,
v2.5.x) had "user must re-paste Code.gs" as an outstanding action item. Skipping the
paste meant the live spreadsheet ran old script logic while the repo had newer logic.

v2.6.0 eliminates the paste step:

- Source of truth: `Code.js` at repo root (was `apps_script/Code.gs`).
- Deploy mechanism: clasp (Google's Apps Script CLI).
- Two deploy paths:
  1. Local: `npm run deploy` from project root → runs `node --check Code.js` then `clasp push`.
  2. GHA: `.github/workflows/deploy-apps-script.yml` triggers on push to main when
     `Code.js`, `appsscript.json`, or `.claspignore` changes. Uses `CLASPRC_JSON` +
     `CLASP_SCRIPT_ID` secrets. Fail-soft when secrets not configured (logs warning,
     skips silently — no email noise).
- Live Apps Script === HEAD on main, always (after secrets configured).

Pattern modeled after the email-automation project (which uses identical layout:
Code.js + appsscript.json + .claspignore + package.json with `npm run deploy`).

Architectural notes:
- Service account cannot deploy via clasp (clasp uses OAuth refresh tokens, not SA JWT).
  GHA uses USER OAuth tokens stored as the CLASPRC_JSON secret.
- The 9 ISR spreadsheets each have their own Apps Script project (Student Cards
  generator). v2.6.0 only auto-deploys to the corrections spreadsheet's project.
  ISRs still rely on manual paste — Code.js is polymorphic (checks for "Copy of MAP
  Roster" tab to gate Student Cards features), so it CAN be pasted into ISRs too,
  just not auto-deployed there.
- The fail-soft secrets check means committing v2.6.0 without secrets configured
  still pushes cleanly — the deploy workflow logs a warning instead of erroring.

## Common Bugs & Fixes

| Symptom | Cause | Fix |
|---------|-------|-----|
| QUERY returns error when name has apostrophe | Single quotes in values (O'Brien) break QUERY string | `_sq()` wraps cell refs with `SUBSTITUTE(cell,"'","''")` |
| Sheet 3 filters show wrong students | `_ApprovedData` has Date in Col1, shifting all cols by 1 | Fixed: Grade=Col3, Level=Col4, StudentGroup=Col8, GuideEmail=Col11 |
| Date Approved shows serial number (46126.5) | QUERY strips number formatting from source | `repeatCell` with `numberFormat` DATE_TIME on Sheet 3 col A |
| textFormatRun index error | "User Guide" at end of string | Only add trailing run if `link_end < len(text)` |
| addBanding error on re-run | Existing banding not cleared | `_clear_banding()` runs before formatting |
| CTAS fails with "mismatched input 'group'" | Reserved word quotes stripped by shell | Read SQL from file, pass via `--cli-input-json` |
| Campus shows 0 enrolled | Notes column at different index | Header auto-detection via `MAP_HEADER_MAP` |
| Dropdown doesn't filter | Using Apps Script hideRows instead of QUERY | Rewrote to use QUERY formulas referencing dropdown cells |
| Slow writes (~35s) | 15+ individual API calls for value ranges | Batched into 2 calls via `values().batchUpdate()` (~16s) |
| Slow pre-writes (~15s) | 28 individual API calls for tab checks, unmerge, clear | Batched into ~5 calls via `_ensure_all_tabs`, batch `unmergeCells`, `batchClear` |
| Duplicate student IDs silently overwritten | Same ID in multiple campus sheets | Warning logged; last-write-wins preserved |
| mergeCells error on re-run | Old merged cells conflict with new layout | `unmergeCells` runs before formatting on all visible sheets |
| "Invalid requests[0].setDataValidation: This operation is not allowed on cells in typed columns" | Column inside a Google Sheets Table (typed column feature) | `setup_unenroll_columns.py` catches this and skips `setDataValidation` (Table already provides checkbox rendering) |
| "Range exceeds grid limits. Max rows: X, max columns: 27" | SR tabs only had 27 cols, can't write to col 28 | Script auto-expands grid via `appendDimension` COLUMNS |
| QUERY `WHERE UPPER(col) = 'TRUE'` returns #VALUE! | Boolean values stored as Google Sheets booleans (not strings). UPPER() expects string input. | Use `WHERE col = TRUE` (no UPPER, no quotes) |
| Pipeline silently misses CMR Unenroll column | Default `A1:AC` range stopped at col 29 (AC). Unenroll is at col AD (29) for 7 campuses. | Expanded to `A1:AE` in `read_map_roster` |
| Live Queue shows `#REF!` or empty | IMPORTRANGE needs one-time human auth in the destination spreadsheet | User opens the tab and clicks 'Allow access' on the prompt |
| Inconsistent date formats in cumulative tabs | Apps Script race: `setNumberFormat(getLastRow(), 1)` fired by concurrent onEdit triggers applies to wrong row, leaving some rows in locale default (`4/23/2026 1:37:44`) and others in ISO (`2026-04-23 01:37:44`). Breaks chronological sort. | `Code.js` now pre-formats the date as ISO string via `Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")` BEFORE appendRow, eliminating the race. For historical data: `normalize_dates.py` migrates existing rows. |
| Accept/Reject col A/B colors missing after checkbox click | Pre-v2.4.3 `Code.js` called `setBackground(null)` on cols 1–15, wiping the permanent `#D4EDDA` / `#FEE2E2` applied by `sheets_writer.py`. Pipeline run re-applies them, but Apps Script wipes them again on next click. | v2.4.3+ `Code.js` only touches cols 3–15 (C:O), leaving A/B untouched. Auto-deployed via clasp post-v2.6.0; manual re-paste is no longer required. |
| Handled students keep reappearing on Sheet 1 | Pipeline had no handled-state tracking. | v2.4.4 `read_handled_student_ids` + `_hide_recently_handled` exclude them. |
| `UnicodeEncodeError: 'charmap' codec can't encode '->'` | `generate_weekly_snapshot.py` print statements | Windows cp1252 console doesn't support U+2192 (`->`) or U+2500 (`-`) box-drawing chars | Use ASCII `->` and `-` in print statements. Em-dash `—` (U+2014) works fine in cp1252. |
| `addBanding: You cannot add alternating background colors to a range that already has alternating background colors` | Re-run of `generate_weekly_snapshot.py` same week | `addBanding` isn't idempotent — errors if range already banded | In `main()`, fetch existing bandings via `spreadsheets.get(fields="sheets(properties.sheetId,bandedRanges)")` and queue `deleteBanding` requests BEFORE the new `addBanding` requests in the same batchUpdate. |
| `deleteSheet: You can't remove all the visible sheets` when 0 unsent rows for the week (deeper issue caught after v2.5.0's `addBanding` fix) | In `generate_weekly_snapshot.py::main()`, old order was: create-file -> read-cumulative -> if all 3 weekly tabs empty, hide all + delete Sheet1 -> Google Sheets rejects (0 visible tabs not allowed) | v2.5.1: restructured to read cumulative tabs FIRST. If `total_rows == 0`, log and exit before any file is created. If a file already exists from a prior run, leave it untouched. |
| Hourly pipeline crashes during `_ensure_all_tabs` with `HttpError 500 "Internal error encountered"` then `TimeoutError`, exits 1, next hourly run self-recovers | Sustained transient Sheets API hiccup (~3+ min). The old `_retry_api` had only 3 attempts × linear backoff (5s, 10s) = ~2 min coverage; outage outlasted the retry budget | v2.5.2: Replaced with shared `retry_helper.retry_api` — 5 attempts × exponential backoff with jitter (~5 min coverage), transient-only catch. Plus GHA `nick-fields/retry@v3` workflow-level retry as belt-and-suspenders. |
| Sent Week stamps land on wrong row in cumulative tab (rare; only when Apps Script edits the same tab during a snapshot run) | `generate_weekly_snapshot.py` stamped by row number stored at read time; Apps Script's `removeStudentFromCumulativeTabs_` could shift rows between read (~T) and stamp (~T+5s) | v2.5.3: stamp by student_id lookup re-fetched immediately before stamping. Re-read of col M is wrapped in `_retry`. Rows that vanished between read and stamp are silently skipped. |
| Apps Script changes don't take effect in the corrections spreadsheet | Pre-v2.6.0: manual paste skipped or forgotten. Post-v2.6.0: GHA secrets not configured, OR `npm run deploy` not run locally. | (a) Check GHA Actions tab — look for "Auto-deploy Apps Script (clasp)" run status. If skipped due to missing secrets, set CLASPRC_JSON + CLASP_SCRIPT_ID per README. (b) Run `npm run deploy` locally. (c) Verify with `clasp pull` to a temp dir + diff against repo. |

## Known Limitations

- **Grade sorts lexicographically** — "10" sorts before "2" because QUERY treats grades as text. Numeric sorting would require a helper column.
- **Cumulative hidden tabs grow indefinitely** — `_ApprovedData`, `_AdditionsData`, `_UnenrollData`, `_RejectedData` are never cleared. Manual cleanup needed periodically.
- **Banding covers 200 data rows** — Alternating row colors extend to row 206. If cumulative sheets exceed 200 approved rows, increase the `end_row` floor in `_format_visible_sheet()`.
- **Reason for Rejection column not in QUERY output** — Sheet 6 QUERY reads from `_RejectedData` A:M (13 cols). The "Reason for Rejection" header is in col N (col 14) but the column content is manually entered by IMs, not populated by QUERY.
- **Row-stamp race**: The stamping pass uses row numbers captured during the
  earlier read pass. If `Code.js::removeStudentFromCumulativeTabs_`
  deletes a row from a cumulative tab between the snapshot's read (~T) and
  stamp (~T+5s), the stored row number can shift, causing the stamp to land
  on the wrong row OR fail with "range not found". Probability is low (5s
  window + sporadic IM clicks). Mitigation for a future PR: stamp by
  student_id lookup at stamp-time instead of stored row number.
- **SA can trash but not permanent-delete in Shared Drive**: Service account
  has Content Manager role on `WEEKLY_SHARED_DRIVE_ID`. That allows
  `files.update(trashed=true)` but NOT `files.delete()`. To permanently
  delete an orphan, either escalate the SA to Manager role or wait for
  Shared Drive's auto-empty policy.
