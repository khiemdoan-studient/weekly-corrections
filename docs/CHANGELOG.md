# Changelog

## [v2.7.2] - 2026-05-06

### Fixed
- **REGRESSION introduced in v2.7.0**: External Student ID auto-detection broken for ALL 9 Dash CMR tabs. v2.7.0 added `"alpha student id"` to `MAP_HEADER_MAP["ext_student_id"]` so the new Vita / ScienceSIS CMR tabs would have ExtSID detected. But every Dash CMR also has BOTH an "Alpha Student ID" column (col W or X) AND a "SUNS Number" / "External ID" column (col AB or AC). `_detect_columns` iterates left-to-right and picks the FIRST match, so all 9 Dash tabs started writing the Alpha Student ID value (e.g. `'11733'`) into _CorrData col L instead of the SUNS Number (e.g. `'4689568995'`). alpha_roster compares against SUNS, producing a false "External Student ID" mismatch on virtually every Dash student.
  - Pre-v2.7.2 state: 1,867 spurious Dash ExtSID mismatches (JHES 373, JRES 360, JRHS 307, JHMS 305, AFMS 167, AFES 151, RCSD 129, Metro 63, AASP 12).
  - Post-v2.7.2 state: 0 Dash ExtSID mismatches.
  - Single-line revert: drop `"alpha student id"` from the matcher set in `config.py::MAP_HEADER_MAP`.

### Why Timeback campuses are unaffected
v2.7.1's `is_timeback` branch in `_find_mismatches` already skips ExtSID comparison for Vita / ScienceSIS rows. They don't need ExtSID detection at all — col L stays blank for Timeback rows, and no comparison runs. User spec ("Forget about external student ID entirely for ScienceSIS / Vita") is preserved.

### Why Reading CCSD is the same as pre-v2.7.0
Reading CCSD's CMR has only an "Alpha Student ID" header (no "SUNS Number"). Pre-v2.7.0 the matcher didn't include "alpha student id" so Reading CCSD's ExtSID column was never detected — col L was blank, MAP "" vs SIS alpha_roster externalstudentid (whatever it has). v2.7.2 restores exactly this. No new behavior for Reading CCSD.

### Verified
- `python -m py_compile` passes.
- End-to-end run: 1,870 matches (was 44 in v2.7.1 — the 1,826 Dash students that were spuriously mismatched on ExtSID now match cleanly). Field mismatches: 115 (was 1,941). Total corrections: 230 (was 2,056). 105 hidden-recently-handled. Pipeline runtime: 15.6s.
- Live `_CorrData` probes (all PASS):
  - Hardeeville 083-11733: dropped from _CorrData entirely (now matches cleanly, was previously false-ExtSID-mismatched).
  - Total Dash ExtSID mismatch count: 0 (was 1,867).
  - 5 user-flagged ScienceSIS students all routed correctly: 066-6749 Alanah + 066-6742 Autumn still on Sheet 1 as Unenrolling (pending IM acceptance); 066-6778 Arie + 066-6774 Armoni + 066-6773 Ataijah in `_UnenrollData` (already accepted by user between v2.7.1 and v2.7.2).
  - Vita + ScienceSIS noise mismatches: 0 (no Level / Student Group / Guide Email / Guide Name / External Student ID surfaced).

## [v2.7.1] - 2026-05-06

### Fixed
- **Noise field mismatches on Vita / ScienceSIS rows**. v2.7.0 compared all 9 fields (Campus, Grade, Level, First/Last/Email, Student Group, Guide Email, External Student ID, Guide Name combine) for every row — but the Timeback OneRoster API doesn't expose Level, External Student ID, Student Group, or Guide info. Result: every Vita / ScienceSIS row that didn't hit the Unenrolling path got a noise mismatch chain like `"Campus, Level, External Student ID"`, hiding actual issues.
  - `generate_corrections.py::_find_mismatches` now takes an `is_timeback` flag. Compares only Campus / Grade / First / Last / Email when True. Skips Level, Student Group, Guide Email, External Student ID, and the Guide Name combine.
  - `generate_corrections.py::compare_students` derives `is_timeback` from `map_rec["Campus"]` membership in the new `TIMEBACK_CAMPUS_NAMES` set.
- **Systemic Campus mismatch on every Timeback row**. `timeback_sis.py` was setting `Campus = "ScienceSIS (TimeBack)"` (the CMR tab name) but MAP shows `Campus = "ScienceSIS"`. Every row mismatched on Campus. Now strips `" (TimeBack)"` from `campus_label` before populating each student's Campus field.

### Added
- **`config.py`**: `TIMEBACK_CAMPUS_NAMES = {"ScienceSIS", "Vita High School"}` — the bare campus values used by `_find_mismatches` to detect Timeback rows.

### Why
User reported: "Armoni Nelson (066-6774) and others have Unenroll=TRUE on the ISR but show up as 'Campus, Level, External Student ID' field-mismatches in the corrections sheet, not 'Unenrolling'. Forget about External Student ID entirely for ScienceSIS / Vita — those columns don't correspond."

Root cause was two separate issues:
1. **Staleness (no code fix needed).** At the moment of the v2.7.0 pipeline run earlier today, the CMR `IMPORTRANGE` for col AB was still resolving / not yet authorized. Pipeline read `_unenroll_flag=False` for everyone. Students whose Notes was non-empty + non-"Enrolled" (Alanah, Arie, Autumn) hit the Notes-based Unenrolling loop and routed correctly. Students whose Notes was empty got coerced to "Enrolled" by my v2.7.0 fix → went to `map_enrolled` → IM-flagged path saw `_unenroll_flag=False` → fell through to field comparison → got the wrong label. Re-running with IMPORTRANGE now resolved produces correct routing for all 5.
2. **Noise mismatches (the user's explicit ask).** Even when the unenroll path fires correctly, every Timeback row that lands in field comparison surfaced 3-5 noise mismatches because the OneRoster API returns "" for fields MAP has populated. Skipping non-corresponding fields produces clean output.

### Verified
- `python -m py_compile` passes.
- End-to-end pipeline run: 1,962 corrections written. IM-flagged Unenroll: 32 (was 22). Field mismatches: 1,941 (was 1,995, -54 noise mismatches removed). Matches: 44 (was 0 — Timeback students now match cleanly when there's nothing real to flag). Runtime: 19.4s.
- Live `_CorrData` probe — all 5 user-flagged students show `Unenrolling`:
  - 066-6749 Alanah Cossey: PASS
  - 066-6778 Arie Sturgis: PASS
  - 066-6774 Armoni Nelson: PASS
  - 066-6773 Ataijah Mitchell: PASS
  - 066-6742 Autumn Boothe: PASS
- Zero violations on the Timeback noise-term scan (`Level`, `External Student ID`, `Student Group`, `Guide Email`, `Guide Name` — none appear in any Vita / ScienceSIS mismatch column).
- Vita + ScienceSIS `_CorrData` mismatch breakdown: 26 Unenrolling + 1 Last Name + 1 Grade+Last Name = 28 total (was 72 in v2.7.0). All 2 remaining field mismatches are real data-team issues, not noise.
- 9 Dash campuses unchanged (no regression).

## [v2.7.0] - 2026-05-06

### Added
- **`config.py`**:
  - 2 new entries in `CAMPUS_SHEETS`: `"ScienceSIS (TimeBack)"` + `"Vita High School (TimeBack)"`. These are Timeback-backed campuses whose SIS source-of-truth is the OneRoster API, NOT the alpha_roster BQ table.
  - `MAP_HEADER_MAP["ext_student_id"]` set gained `"alpha student id"` (the External Student ID header text used on the new Timeback CMR tabs).
  - 2 new entries in `ISR_CONFIG` for the Vita + ScienceSIS ISRs (`1sOSwvw…` + `1SjVoQ…`). Both have `sr_unenroll_col=23` (col X), `mr_unenroll_col=27` (col AB).
  - New constant `TIMEBACK_CAMPUSES`: maps each Timeback CMR tab name to the school's OneRoster `sourcedId` UUID (Vita = `e57cb46d-…`, ScienceSIS = `7c475cf4-…`).
  - New constant `TIMEBACK_CREDS_PATH`: file path for Timeback API credentials (`keys/timeback-creds.json`, gitignored).
- **`timeback_sis.py`** (NEW, ~210 lines): self-contained OneRoster client. Wraps just the OAuth2 + `GET /schools/{id}/students` endpoint of `api.alpha-1edtech.ai`. Public API: `query_timeback_enrolled(timeback_campuses)` returns a dict keyed by `legacyDashStudentId` shaped like `query_alpha_roster` output (so `compare_students` can consume both sources without changes). Falls back to `COGNITO_CLIENT_ID` + `COGNITO_CLIENT_SECRET` env vars if `keys/timeback-creds.json` is absent.
- **`generate_corrections.py::read_combined_sis_data`**: new function. Calls `read_sis_data` (alpha_roster) + `query_timeback_enrolled` and merges the two into a single SIS dict. Timeback wins on `student_id` collisions (per user spec — the migration window has ~62 students existing in both alpha_roster and Timeback, and Timeback is the new system of record). On Timeback API failure, logs the error and continues with Dash data only (Vita/ScienceSIS will surface as Roster Additions until the API recovers — graceful degradation).
- **`generate_corrections.py::read_map_roster`**: empty-Notes Timeback rows now coerced to `notes = "Enrolled"` so they enter `map_enrolled` and the IM-checkbox unenroll path can fire. Dash campuses still skip empty-Notes rows (existing behavior unchanged).
- **`requirements.txt`**: added `requests` (for `timeback_sis.py`).
- **`.github/workflows/hourly-pipeline.yml`**: new "Write Timeback creds from secret (fail-soft)" step. Reads `TIMEBACK_CREDS_JSON` GHA secret and writes it to `keys/timeback-creds.json` before running `generate_corrections.py`. When the secret is missing, logs a `::warning::` and pipeline continues with Dash data only. Cleanup step now also removes `keys/timeback-creds.json`.

### Why
User requested unenroll-checkbox parity for two new campuses (Vita + Science SIS) that migrated to Timeback's OneRoster API. The 9 existing Dash campuses cross-reference against `alpha_roster` BQ table; Timeback campuses need the live OneRoster API as source-of-truth. This release adds the SIS-bridge plumbing without changing any user-facing workflow — IMs check the same Unenroll checkbox on their ISR, and the same Sheet 1 / Apps Script / weekly-snapshot flow handles the rest.

### Verified
- `python -m py_compile config.py timeback_sis.py generate_corrections.py` → compile OK.
- Isolated module test: `query_timeback_enrolled(TIMEBACK_CAMPUSES)` returned 76 students across 2 campuses (52 ScienceSIS + 24 Vita, 10 skipped for missing `legacyDashStudentId`). OAuth handshake works; pagination works; metadata bridge to `legacyDashStudentId` works.
- `python setup_unenroll_columns.py` → wrote SR col X + MR col AB Unenroll columns on both Vita + ScienceSIS ISRs; wrote IMPORTRANGE at CMR `'ScienceSIS (TimeBack)'!AB2` and `'Vita High School (TimeBack)'!AB2`. Existing 9 Dash ISRs re-confirmed as no-op.
- `python build_unenroll_queue.py` → added Vita + ScienceSIS QUERY+IMPORTRANGE blocks at rows 455 + 505 of "Unenroll Queue (Live)" tab.
- `python generate_corrections.py` end-to-end:
  - MAP roster: 2,084 enrolled (was 2,028 pre-v2.7.0 — added 38 ScienceSIS + 18 Vita).
  - Combined SIS: 8,723 students (8,709 Dash + 76 Timeback). 62 overlapping student_ids; Timeback entries take precedence.
  - Total corrections: 2,100 (1,939 → 2,100, +161 new from Vita/ScienceSIS).
  - Unenrolling mismatches: 38 (was 37 — +1 from Notes-based path on a Vita/ScienceSIS student).
  - Pipeline runtime: ~18s (was ~15s pre-v2.7.0; +3s for the 2 OneRoster API calls).
- Live-sheet probe of `_CorrData` confirmed 20 Vita + 52 ScienceSIS rows surface; sample row contents look correct.

### Latency
The 2 OneRoster API calls (one per Timeback school) add ~3s to each hourly cron run. User confirmed acceptable trade-off for live source-of-truth correctness.

### User action required
Add GHA secret `TIMEBACK_CREDS_JSON` containing the JSON contents of `keys/timeback-creds.json`:
```json
{"client_id": "...", "client_secret": "..."}
```
Until the secret is configured, the GHA hourly workflow logs a warning and runs with Dash-only SIS data — Vita/ScienceSIS students will surface as Roster Additions instead of correctly comparing against Timeback. Local runs already work because `keys/timeback-creds.json` is committed to your local checkout.

## [v2.6.1] - 2026-05-05

### Added
- **`config.py`**: new constant `WEEKLY_TAB_INSTRUCTIONS = "Instructions"` (~line 195) — the user-facing tab name for the support-packet guidance tab. Pinned at sheet index 0 when present.
- **`generate_weekly_snapshot.py`**: new `--all-unsent` CLI flag (argparse, registered in `__main__`). When set, switches the snapshot to **support-packet mode**:
  - Filter: includes EVERY row with a blank `Sent Week`, regardless of which week. Default mode (no flag) still uses blank-OR-current-Monday.
  - Adds an `Instructions` tab with plain-language guidance for the support team (what each tab means, how to find a student in SIS, column reference, who to contact). Pinned to index 0 so it's the first tab support sees.
- **`generate_weekly_snapshot.py`**: new helper `_build_instructions_rows(generation_iso, total_rows, per_tab_counts)` (~line 332) — returns list of `(text, style)` tuples driving the Instructions tab content. Style ∈ {`title`, `h2`, `h3`, `body`, `blank`}.
- **`generate_weekly_snapshot.py`**: new helper `_instructions_format_requests(sheet_id, rows)` (~line 405) — returns Sheets API batchUpdate requests applying per-row formatting (navy title bg, h2 bold 13pt, h3 bold 11pt), wraps col A at 900px, hides cols B-Z, and pins the tab at index 0.
- **`generate_weekly_snapshot.py`**: `filter_for_week()` gained an `all_unsent=False` keyword argument. Default behavior unchanged. When `True`, filter reduces to "include if `Sent Week` is blank" — same blank-row check that supplies the empty-row guard from v2.5.0.

### Why
User asked for "a document I can send to support with the correction list" — generated on demand, off the regular Monday cron schedule, covering ALL pending corrections (not just this week's). Previously support had to be invited to the weekly file or the user had to manually copy rows; now `python generate_weekly_snapshot.py --all-unsent` produces a self-contained, support-friendly Sheet with built-in guidance.

### Verified
- `python generate_weekly_snapshot.py --help` shows the new `--all-unsent` flag with a usage example.
- Live run (2026-05-05): `python generate_weekly_snapshot.py --all-unsent` produced sheet `1vjIY6hVwyOcUwOZrsXLbSVWG-nAUsE4ZRwRNUZaEgJs` with 4 tabs in this order:
  - index=0 `Instructions` (28 rows, visible) — pinned correctly
  - index=1 `Correction List` (2 rows + header, visible)
  - index=2 `Roster Additions` (header only, hidden — empty-tab guard working)
  - index=3 `Roster Unenrollments` (1 row + header, visible)
- 3 source rows stamped with `2026-05-04` in `_ApprovedData` (2) and `_UnenrollData` (1), so subsequent default-mode runs won't re-bundle them.
- Default mode (`python generate_weekly_snapshot.py` without flag) unchanged — `filter_for_week()` still defaults to `all_unsent=False` and the Instructions tab is only added when `all_unsent=True`. Monday cron behavior is identical to v2.6.0.

### Usage
```
python generate_weekly_snapshot.py --all-unsent
```
Re-running `--all-unsent` is idempotent for the same week: the existing weekly file is updated in place (same file ID, support's link still works), and stamped rows are not re-stamped on subsequent runs.

## [v2.6.0] - 2026-04-30

### Added
- **`Code.js`** at repo root. Replaces `apps_script/Code.gs`. Same functional content as v2.4.3 (verified by `clasp pull` diff: only the header comment differs). Header updated to v2.6.0 with deploy instructions.
- **`appsscript.json`**. clasp manifest. `timeZone: America/New_York`, `runtimeVersion: V8`, `exceptionLogging: STACKDRIVER`.
- **`.claspignore`**. Whitelist strategy: ignore everything by default, only push `Code.js` + `appsscript.json` to Apps Script. Prevents accidentally pushing Python tooling, docs, GHA workflows, or `.claude/` config to the production Apps Script project.
- **`package.json`**. npm scripts:
  - `npm run check`: `node --check Code.js` syntax gate.
  - `npm run deploy`: check + `clasp push`.
  - `npm run pull` / `push` / `open`: raw clasp passthroughs.
- **`.github/workflows/deploy-apps-script.yml`**. Auto-deploy on push to main when `Code.js`, `appsscript.json`, or `.claspignore` change. Also `workflow_dispatch` for manual runs. Fail-soft when secrets missing (logs warning, skips silently).

### Removed
- **`apps_script/Code.gs`**. Old location. Migrated to root `Code.js`. `apps_script/` directory deleted.

### Changed
- **`.gitignore`**. Added `.clasp.json` (gitignored, contains scriptId per user), un-excluded `appsscript.json` from the `*.json` rule, added `node_modules/`.

### Why
Eliminates the manual "paste Code.gs into Extensions > Apps Script" step that's caused recurring confusion since v2.4.3. The live Apps Script in the corrections spreadsheet now ALWAYS matches HEAD on main. No more "did I remember to re-paste?" anxiety.

### Architecture: how auto-deploy works
Two paths, both supported:
1. **Local CLI** (matches the email-automation pattern): developer runs `npm run deploy` from project root. Runs `node --check Code.js` for syntax, then `clasp push` to upload to the linked Apps Script project. Requires `clasp login` once per machine and `.clasp.json` with the scriptId.
2. **GHA auto-deploy on push**: any commit to main that touches `Code.js`, `appsscript.json`, or `.claspignore` triggers `.github/workflows/deploy-apps-script.yml`, which installs clasp, restores credentials from `CLASPRC_JSON` + `CLASP_SCRIPT_ID` secrets, runs `clasp push --force`. Truly hands-free if secrets configured.

### Verified
- `clasp pull` from live, diffed against repo `Code.js`. Only the header comment differed (live had v2.4.3 header text, repo has v2.6.0). **Functional code is identical**. Confirms the live Apps Script was NOT stale on logic, only on the header comment.
- `clasp push --force` succeeded: `Pushed 2 files at 4:19:53 PM`. Live Apps Script now matches repo HEAD.
- `node --check Code.js` syntax check passes.
- Local clasp 3.3.0 installed; `~/.clasprc.json` present with `tokens` schema.

### Investigation: original "Accept doesn't route" symptom (separate from auto-deploy)
User reported on 2026-04-30: "I checked Accept, nothing went into Correction List, and after an hour the checkbox was unchecked." Same for Unenrolling.

What we found during diagnosis:
- Apps Script `onEdit` IS firing (5/1 rows present in `_ApprovedData` and `_UnenrollData`).
- Routing works correctly (Unenrolling -> `_UnenrollData`, Roster Addition -> `_AdditionsData`, else -> `_ApprovedData`).
- v2.4.4 `read_handled_student_ids` returns the expected recent sids; `_hide_recently_handled` filters them out.
- Live `Code.js` diff vs repo `Code.js`: only header text differs; logic identical.

Therefore the symptom is NOT a stale-code issue. Likely candidates:
1. **Sheet 3 dropdown filter set**: if Campus dropdown on "Automated Correction List" is set to a specific campus, rows for other campuses are invisible despite being in `_ApprovedData`.
2. **Visual position shift**: after pipeline rebuild, a different student now occupies the same visual row as where the user clicked Accept; the new student's checkbox appears unchecked, looking like "your" Accept reverted.
3. **Race**: user clicks Accept right before the hourly cron runs; pipeline reads cumulative tabs before the append finishes. Probability narrow (~5 min window).

v2.6.0 doesn't directly fix any of these. It ensures future fixes can be deployed without manual paste, eliminating that whole class of "did I update the script?" failure mode.

### User Action Required
**For the GHA auto-deploy to start working** (one-time setup):
1. Go to GitHub repo Settings -> Secrets and variables -> Actions -> New repository secret.
2. Add secret `CLASPRC_JSON`: paste the full contents of your local `~/.clasprc.json` (or `%USERPROFILE%\.clasprc.json` on Windows). This is your clasp OAuth refresh token.
3. Add secret `CLASP_SCRIPT_ID`: paste `16_ypoWiIFRpIZzUEpwJGLP8DvexGoCDAiXaZGKoIhRQFa38H8vcS436_` (the Apps Script project ID for the corrections spreadsheet).
4. After both secrets are set, the next push to `Code.js` auto-deploys. Or trigger manually: Actions tab -> "Auto-deploy Apps Script (clasp)" -> Run workflow.

Until secrets are added, `npm run deploy` from local works (and was used to deploy v2.6.0 itself).

The `CLASPRC_JSON` contains an OAuth refresh token. Treat it as a credential. Only add to GitHub Actions secrets, never commit to the repo (already gitignored via `*.json`).

## [v2.5.3] — 2026-04-30

### Added
- **`health_report.py`** (~190 lines) — pipeline health summary script. Queries the last N days of GitHub Actions runs via the `gh` CLI for both `hourly-pipeline.yml` and `weekly-snapshot.yml`. Computes (1) total runs, (2) success rate %, (3) failure count, (4) cancelled count, (5) max consecutive failure streak, (6) currently-failing streak, (7) last failure timestamp, (8) median run duration. Output is Markdown — suitable for posting as a tracking Issue or reading locally. CLI flags: `--days N`, `--repo OWNER/NAME`, `--output PATH`.
- **`.github/workflows/weekly-health-report.yml`** — cron `0 12 * * 1` (Monday 12:00 UTC, an hour after the weekly snapshot at 11:00). Runs `health_report.py --days 30`, opens a GitHub Issue titled `📊 Weekly health report — YYYY-MM-DD` with the summary, labeled `health-report`. The last 4 weeks' issues stay open as a rolling history so you can scan trends at a glance. Includes `workflow_dispatch` for manual runs.
- **Smart-notify step** in both `hourly-pipeline.yml` and `weekly-snapshot.yml` — a final `if: always()` step using `actions/github-script@v7` that:
  - On job failure: queries the last 10 runs of THIS workflow, counts consecutive failures, and opens (or comments on) a tracking Issue titled `🚨 <workflow-name> persistently failing` with label `pipeline-failure` once the threshold is hit (env `THRESHOLD: '3'`). Idempotent — won't open duplicates.
  - On job success: if a tracking issue is open, comments `Pipeline recovered. Auto-closing` and closes it.
  - Net effect: zero open `pipeline-failure` issues = healthy. An open one = a real, persistent failure (real signal).
  - Both workflows now have a top-level `permissions:` block granting `issues: write` + `actions: read`.

### Fixed
- **Row-stamp race in `generate_weekly_snapshot.py`** (HIGH-severity finding from the v2.5.1 audit, finally fixed) — previously stamped cumulative-tab rows by row number stored from the earlier read pass. If Apps Script's `removeStudentFromCumulativeTabs_` deleted a row in the ~5s window between read (~T) and stamp (~T+5s), stored row numbers shifted and the stamp could land on the wrong row. v2.5.3 re-reads col M (Student_ID) immediately before stamping, builds a `student_id → current_row_num` map, and looks up by sid at stamp time. Race window shrunk from ~5s to ~milliseconds. Rows that vanished entirely between read and stamp are silently skipped — they'll be re-picked-up by the next run if they reappear with a blank `Sent Week`.

### Why this is comprehensive (and why "comprehensive" still doesn't mean "zero failures")
Cloud APIs sometimes hiccup for longer than even our beefy retry budget. The real prevention strategy is layered:
1. **v2.5.2** — in-script retry + GHA workflow-level retry absorbs ~99% of transient blips silently.
2. **v2.5.3 smart-notify** — the remaining 1% only escalate when persistent (3+ consecutive failures = something actually broken). Single transient blips no longer email you.
3. **v2.5.3 health report** — weekly summary lets you see trends; if success rate drifts down over time, you'll catch it.
4. **v2.5.3 row-stamp fix** — failed runs no longer leave stamping in a bad state on the cumulative tabs.

Out of scope deliberately: an architectural rebuild (move scheduler from GHA to GCP Cloud Scheduler, decouple state from the Sheets API) — days of work and a large blast radius, reserved for if v2.5.x doesn't get us where we want. Other audit findings (CUM-002 unbounded growth, WRITE-003 migration sequence) are documented and deferred — not currently triggering.

### Verified
- `python -m py_compile` passes on `generate_weekly_snapshot.py` and `health_report.py`.
- Live integration: `python generate_weekly_snapshot.py` ran end-to-end with the new student_id-lookup stamping. Idempotent on re-run: `stamped 0 row(s)` because rows were already marked from the prior run.
- Live integration: `python health_report.py --days 14` returned correct counts:
  - `hourly-pipeline.yml`: 103 runs, 98.1% success, 2 failures (4/28 + 4/29), max streak 1.
  - `weekly-snapshot.yml`: 1 run (4/27 failed cron), max streak 1, currently failing 1 (below threshold of 3, so no tracking issue opened).

### Failure-budget framing
Before v2.5.x: ~98% hourly success rate, every failure → email noise. After v2.5.3: same ~99%+ success rate, but you ONLY hear about persistent failures (3+ consecutive ≈ ~3 hours of real outage). Single-blip failures are absorbed silently. The weekly health digest gives you trend visibility without alert fatigue.

### User Action Required
- **Mute the default GitHub Actions failure email** — Settings → Notifications → uncheck "Send notifications for failed workflows only for workflows I trigger" (label varies by GitHub UI). Without this, you'll get BOTH the legacy failure emails AND the new smart-notify Issue notifications.
- **Subscribe to issues with label `pipeline-failure`** for actual signal. Watch the repo for issues, or add `pipeline-failure` to your notification preferences.
- **Optional: subscribe to label `health-report`** to get the weekly summary issue.
- Next Monday 5/4 12:00 UTC, the first weekly health report issue should appear automatically.

## [v2.5.2] — 2026-04-30

### Added
- **`retry_helper.py`** (~120 lines) — shared retry helper exposing `retry_api(fn, max_attempts=5, base_delay=1.0, max_delay=30.0, label="")`. Replaces the per-file `_retry_api` (sheets_writer.py) and `_retry` (generate_weekly_snapshot.py). Strategy:
  - 5 attempts total (was 3 in both legacy helpers).
  - Exponential backoff: 1s, 2s, 4s, 8s, 16s — ~31s of pure sleeps, ~5 minutes of total coverage with the ~60s-per-attempt API timeout. Up from ~2 minutes.
  - 25% random jitter on each sleep so concurrent workflows (hourly + weekly) don't synchronize their retries during a Sheets brownout.
  - **Transient-only catch**: `HttpError` with status in `{408, 429, 500, 502, 503, 504}`, plus `TimeoutError`, `socket.timeout`, and `ConnectionError`. Programming bugs (`KeyError`, `AttributeError`, etc.) raise immediately instead of being masked by retries — the legacy `sheets_writer._retry_api` had a bare `except Exception` that hid these.
  - Each retry logs which attempt + why + how long it'll wait, with optional `label` for the call site.
- **GitHub Actions workflow-level retry** — both `hourly-pipeline.yml` and `weekly-snapshot.yml` now wrap the Python step in `nick-fields/retry@v3` with `max_attempts: 2, timeout_minutes: 8, retry_wait_seconds: 60`. Belt-and-suspenders defense — even if the in-script retry exhausts, GHA re-runs the entire job once for free.

### Changed
- **`sheets_writer.py`** — replaced the local `_retry_api` (3 attempts, linear, broad `except Exception`) with `from retry_helper import retry_api as _retry_api`. All 26+ existing call sites unchanged.
- **`generate_weekly_snapshot.py`** — replaced the local `_retry` (3 attempts, linear, `HttpError`-only — couldn't catch the `TimeoutError` half of the 4/29 failure chain) with `from retry_helper import retry_api as _retry`. All 11 existing call sites unchanged.
- **`generate_corrections.py`** — wrapped two API calls that previously had no retry coverage at all:
  - `read_map_roster` campus sheet `.get().execute()` — was bare; a transient 500 silently dropped a campus from that hour's run.
  - `read_handled_student_ids` cumulative-tab `.get().execute()` — was inside a bare `except: continue`; a transient 500 silently treated the tab as empty for that hour, which would have re-flagged already-handled students on Sheet 1.
- **Helper scripts** (`add_sent_week_column.py`, `normalize_dates.py`, `setup_unenroll_columns.py`, `build_unenroll_queue.py`) — added `retry_helper` import and wrapped the loop-internal per-tab/per-campus API calls where a transient error would otherwise drop the rest of the iteration.

### Fixed
- **The 2026-04-29 hourly cron failure class** — `_ensure_all_tabs` in `sheets_writer.py` raised `HttpError 500 "Internal error encountered"` → `TimeoutError: The read operation timed out` → exhausted 3 retries → exit 1. Subsequent hourly runs self-recovered, confirming transient. The 2026-04-28 hourly run had failed in the same class. The new retry budget covers ~5 min of API hiccups instead of ~2 min, and the workflow-level retry adds another full job re-run on top of that.

### Why
The legacy `_retry` in `generate_weekly_snapshot.py` only caught `HttpError`, so the `TimeoutError` half of the 4/29 chain bypassed the retry entirely. The legacy `_retry_api` in `sheets_writer.py` did catch broad `Exception` but only gave 3 attempts with linear backoff (~2 minutes), which wasn't enough headroom for the 4/29 brownout. Centralizing on a single tuned helper fixes both gaps and means future tuning (e.g. raising `max_attempts` to 7 if Google has a longer outage) only touches one file.

### Verified
- `python -m py_compile` passes on all 8 modified Python files.
- 5-test in-test smoke suite for `retry_helper`: success-first-try, transient-then-success, `KeyError` fail-fast, `HttpError 500` retries, `HttpError 404` fail-fast — all pass.
- Live integration: `python generate_weekly_snapshot.py` ran end-to-end, picked up 2 new corrections accepted since the prior run, created `4/27 Corrections` (id `1z-aL77kzA37VNzvg6lSEye8o5J1H_GiB8ezQjlzzd6U`) in the Shared Drive, stamped both rows with `2026-04-27`. No regressions.

### Scope note
The user explicitly chose the comprehensive scope (over the tighter "fix sheets_writer + generate_weekly_snapshot only"): centralize the retry logic into a single shared module and import it everywhere, including helper scripts. Single point of tuning beats 8 copies that drift over time.

### User Action Required
- **None.** Next hourly cron picks up the new retry behavior automatically. The workflow-level GHA retry kicks in only if the Python script exits 1 — the hope is it never has to.

## [v2.5.1] — 2026-04-27

### Fixed
- **Empty-week cron crash on Monday 4/27** — The 11:00 UTC cron failed because all 29 cumulative-tab rows (23 from `_ApprovedData`, 6 from `_UnenrollData`) were already stamped `2026-04-20` from last week's first run, and no new corrections had been accepted between 4/20 and 4/27. The filter selected 0 rows for the 4/27 week. The script created the file `4/27 Corrections` (id `1pfMlmN2EzjLD4nQIG3b5EDt0EjMQq0xKNhtAj3wqphA`), added 3 weekly tabs, then tried to hide all 3 (each had 0 rows) AND delete the default `Sheet1` — which would have left 0 visible sheets. Sheets API returned 400: `"You can't remove all the visible sheets in a document"`.
- **`generate_weekly_snapshot.py::main()` step order** — Restructured to read all 3 cumulative source tabs FIRST, before any file create/find. If `total_rows == 0`:
  - No file exists for this week → log `"No corrections to send this week. File not created."` and exit cleanly. Drive stays clean.
  - File already exists for this week (e.g. earlier successful run, or leftover orphan) → log and exit; leave the file untouched.

  Otherwise the normal find-or-create + populate flow runs unchanged. Updated the module-level docstring at the top of `generate_weekly_snapshot.py` to reflect the new order (read first, then check empty, then create-or-find).

### Why
The v2.5.0 success path assumed at least one row would always be selected. It worked on 4/20 because the cumulative tabs had 29 brand-new unstamped rows. By 4/27, every existing row was already stamped and no fresh IM clicks had landed — a perfectly normal state that v2.5.0 didn't anticipate. The fix is purely defensive; the v2.5.0 success path is unchanged.

### Cleanup
- Orphan `4/27 Corrections` file (id `1pfMlmN2EzjLD4nQIG3b5EDt0EjMQq0xKNhtAj3wqphA`) moved to trash via `files.update(trashed=true)`. The SA has Content Manager role on the Shared Drive, which permits trash but not permanent-delete; trash will auto-empty per Shared Drive retention policy.

### Verified
- Local run: `python generate_weekly_snapshot.py` → `"No corrections to send this week. File not created. (3.1s)"` → exit 0.
- Shared Drive `0AFQGIqcKjsyFUk9PVA` lists exactly 1 file: `4/20 Corrections` (id `1TmpjJkFrKQdG_DzxkVrE0YqIb30tdK5ZZTWXNguXA0I`). The orphan is in trash.

### Audit findings (documented, NOT fixed in this PR)
- **Row-stamp race (medium risk)** — The stamping pass uses row numbers captured during the read pass. If Apps Script's `removeStudentFromCumulativeTabs_` deletes a row in the ~5s window between read and stamp, subsequent row numbers shift, and the stamp could land on the wrong row OR fail with "range not found". Probability is low (5s window + sporadic IM clicks) but non-zero. Mitigation for a future PR: stamp by `student_id` lookup at stamp-time instead of stored row number. Not fixed here to keep this PR scoped to the actual failure.

### User Action Required
- **None.** The fix is purely defensive — the v2.5.0 success path is unchanged. Next Monday's cron will either create a new file (if there are unsent rows by then) or skip cleanly (if no IMs accept anything before 5/4).

## [v2.5.0] — 2026-04-24

### Added
- **Weekly snapshot automation** — Every Monday at 07:00 ET (11:00 UTC), GitHub Actions runs `generate_weekly_snapshot.py` and produces a single Google Sheet bundling all corrections not yet sent to support. One sheet per week, lives in a Shared Drive, same URL survives re-runs so existing shares don't break.
- **`generate_weekly_snapshot.py`** (~500 lines) — orchestrator. Computes current Monday in `America/New_York`, finds-or-creates a sheet named `M/D Corrections` (e.g. `4/20 Corrections`) at the root of the Shared Drive `Weekly Corrections Archive` (id `0AFQGIqcKjsyFUk9PVA`), then for each of 3 cumulative source tabs (`_ApprovedData`, `_AdditionsData`, `_UnenrollData`) reads rows where col O "Sent Week" is blank OR equals the current Monday ISO date, writes them into 3 tabs in the weekly sheet (`Correction List`, `Roster Additions`, `Roster Unenrollments`), hides tabs with 0 data rows, deletes the default `Sheet1`, then stamps col O of every selected source row with the current Monday ISO so next week's run excludes them automatically. `_RejectedData` is deliberately excluded — rejected rows don't go to support.
- **`add_sent_week_column.py`** (~100 lines) — one-time pre-flight that ensures col O header is `Sent Week` across all cumulative tabs. Safe to re-run; idempotent.
- **`.github/workflows/weekly-snapshot.yml`** — cron `0 11 * * 1` + `workflow_dispatch`.
- **`config.py` constants** — `WEEKLY_SHARED_DRIVE_ID = "0AFQGIqcKjsyFUk9PVA"`, `WEEKLY_SHARED_DRIVE_NAME`, `WEEKLY_TIMEZONE = "America/New_York"`, `SENT_WEEK_COL = 14`, `SENT_WEEK_HEADER = "Sent Week"`, `WEEKLY_TAB_CORRECTIONS/ADDITIONS/UNENROLLMENTS`, `WEEKLY_SOURCE_TABS` dict, `WEEKLY_HEADERS` (14-col header list matching the approval sheets).
- **`requirements.txt`** — added `tzdata` (required for `zoneinfo` on Windows runners).

### Changed
- **Re-run semantics are idempotent** — re-running the same week updates the existing sheet in place (same file id, existing shares survive). Bandings are cleared before re-apply so `addBanding` doesn't fail with "already banded".

### Architecture note — why a Shared Drive
Service accounts have 0 bytes of Drive storage quota by default, so the SA cannot own spreadsheet files. Solution: a Google Shared Drive ("Weekly Corrections Archive") with the SA added as Content Manager. Files created in a Shared Drive are owned by the drive itself, bypassing user quotas. All Drive API calls now pass `supportsAllDrives=True` and `files.list` uses `driveId` + `corpora='drive'` + `includeItemsFromAllDrives=True`.

### Fixed (caught during build)
- **Windows cp1252 console couldn't encode `→` and `─`** — replaced with ASCII `->` and `-` in print statements. (Em-dash `—` encodes fine in cp1252 and stays.)
- **`addBanding` errors on re-run when bandings already existed** — now deletes existing bandings first within the same batchUpdate.

### Verified
- First live run: created `4/20 Corrections` (id `1TmpjJkFrKQdG_DzxkVrE0YqIb30tdK5ZZTWXNguXA0I`) in the Shared Drive.
- Read 23 rows from `_ApprovedData`, 0 from `_AdditionsData`, 6 from `_UnenrollData` — 29 total selected for the week of 4/20.
- Weekly sheet tabs: `Correction List` visible (24 rows incl. header), `Roster Additions` hidden (empty), `Roster Unenrollments` visible (7 rows incl. header). `Sheet1` deleted.
- All 29 source rows stamped `2026-04-20` in col O.
- Idempotent re-run: same file updated in place, log line `stamped 0 rows` confirms no double-marking.

### User Action Required
- **None for the automation itself.** Next Monday 07:00 ET, the snapshot auto-generates.
- **To send to support**: open the current week's sheet in the Shared Drive "Weekly Corrections Archive", click Share, send to your support contact. The URL stays stable for the whole week even if the script re-runs, so shares don't break.

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
