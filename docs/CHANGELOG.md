# Changelog

## [v2.9.2] - 2026-05-25

Add JRES (Ridgeland Elementary School) to summer school.

### Added
- **JRES added to `SUMMER_TABS`**. Standard 39-col Jasper layout, so summer cols land at SR AC..AG / MR+CMR AE..AI. All 20 provided students matched existing roster rows (incl. variants Axel Santos Lopez -> Axel Lopez Santos, Miirian -> Mirian Diaz Pinto, Shania Delonila -> Shania Deonilla, and Ava Deloach -> Avah Deloach via the new lenient tier); 0 ambiguous, 0 unmatched.
- JRES is Jasper: Subjects = "Language and Fast Math" (uniform), no teacher (blank), grade = provided (4/5). Parent / contact / transportation columns from the source list were disregarded per request.
- **Lenient match fallback tier** in the loader: fires ONLY when the strict matcher (clean fuzzy-subset + 3+token) returns zero candidates, allowing short first-name typos (prefix, len diff <=1, e.g. "Ava" -> "Avah") to pair when the surname matches exactly and the result is unique. Strict matches keep precedence, so no regression to existing clean matches.

### Verified
- `python -m py_compile` passes; provisioning idempotent (4 existing schools unchanged, JRES added, combined tab rebuilt) in ~34s.
- Live: combined `Summer School Roster` = 263 rows across 5 schools (AFMS 24, JHES 25, JHMS 53, JRES 20, JRHS 141), every row Summer School = TRUE; JRES samples show grade + "Language and Fast Math".

### Files changed
- `setup_summer_school_columns.py`, `docs/CHANGELOG.md`, `docs/AI_INSTRUCTIONS.md`.

## [v2.9.1] - 2026-05-25

Add AFMS (Allendale Fairfax Middle School) to summer school, and codify the combined Summer School Roster tab.

### Added
- **AFMS added to `SUMMER_TABS`** in `setup_summer_school_columns.py`. Provisioning is layout-aware (find-or-append by header), so AFMS's narrower 29-col CMR tab gets its summer columns at SR AB..AF / MR AC..AG / CMR AD..AH (vs the Jasper schools' SR AC..AG / MR+CMR AE..AI). All 24 provided AFMS students matched existing roster rows (incl. variants Clarke Antoine -> Antoine Clark, Sanders Zayala -> Zayla Sanders, Osborne Ezekiel -> Ezekiel Osborn); 0 ambiguous, 0 unmatched.
- AFMS specifics vs the Jasper schools: per-student subject ("Language" for Lang, "Math" for Math) instead of the uniform "Language and Fast Math"; no teacher (blank, like Ridgeland); grade = provided (6/7/8).
- **`build_summer_roster_tab()`** codifies the combined `Summer School Roster` tab (previously created ad-hoc) into the script. One live QUERY over all `SUMMER_TABS`, normalized to core A:N + the 5 summer columns via a per-school `{core, summer}` horizontal join, so the summer flag is always output column 15 and ONE QUERY (`where Col15 = true`) filters across schools regardless of their differing absolute summer-column positions. Re-run the script to refresh after the school list changes.

### Verified
- `python -m py_compile setup_summer_school_columns.py` passes; provisioning idempotent (3 existing schools unchanged, AFMS added, combined tab rebuilt) in ~26s.
- Live: combined `Summer School Roster` = 242 rows (AFMS 24, JHES 24, JHMS 53, JRHS 141), every row Summer School = TRUE (the normalized QUERY handles both the 39-col Jasper and 29-col AFMS layouts). AFMS samples show correct subject (Language/Math) and grade as plain numbers.

### Files changed
- `setup_summer_school_columns.py`, `docs/CHANGELOG.md`, `docs/AI_INSTRUCTIONS.md`.

## [v2.9.0] - 2026-05-25

Summer School columns across the roster chain (ISR -> MAP Roster -> Combined MAP Roster) for 3 Jasper schools.

### Added
- **`setup_summer_school_columns.py`** (idempotent provisioner, mirrors `setup_unenroll_columns.py`). For the 3 summer-school campuses (JHMS Hardeeville Jr/Sr, JHES Hardeeville Elementary, JRHS Ridgeland Secondary) it appends 5 columns AFTER the existing last-used column at each layer (nothing existing shifts):
  - **Student Roster (SR)**: `Summer School` (checkbox) + `Summer School Teacher Email` + `Summer School Teacher` + `Summer School Grade` + `Summer School Subjects`. Typed source data.
  - **MAP Roster (MR)**: same 5 as `=ARRAYFORMULA('Student Roster'!<col>2:<col>)` mirrors (matches the existing MR column style).
  - **Combined MAP Roster (CMR)**: same 5 as `=IMPORTRANGE(ISR,"MAP Roster!<col>2:<col>")` (the Unenroll precedent). Appended at cols AE..AI on all 3 campus tabs.
  - The grade column is forced to plain NUMBER format (appended cells otherwise inherited a date format and rendered grades like `01/07/1900`).
- One-time data load (via a gitignored `_scratch_*` loader; student PII never committed) set the flag + teacher + grade + subjects for the provided students. Subjects = "Language and Fast Math" for all 3 (all Jasper/JCSD). Teachers: JHMS per-student (Janice Allen / Tashanique Douglas / Avlen Edwards), JHES Mahogany Salisbury / Queenie Henry, JRHS left blank per request. JHES summer grade taken from each student's existing roster grade (the provided list gave only teacher bands).

### Matching (names -> existing roster rows)
The provided students were name-only (no IDs) and mostly already in each roster under fuller/variant names (compound surnames, middle names, hyphenation, occasional typos). Matcher: normalize (NFKD ascii, lowercase, strip punctuation, drop <=1-char middle initials) then a clean fuzzy-subset match (one token set is a fuzzy subset of the other, >=2 shared, unique) with a fallback for 3+token names sharing >=2 tokens (one mismatched token each side, e.g. "Ivan Dario Funez" -> "Ivan D Funez Banegas"). Clean matches take precedence; uniqueness is required, so any over-match becomes ambiguous and is never written. Result: **213 written, 3 ambiguous, 11 unmatched** (reported for manual review; no new roster rows created, no guesses).

### Verified
- `python -m py_compile setup_summer_school_columns.py` passes; provisioning idempotent (ran twice clean).
- Live SR -> MR -> CMR propagation confirmed on sample students per school (flag TRUE, correct teacher/grade/subjects; grade displays as a plain number; non-summer students blank; existing Unenroll column untouched).

### Files changed
- `setup_summer_school_columns.py` (new), `.gitignore` (add `_scratch_*`), `docs/CHANGELOG.md`, `docs/AI_INSTRUCTIONS.md`.

## [v2.8.4] - 2026-05-25

Email-fallback matching so students who are in both MAP and SIS but missing their MAP Student ID stop showing as false "Add to MAP Roster", plus a loud warning when a campus sheet yields zero students.

### Added
- **Email-fallback matching (`generate_corrections.py`)**. Matching was 100% Student-ID-keyed with no fallback, so a MAP row with a blank or wrong Student ID could never match its SIS record.
  - `read_map_roster` now keeps blank-Student-ID rows that have an email (previously dropped at the `if not student_id: continue` gate) in a new `map_emailonly` list, and builds `all_map_emails` (every lowercased MAP email).
  - `compare_students` builds a `sis_by_email` index. Loop 1 tries an email match before declaring "Roster Addition" (catches typo'd ids). A new loop flags email-only rows that match the SIS by email as a **"Student ID"** correction, stamped with the correct SIS id so the IM knows what to enter and each row gets a distinct hide key. The "Add to MAP Roster" loop now skips any SIS student whose email already exists in the MAP via an email-only row.
  - Result for the reported students: Vita's Yember Aguilera (033-2634) and Dauz Lee-Zablotny (033-8879), enrolled in OneRoster but with blank MAP ids, now show as "Student ID" corrections (fill in the id) instead of false "Add to MAP Roster". 10 such students surfaced across Vita / JHMS / JHES / AFMS / Reading.
- **Loud zero-student warning (`read_map_roster`)**. If a campus sheet processes 0 students out of N data rows, it prints a `*** WARNING` with the detected columns. This surfaced that Reading CCSD (Dash) processes 0 of 1,199 rows: its Notes column is all-blank on a non-Timeback sheet, so every row is skipped. Pre-existing gap, now visible.

### Investigated (no code change)
- **Metro Mariam Hassan (079-10490) + JHES dylan/eleazar/keyon (083-15135/15138/15140)**: genuinely absent from `alpha_roster` (verified by fullid + name). Correct "Roster Addition" (real adds to SIS).
- **Gabriela Condari**: no "Condari" in the SIS and not in the managed MAP; "domain_disabled / external user" is an unmanaged Google account. Not actionable from this pipeline.
- **JHMS (Hardeeville Junior & Senior High) is NOT invisible.** Its MAP col-A header reads '4', but `read_map_roster` already has a fallback (assume col A = Student ID when Notes is detected and >=5 cols matched), so all ~330 students ARE processed (a `NOTE` prints each run). The '4' lives in the SOURCE MAP Roster master (`1g8KU...`), pulled into the JHMS tab via `=IMPORTRANGE(...,"MAP Roster!A:N")`. An attempt to overwrite the JHMS-tab A1 deleted that importrange and briefly blanked the campus; it was immediately restored. A literal 'Student ID' header can only be set in the source sheet, and it is cosmetic since the fallback works.

### Verified
- `python -m py_compile generate_corrections.py` passes.
- Live run: Student ID = 10 (incl. Yember 033-2634 + Dauz 033-8879, each carrying their SIS id in `_CorrData`); `_MapAdditionsData` has no Yember/Dauz; Mariam + the 3 JHES students still "Roster Addition"; JHMS processed (309 enrolled); the Reading CCSD 0-student warning fires.

### Files changed
- `generate_corrections.py`, `docs/CHANGELOG.md`, `docs/AI_INSTRUCTIONS.md`.

## [v2.8.3] - 2026-05-25

Bulletproof "Reason for Rejection" persistence: a pipeline-side capture step so typed reasons survive every refresh even if the onEdit Apps Script bridge is dead.

### Added
- **`sheets_writer.py::_capture_typed_reasons`** + a call at the start of `write_corrections` (before any clear/reorder). It reads Sheet 6 `M7:O` (student_id + reason, still aligned to the prior render) and upserts any non-blank reason into the durable `_RejectionReasons` tab BEFORE the A:N batchClear can orphan col O. `_hydrate_rejection_reasons` then re-renders col O from the now-complete store.
  - This is the THIRD persistence guard, joining the onEdit write path (`handleRejectionReasonEdit_` / `upsertRejectionReason_`) and the hydrate read path. Crucially it makes reason persistence INDEPENDENT of the Apps Script. The root cause of the repeated wipes was the onEdit bridge being dead from 2026-05-08 to 2026-05-25 (the stale-deploy gap fixed in v2.8.1/v2.8.2), during which typed reasons were never saved and the hourly hydrate overwrote col O with blanks within the hour.
  - Only non-blank reasons are written (never a blank over a stored reason), so the capture never fights a deliberate clear made through the live onEdit.

### Verified (live end-to-end, the decisive proof)
- Wrote a unique marker reason directly via the Sheets API to a rejected student's col O (sid 086-12777). This SIMULATES a reason typed while onEdit is dead (API writes do not fire onEdit). Ran the full `generate_corrections.py`. Result: the pipeline logged "Captured 1 newly-typed reason(s) into _RejectionReasons before hydrate", the marker persisted in `_RejectionReasons`, AND survived on Sheet 6 col O after the rebuild. Snapshot/restore cleanup returned `_RejectionReasons` to its exact 11 rows. TEST RESULT: PASS.
- `python -m py_compile sheets_writer.py generate_corrections.py` passes.

### Recovery of older wiped reasons (investigated, mostly unrecoverable)
- Exhaustive probe: 11 reasons survive in `_RejectionReasons`. All 39 retained Google Drive revisions of the corrections sheet were downloaded and scanned, recovering 0 reasons beyond the 11. Drive prunes aggressively for this heavily-edited file: only 2026-05-22 to 2026-05-25 is retained (everything earlier is gone), the Sheets UI Version History draws from the same pruned set, and the weekly snapshot files never stored col O. Pre-2026-05-22 reasons that were never durably saved are unrecoverable except from an external copy, which `restore_rejection_reasons.py` can still ingest.

### Investigated, no change (ScienceSIS "in MAP not SIS" report)
- 10 ScienceSIS students a user could not find in the SIS are all currently Enrolled per the OneRoster API. They sit in `_UnenrollData` because an IM flagged them to leave (2026-05-07 x2, 2026-05-13 x8 in one batch) while still Enrolled in the SIS. "Unenrolling" is working as designed: `compare_students` requires SIS admissionstatus==enrolled to flag it, and a student MISSING from OneRoster becomes "Roster Addition", not "Unenrolling". The alpha_roster BQ returning 0 for them is expected (ScienceSIS lives in OneRoster, not that export). No code change.

### Files changed
- `sheets_writer.py`, `docs/CHANGELOG.md`, `docs/AI_INSTRUCTIONS.md`.

## [v2.8.2] - 2026-05-25

GHA Apps Script auto-deploy is now actually wired up. Closes the v2.8.1 audit's biggest finding (live script was 4 versions stale because the deploy workflow silently skipped every run).

### Fixed
- **`deploy-apps-script.yml` was silently skipping every run** because `CLASPRC_JSON` + `CLASP_SCRIPT_ID` were never set (the workflow is fail-soft: missing secrets log a `::warning::` and exit success). Configured both repo secrets:
  - `CLASPRC_JSON` piped straight from the local `~/.clasprc.json` (clasp 3.x `tokens.default` schema) via `gh secret set CLASPRC_JSON < ~/.clasprc.json`. The OAuth refresh token goes to GitHub's encrypted secret store without ever touching the transcript or logs.
  - `CLASP_SCRIPT_ID` = `16_ypoWiIFRpIZzUEpwJGLP8DvexGoCDAiXaZGKoIhRQFa38H8vcS436_`.
- Going forward, any push to main touching `Code.js` / `appsscript.json` / `.claspignore` auto-deploys. No more manual `npm run deploy`. If the OAuth token ever expires the deploy FAILS loudly (red X) instead of silently skipping: strictly better than before. Recovery: local `clasp login`, then re-run the `gh secret set CLASPRC_JSON` pipe.

### Verified (the deploy run IS the build verification; no app code changed)
- `gh secret list` shows both `CLASPRC_JSON` + `CLASP_SCRIPT_ID` (timestamps 2026-05-25).
- Test deploy (`gh workflow run deploy-apps-script.yml`, run 26385874847) completed in 21s with the push steps RUN, not skipped: `Check secrets configured` then `Restore clasp credentials` then `Push to Apps Script` then `Summary` all green, and the Summary emitted the `::notice::` "Apps Script deployed via clasp push to scriptId ***". A successful "Push to Apps Script" also proves the OAuth refresh token is still valid.
- `clasp pull` of the live script confirms all 10 feature markers (handleRejectionReasonEdit_, upsertRejectionReason_, Add to MAP Roster, _MapAdditionsData, LockService). No-op push, since the live script was already current from the v2.8.1 manual deploy.

### Note
- GitHub flagged Node.js 20 action deprecation (actions/checkout@v4, actions/setup-node@v4): GitHub forces these to Node 24 on 2026-06-02. Non-blocking for now; bump the action versions before then.

### Files changed
- GHA repo secrets (`CLASPRC_JSON`, `CLASP_SCRIPT_ID`), `docs/CHANGELOG.md`, `docs/AI_INSTRUCTIONS.md`.

## [v2.8.1] - 2026-05-25

Audit-driven fixes (full-pipeline /audit). No CRITICAL bugs found; these are the MEDIUM/LOW items worth fixing.

### Fixed
- **`Code.js` onEdit race (MEDIUM, observed)**: `removeStudentFromCumulativeTabs_` + `appendRow` was not lock-protected. Concurrent onEdit instances (rapid checkbox toggling, or the mutual-exclusion `setValue(false)` firing a second onEdit) could interleave and leave DUPLICATE cumulative-tab rows. This is the mechanism behind the 2 duplicate tuples (079-10545, 083-11566) the v2.8.0 audit found. Now wrapped in `LockService.getDocumentLock().tryLock(10000)`: if another instance holds the lock the edit is skipped (next toggle / hourly run reconciles), and `releaseLock()` runs at the end (the document lock also auto-releases on script termination). Impact was cosmetic (handled_keys is a set, dedupes), but the dup rows are now prevented at the source.
- **`timeback_sis.py::_OneRosterClient._get` silent partial failure (MEDIUM)**: on persistent 401/429 (rate-limit) across all retry attempts, `_get` returned `None`, which made `get_students` treat it as end-of-pagination (`if not data: break`) and silently return partial/empty results. `query_timeback_enrolled` then returned an incomplete dict with NO exception, so Vita/ScienceSIS students silently vanished (surfacing as Roster Additions) with nothing logged. Now raises `RuntimeError` on exhaustion, so `read_combined_sis_data`'s existing try/except logs it and degrades to Dash-only loudly. Same data outcome, far better observability. (At current scale these schools fit in one OneRoster page, so this is a robustness/observability fix, not an active data bug.)
- **`sheets_writer.py` `_RejectionReasons` not re-hidden by the pipeline (LOW-MED)**: the tab was hidden only by the one-time v2.7.4 migration; `write_corrections` never re-hid it. Now captured as `rejection_reasons_id` and added to the hide loop (consistent with `_MapAdditionsData`). No-op if already hidden; re-hides if anyone unhides it.
- **`sheets_writer.py:939` stale comment**: the `clear_ranges` comment said col O hydrates from `_RejectedData`; v2.7.4 moved storage to `_RejectionReasons`. Comment corrected (and a pre-existing em dash removed).

### Audit dismissals (verified false positives)
- compare_students "double-count": the `student_id in map_enrolled` guard already covers matched students.
- MAP_HEADER_MAP "substring collision": `_detect_columns` uses exact set membership, not substring.
- "Add to MAP Roster appends SIS to both lists": intentional v2.8.0 design (data lives in SIS).
- compute_monday "timezone drift": `datetime.now(ZoneInfo(...))` is tz-aware.
- Code.js deleteRow "off-by-one": cumulative tabs have no header row.
- Shared-Drive find-or-create race: mitigated by the GHA single-flight concurrency group.

### Deferred (LOW, not fixed)
- `read_handled_student_keys` 5 sequential reads could be one `batchGet` (~1-2s; changes per-tab partial-failure semantics).
- `alpha_roster_ctas.sql` dedup tiebreaker non-deterministic when admissionstatus + student_group both NULL (one-time/sibling export).
- `config.py` ISR_CONFIG stale "NEEDS TO BE ADDED" comments (cosmetic).
- Empty-week orphan files accumulate in the Shared Drive (known clutter).

### Verified
- `python -m py_compile` (timeback_sis, sheets_writer, generate_corrections, config) + `node --check Code.js` pass.
- Regression run: 27 students written, 39.6s, no errors. `_RejectionReasons` and `_MapAdditionsData` confirmed hidden=True.

### Deploy gap discovered + resolved (CRITICAL, the audit's biggest catch)
While verifying the v2.8.1 Code.js deploy, found the **live Apps Script in the corrections spreadsheet was still v2.6.0** (2026-05-01). The GHA `deploy-apps-script.yml` workflow has been **silently skipping the clasp push on every run since v2.6.0** because the `CLASPRC_JSON` / `CLASP_SCRIPT_ID` secrets were never configured. The workflow is fail-soft (logs a `::warning::` and exits success), so every run reported "success" while actually deploying nothing. Prior sessions (mine included) read that "success" as "deployed".

Impact: every Apps Script change since v2.6.0 was never live:
- v2.7.3/v2.7.4 Reason-for-Rejection onEdit bridge (`handleRejectionReasonEdit_`, `upsertRejectionReason_`): NOT live. New reasons IMs typed on Sheet 6 col O since 2026-05-08 were never captured to `_RejectionReasons`; the Python hydration only ever surfaced the 11 reasons from the one-time migration. Any newly-typed reason was effectively dropped on the next pipeline run.
- v2.8.0 "Add to MAP Roster" accept routing + `_MapAdditionsData`: NOT live. An accepted Add-to-MAP row would have fallen through to `_ApprovedData` (wrong tab). No damage yet (Sheet 7 had 0 accepted rows).
- v2.8.1 onEdit lock: NOT live (just shipped).

Resolution: ran `npm run deploy` locally (clasp is configured on this machine: `.clasp.json` + `~/.clasprc.json`). Pushed the current Code.js. Verified via `clasp pull` into a temp dir: all feature markers now present in the live script (handleRejectionReasonEdit_, upsertRejectionReason_, Add to MAP Roster, _MapAdditionsData, LockService). The Reason bridge + Add-to-MAP routing + lock are finally live.

### User action required (prevent recurrence)
Configure the two GitHub secrets so the auto-deploy actually pushes (per README "Apps Script auto-deploy setup"): `CLASPRC_JSON` (contents of `~/.clasprc.json`) + `CLASP_SCRIPT_ID` (`16_ypoWiIFRpIZzUEpwJGLP8DvexGoCDAiXaZGKoIhRQFa38H8vcS436_`). Until then, run `npm run deploy` locally after ANY Code.js change. (Note: the live header comment still reads v2.8.0; the v2.8.1 lock code IS present. Cosmetic, will sync on next Code.js change.)

### Files changed
- `Code.js`, `timeback_sis.py`, `sheets_writer.py`, `docs/CHANGELOG.md`, `docs/AI_INSTRUCTIONS.md`.

## [v2.8.0] - 2026-05-19

### Added
- **New mismatch type "Add to MAP Roster" + new Sheet 7 "Missing from MAP Roster".** Detects students enrolled in the SIS who have NO row in the MAP roster (the reverse of "Roster Addition", which is in-MAP-not-in-SIS). Previously these were invisible: `compare_students` only iterated MAP students, so a SIS-only student was never flagged in any direction.
- **`generate_corrections.py::compare_students`**: a 3rd detection loop iterates `sis_students`. A student is flagged "Add to MAP Roster" when: not in MAP (enrolled or non-enrolled), `admissionstatus == "Enrolled"`, Campus is one of the managed campuses, and not a test account.
  - **Managed-campus scoping**: `alpha_roster` is a global Alpha export (~9,400 students incl. hundreds from unmanaged schools: TSA, Colearn, Alpha Miami, etc.). Scope = the set of distinct Campus values present in the MAP roster. Verified that MAP Campus values exactly match SIS campus values, so membership cleanly isolates the 11 managed campuses.
  - **Test-account filter** (`_is_test_account`): skips rows whose first+last name contains "test" (case-insensitive). Removed 3 on first run (Test Metro, Pasco Test, Test JHES).
- **`config.py`**: `TAB_MAP_ADDITIONS = "Missing from MAP Roster"`.
- **`sheets_writer.py`**: Sheet 7 plumbing mirroring Sheet 4 (Roster Additions): `SORT_OPTS_SHEET7`, `_build_sorted_query_sheet7` (src `_MapAdditionsData!A:N`), new `_MapAdditionsData` cumulative tab (14-col, hidden) + `TAB_MAP_ADDITIONS` visible sheet, `_Lists` sort col L (`sort7`), row block, `_format_visible_sheet` call, red Mismatch-header + DATE_TIME-format loops, `_migrate_cumulative_tabs` + `_backfill_mismatch_summary` coverage, and a Sheet 1 conditional-format rule "Add to MAP Roster" -> light blue (`#CCE5FF`), inserted before the NOT_BLANK yellow catch-all.
- **`Code.js`**: accept routing for "Add to MAP Roster" -> `_MapAdditionsData`; `_MapAdditionsData` added to `removeStudentFromCumulativeTabs_`.
- **`generate_corrections.py::read_handled_student_keys`**: `_MapAdditionsData` added so accepted Add-to-MAP students hide from Sheet 1 (v2.7.5 tuple logic).

### Workflow
On Sheet 1 ("Corrected Roster Info") an Add-to-MAP row shows the SIS student data (so the IM sees who to add) with a light-blue Mismatch Summary. The IM accepts -> Apps Script routes to `_MapAdditionsData` -> the student appears on Sheet 7 "Missing from MAP Roster". The IM then adds the student to the MAP roster manually. Reject routes to `_RejectedData` like any other type.

### Direction note
"Roster Addition" = in MAP, not in SIS (data team adds to SIS). "Add to MAP Roster" = in SIS, not in MAP (IMs add to MAP). Opposite directions, distinct label + distinct sheet, so the two are never conflated.

### Verified
- `python -m py_compile config.py generate_corrections.py sheets_writer.py`; `node --check Code.js`. All pass.
- `python generate_corrections.py`: "Add to MAP Roster (in SIS, not in MAP): 23". Both new tabs auto-created (`_MapAdditionsData` hidden, "Missing from MAP Roster" visible).
- Live `_CorrData` probe: 23 Add-to-MAP rows. By campus: JHMS 7, Vita 5, JHES 3, JRHS 3, AASP 2, JRES 2, AFMS 1. Test accounts present: 0. Unmanaged-school rows: 0 (scoping correct). (26 raw SIS-only minus 3 test accounts = 23; Metro dropped out entirely since its only SIS-only student was a test account.)
- Sheet 7 header correct (`Date Approved, Mismatch Summary, Campus, ...`), 0 data rows until IMs accept.
- Sheet 1 conditional-format rules read back via API: 4 rules in correct priority order, "Add to MAP Roster" -> blue (0.80, 0.90, 1.00) ahead of NOT_BLANK -> yellow.
- Existing 3 detection paths untouched (additive 3rd loop; existing loops unchanged).

### Excluded (out of scope)
- Weekly snapshot: `_MapAdditionsData` is NOT in `WEEKLY_SOURCE_TABS` (these go to IMs to add to MAP, not the SIS data team).
- No automated MAP writes; IMs add students manually after review.

### Files changed
- `config.py`, `generate_corrections.py`, `sheets_writer.py`, `Code.js`, `docs/CHANGELOG.md`, `docs/AI_INSTRUCTIONS.md`.

### User action required
- Deploy lands the Apps Script routing automatically (clasp). After the next pipeline run, review the new Sheet 7 candidates on Sheet 1, accept the real ones, and add them to the MAP roster.

## [v2.7.6] - 2026-05-19

### Added
- **`generate_weekly_snapshot.py` `--since YYYY-MM-DD --name "TITLE"` date-range export mode.** Produces a consolidated snapshot of every correction whose Date Approved (col A) is on or after the `--since` date, into a Shared Drive file named `TITLE`. Fills the gap where the existing tooling could only filter by Sent Week status (`--all-unsent`) or current Monday (default). Neither can express "all changes after date X across multiple weeks".
  - `_parse_date_approved(date_str)`: parses col A into a `date`. Handles `%Y-%m-%d %H:%M:%S`, `%m/%d/%Y %H:%M:%S`, `%Y-%m-%d`, `%m/%d/%Y`. Returns `None` for blank/unparseable.
  - `filter_since_date(rows, since_date)`: selects `(row_num, row)` where col A date >= `since_date`. The only filter mode keyed on col A; default + `--all-unsent` key on col O.
  - `main()` extended to `main(all_unsent=False, since_date=None, custom_name=None)`.
  - argparse: `--since` + `--name`. Validation: `--since` requires `--name`; `--since` mutually exclusive with `--all-unsent`; `--name` only valid with `--since`.

### Key behavior
- **`--since` is inclusive** of the given day. The first run used `--since 2026-05-05` to honor the user's "strictly after 5/4" intent.
- **No Sent Week stamping.** The stamping block is wrapped in `if since_date is None:`. `--since` selected rows keep their original per-week Sent Week values. Re-stamping would corrupt per-week history and break future default-mode cron runs. `--since` never writes the cumulative tabs (read-only re-export).
- **No Instructions tab.** Gated on `all_unsent` (stays False). 3 data tabs only.
- **Idempotent.** Re-running with the same `--name` updates the file in place (find-or-create by exact name).

### Why
User asked for "a new corrections sheet for all the changes after 5/4 to present". Probe showed every cumulative-tab row was already stamped sent (Sent Week 04-20 / 05-04 / 05-11 / 05-18), so `--all-unsent` would have returned zero rows; the default mode only bundles the current Monday. A date-range mode was the only way to express the ask.

### Verified
- `python -m py_compile generate_weekly_snapshot.py` passes. `--help` shows the new flags + example.
- First run: `python generate_weekly_snapshot.py --since 2026-05-05 --name "5/19 Corrections"` CREATED file id `1BXCdHyRQhUtL4y4oEYfVD6hz4_uDRynqaBG8AqJ5sqM` in the Weekly Corrections Archive Shared Drive. 3 tabs: Correction List 77, Roster Additions 44, Roster Unenrollments 47 (168 total).
- **col O integrity check**: SHA-256 of col O on all 3 cumulative tabs (`_ApprovedData`, `_AdditionsData`, `_UnenrollData`) byte-identical before and after the run. Confirms no stamping. `--since` is genuinely read-only w.r.t. the cumulative tabs.
- New file structure: 3 tabs at index 0/1/2, no Instructions tab, default `Sheet1` deleted. Date ranges per tab all >= 2026-05-05 (5/4 correctly excluded). 14-col header (`Date Approved`, `Mismatch Summary`, `Campus`, ...).
- Default mode + `--all-unsent` mode untouched. Additive change: both paths bypass the new `since_date` branches.

### Files changed
- `generate_weekly_snapshot.py`: `_parse_date_approved`, `filter_since_date`, `main()` branches, argparse flags, docstring + epilog
- `docs/CHANGELOG.md`, `docs/AI_INSTRUCTIONS.md`: this entry + Date-range export mode section

### User action required
- **None.** The "5/19 Corrections" file is in the Shared Drive. Open it, click Share, send to support. Re-run `python generate_weekly_snapshot.py --since 2026-05-05 --name "5/19 Corrections"` any time to refresh it in place.

## [v2.7.5] - 2026-05-08

### Changed
- **Sheet 1 "Corrected Roster Info" now hides previously-handled student-mismatch pairs FOREVER**, not just within a 7-day window. v2.4.4 introduced `HIDE_HANDLED_DAYS = 7` to hide Accept/Reject'd students temporarily — the rationale was "signal IMs about stale corrections" but in practice IMs were confused by previously-actioned students reappearing every 7 days. Most notable example: John Bradley Apostol's students kept resurfacing, IMs kept re-rejecting them.
- **Re-flag policy**: if a NEW *different* mismatch type arises for a student you've already handled (e.g. fixed Grade last week, now SIS also shows wrong Email → mismatch is now "Grade, Email"), the student WILL reappear on Sheet 1 because the new tuple `(sid, "Grade, Email")` is not in handled_keys. The old `(sid, "Grade")` tuple is still hidden. This catches genuinely new data-quality issues without re-nagging on stale ones.

### How the new logic works
- `generate_corrections.py::read_handled_student_keys(sheets_service)` reads all 4 cumulative tabs (`_ApprovedData`, `_AdditionsData`, `_UnenrollData`, `_RejectedData`) and collects every `(student_id, mismatch_summary)` tuple — no time cutoff.
- `generate_corrections.py::_hide_handled(corrections_map, corrections_sis, handled_keys)` filters by tuple membership: a current correction is hidden iff its `(Student_ID, mismatch_summary)` exists in the set.
- Pre-v2.7.5 only collected `student_id` and matched on that alone (within the 7-day window).

### Removed
- `config.HIDE_HANDLED_DAYS = 7` — the constant is gone. The comment block in `config.py` explains the v2.7.5 semantics for future reference.
- `from datetime import datetime, timedelta` in `generate_corrections.py` — no longer needed since handled_keys ignores date entirely. Drops a class of latent bugs (the unparseable-timestamp silent-skip issue from pre-v2.4.2 era).

### Renamed
- `read_handled_student_ids` → `read_handled_student_keys` (returns tuples now, not bare sids).
- `_hide_recently_handled` → `_hide_handled` (no longer "recently"; "previously" or "ever").

### Verified
- `python -m py_compile generate_corrections.py config.py` → OK.
- Local pipeline run: **Hidden 153 previously-handled student-mismatch pair(s) (235 total handled keys). 5 corrections remain on Sheet 1.** (Prior v2.7.4 run with 7-day window: 200 hidden, 30 remaining.)
- Spot-check of the 5 remaining: 3 are previously-Reject'd JHES students whose mismatch CHANGED ("Guide Email" → "Guide Email, Guide Name" — additional Guide Name mismatch surfaced after their original rejection). Working as designed: the new mismatch type re-surfaces them. The other 2 are genuinely new (one JHMS Student Group mismatch, one Vita Unenrolling) — neither has ever been handled.
- Apostol-batch / Lemuel Mosquito / JHES Guide Email-only students from the 2026-05-04 batch: confirmed absent from Sheet 1.

### Cumulative tab safety
The 4 approval sheets (Sheets 3-6) are unchanged. Every handled row remains on its corresponding approval sheet forever. The data team's job board is still the approval sheets. v2.7.5 only changes when students re-appear on the input sheet (Sheet 1).

### Data-loss audit (post-deploy)
After the 50+ → 5 transition on Sheet 1 raised the obvious "did we lose data?" question, ran a comprehensive read of every tab:

| Tab | Row count | Note |
|---|---|---|
| `_ApprovedData` | 86 | preserved |
| `_AdditionsData` | 44 | preserved |
| `_UnenrollData` | 48 | preserved |
| `_RejectedData` | 59 | preserved |
| `_RejectionReasons` | 11 | preserved |
| `_CorrData` | 5 | only the new-mismatch/never-handled students |
| Sheet 3 "Automated Correction List" | 86 | matches `_ApprovedData` |
| Sheet 4 "Roster Additions" | 44 | matches `_AdditionsData` |
| Sheet 5 "Roster Unenrollments" | 48 | matches `_UnenrollData` |
| Sheet 6 "Rejected Changes" | 59 | matches `_RejectedData` |

**Total handled tuples**: 235. **Distinct student_ids**: 235. All sampled tuples found in handled_keys. The 5 visible students on Sheet 1 all have either a NEW mismatch_summary value (different from what they were previously actioned for) or have never been handled — verified by cross-referencing each sid against `sid_to_mismatches` map.

**Conclusion: NO DATA LOSS.** Every previously-actioned correction is preserved on the corresponding approval sheet (3/4/5/6). Reasons preserved in `_RejectionReasons`. The Sheet 1 transition is purely the intentional v2.7.5 hide-forever behavior.

### Minor finding: 2 cross-tab duplicate-tuple sids (cosmetic only)
- `079-10545` Bader Sammoud — appears in BOTH `_AdditionsData` (Accept'd 2026-05-07 12:07:32) AND `_RejectedData` (Reject'd 10 min later, 12:17:31) with the same `Roster Addition` mismatch. Should have been deduped by `Code.js::removeStudentFromCumulativeTabs_` on toggle. Likely cause: race condition or pre-v2.6.0 manual-paste-era Apps Script ran without the dedup.
- `083-11566` — TWO entries in `_UnenrollData` with same `Unenrolling` mismatch, dates 2026-05-04 and 2026-05-13. Same root cause.

Functional impact: NONE. handled_keys is a `set` of tuples — duplicates collapse. Both sids are correctly hidden from Sheet 1. The duplicate rows are cosmetic cruft on the approval sheets, will appear as a duplicate visual row on Sheet 4/5/6 respectively. Optional cleanup: delete the older row manually via Apps Script editor. Not auto-fixing — too risky to write a dedup pass when there's no functional impact.

### Files changed
- `generate_corrections.py` — `read_handled_student_keys` + `_hide_handled` rewrites + import cleanup
- `config.py` — removed `HIDE_HANDLED_DAYS`, added v2.7.5 comment block
- `docs/CHANGELOG.md` — this entry
- `docs/AI_INSTRUCTIONS.md` — Row-Hiding section + Key Design Decisions #10 updated
- `docs/HUMAN_INSTRUCTIONS.md` — "student re-appearing" troubleshooting row updated

### User action required
- **None.** Next hourly cron picks up the new code.
- **To force a previously-handled student back onto Sheet 1**: open the relevant cumulative tab via Apps Script editor (the hidden `_ApprovedData` / `_AdditionsData` / `_UnenrollData` / `_RejectedData`) and delete the row for that student. The next pipeline run will re-surface them.

## [v2.7.4] - 2026-05-08

### Fixed
- **Reason for Rejection STILL disappearing after v2.7.3 architecture.** Two compounding issues:
  1. **OPERATIONAL**: v2.7.3 was implemented but never committed/pushed. Origin/main stayed at `63e7a09` (v2.7.2). Hourly cron continued running pre-v2.7.3 code that wiped Sheet 6 A:Z each run. Lesson: a fix that doesn't deploy is the same as no fix.
  2. **ARCHITECTURAL**: even with v2.7.3 deployed, storing reasons in `_RejectedData` col O leaves them exposed to 3 destructive paths in `_RejectedData`'s ecosystem:
     - `sheets_writer.py::_realign_row` truncates rows to 14 cols (line 626, 646)
     - `sheets_writer.py::_backfill_mismatch_summary` clears `_RejectedData` A:Z then rewrites with 14-col rows (line 727-740)
     - `Code.js::removeStudentFromCumulativeTabs_` deletes `_RejectedData` row on Reject toggle (line 152-169) — uncheck-and-recheck silently wipes the reason

### Added — separate-tab architecture
- **NEW hidden tab `_RejectionReasons`** (2 cols: `student_id`, `reason`). Decoupled from the 4 cumulative tabs (`_ApprovedData`, `_AdditionsData`, `_UnenrollData`, `_RejectedData`). No existing rebuild path touches `_RejectionReasons`. Created as hidden on first run.
- **`Code.js::upsertRejectionReason_`** — find-or-append pattern for `_RejectionReasons`. Linear scan over col A by student_id. If found, update col B. If not, append. Mirrors the read pattern of `removeStudentFromCumulativeTabs_` but never deletes.
- **`Code.js::handleRejectionReasonEdit_`** rewritten — calls `upsertRejectionReason_` instead of writing to `_RejectedData` col O.
- **`sheets_writer.py::_hydrate_rejection_reasons`** rewritten — reads from `'_RejectionReasons'!A:B` instead of `'_RejectedData'!M:O`.
- **`sheets_writer.py::write_corrections::all_tab_names`** — added `"_RejectionReasons"` so `_ensure_all_tabs` creates it idempotently on first run.
- **`restore_rejection_reasons.py`** — fully rewritten for `_RejectionReasons` upsert semantics. Reads pre-wipe XLSX, ensures `_RejectionReasons` exists (hidden), and upserts (sid, reason) pairs. `--force` overwrites existing non-blanks; default skips them.

### Bullet-proof guarantees vs prior versions

| Failure mode | v2.7.2 | v2.7.3 | v2.7.4 |
|---|---|---|---|
| Pipeline `batchClear` wipes Sheet 6 A:Z | YES | mitigated (A:N) | mitigated (A:N) |
| `_realign_row` truncates `_RejectedData` to 14 cols | N/A | RISK | N/A — col O isn't on `_RejectedData` |
| `_backfill_mismatch_summary` clears `_RejectedData` A:Z | N/A | RISK | N/A |
| User toggles Reject (uncheck → recheck) wipes reason | YES (no protection) | RISK (`removeStudent…` deletes row) | RESOLVED — `_RejectionReasons` not touched |
| Pipeline never deployed | RISK | RISK | RISK (operational only) |

The only remaining loss path is "user explicitly clears the cell on Sheet 6 col O" → onEdit fires with empty value → upsert writes empty string. That is intentional user action, not a system bug.

### Restoration steps (this incident)
1. Confirmed: `_RejectedData` col O still had all 11 reasons from yesterday's restore (cumulative tab; not wiped).
2. Migrated those 11 reasons into the new `_RejectionReasons` tab (one-shot inline script).
3. Local pipeline run with v2.7.4 code → hydration count `11/39` ✓ — all reasons appeared at correct student rows on Sheet 6.
4. Deploy v2.7.4 (commit + push). GHA auto-deploys Code.js via clasp. Hourly cron picks up new sheets_writer.py on next run.

### Deploy verification
After push, confirmed:
- `gh run list --workflow=deploy-apps-script.yml` → most recent run = SUCCESS, headSha matches new commit
- `gh run list --workflow=hourly-pipeline.yml` (next :00 run) → headSha matches new commit, Sheet 6 col O still populated

### Files changed
- `sheets_writer.py` — `_ensure_all_tabs` adds `_RejectionReasons`; `_hydrate_rejection_reasons` reads from new tab
- `Code.js` — `upsertRejectionReason_` (new); `handleRejectionReasonEdit_` rewritten
- `restore_rejection_reasons.py` — fully rewritten for v2.7.4 semantics
- `docs/CHANGELOG.md` — this entry
- `docs/AI_INSTRUCTIONS.md` — Sheet 6 architecture + hidden-tabs list updated

### User action required
- **None.** Reasons restored. Future Reject toggles + pipeline rebuilds preserve them automatically.

## [v2.7.3] - 2026-05-07

### Fixed
- **"Reason for Rejection" column (Sheet 6 col O) was being silently wiped on every hourly pipeline run.** `sheets_writer.py::write_corrections` cleared `'Rejected Changes'!A:Z` via `batchClear`, which included col O. Sheet 6's QUERY pulls A:N (14 cols) from `_RejectedData`; col O is a 15th, manually-entered column. Pre-v2.7.3, col O had no backing data anywhere — every IM-typed reason vanished on the next hourly cron. Bug present since v2.1.0 (when Sheet 6 was first introduced).

### Why a stop-the-wipe-only fix would not have worked
Sheet 6's QUERY uses `SORT()`. New rejected rows appear at the top, pushing existing rows down. A reason typed at row 7 today is for a different student tomorrow if any new rejection arrives between runs. Reasons must be tied to `student_id` (stable), not row position (unstable).

### Added
- **`_RejectedData` schema extended from 14 → 15 cols.** Col O = persistent storage for "Reason for Rejection", keyed by `student_id` in col M. Existing rows have blank col O (populated by the restore script for historical rows). New rejections have blank col O (populated when the user types a reason).
- **`Code.js::handleRejectionReasonEdit_`** — Apps Script `onEdit` branch for Sheet 6 col O. Reads student_id from col M same row, walks `_RejectedData` col M to find the matching student, writes the reason to that row's col O. Mirrors the existing `removeStudentFromCumulativeTabs_` pattern.
- **`sheets_writer.py::_hydrate_rejection_reasons`** — Called near end of `write_corrections`. Reads Sheet 6 col M (rows 7+, the QUERY-rendered student_ids) and `_RejectedData` cols M+O. Writes Sheet 6 col O aligned to current row order. Reasons follow students through QUERY reorders.
- **`restore_rejection_reasons.py`** (~180 lines) — One-time restoration. Reads a pre-wipe XLSX export of the corrections spreadsheet (downloaded by user from File → Version history → Download as .xlsx), extracts (student_id, reason) pairs from Sheet 6, writes them to `_RejectedData` col O for matching students. `--force` flag overwrites existing non-blank reasons; default skips them.
- **`requirements.txt`**: added `openpyxl` (used by `restore_rejection_reasons.py`).

### Changed
- **`sheets_writer.py::write_corrections`** — Replaced the `[f"'{t}'!A:Z" for t in clear_tabs]` list-comprehension with a per-tab `clear_ranges` dict. Sheet 6 cleared as `A:N` (preserves col O). Other tabs unchanged at `A:Z`.

### Restoration steps (one-time, user-driven)
1. Open the corrections spreadsheet → File → Version history → See version history.
2. Find a revision from before the wipe (any timestamp where col O on Sheet 6 still has the typed reasons — typically several days back).
3. Click the ⋮ menu on that revision → Make a copy → name it e.g. `pre-wipe-snapshot`.
4. Open that copy → File → Download → Microsoft Excel (.xlsx).
5. Run: `python restore_rejection_reasons.py path/to/downloaded.xlsx`
6. Verify the report shows the expected number of restored reasons.
7. Delete the snapshot copy when done.

### Architecture (in one diagram)
```
USER TYPES reason in Sheet 6 col O at row R
        ↓
Code.js onEdit → handleRejectionReasonEdit_ fires
        ↓
Reads student_id from Sheet 6 col M row R
Walks _RejectedData col M → writes reason to that row's col O
        ↓
NEXT PIPELINE RUN:
  sheets_writer.py clears Sheet 6 A:N (NOT A:Z — col O preserved-then-rewritten)
  QUERY rebuilds Sheet 6 A:N from _RejectedData A:N
  _hydrate_rejection_reasons reads Sheet 6 col M + _RejectedData M+O
  Writes Sheet 6 col O matched by student_id
        ↓
Reason appears at the CORRECT row regardless of QUERY sort order
```

### Verified
- `python -m py_compile sheets_writer.py restore_rejection_reasons.py config.py generate_corrections.py` → all OK.
- `node --check Code.js` → OK.
- (Live verification pending — runs after user provides pre-wipe XLSX export.)

### User action required
1. Restore lost reasons via the restoration steps above.
2. Confirm next hourly cron run still preserves col O after restoration (passive — no action needed; just check the sheet next morning).
3. To verify the fix works for new entries: type a reason in Sheet 6 col O for any rejected student, manually trigger the pipeline (`python generate_corrections.py` or trigger the GHA workflow), and confirm the reason still appears at the correct row after rebuild.

### Files changed
- `sheets_writer.py` — narrow Sheet 6 wipe + `_hydrate_rejection_reasons` (+72 -16)
- `Code.js` — `handleRejectionReasonEdit_` + onEdit dispatch (+58 -1)
- `restore_rejection_reasons.py` — NEW (~180 lines)
- `requirements.txt` — `openpyxl` (+1)
- `docs/CHANGELOG.md`, `docs/AI_INSTRUCTIONS.md` — this entry + fix Sheet 6 docs

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
