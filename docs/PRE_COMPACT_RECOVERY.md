# POST-COMPACTION RECOVERY - 2026-05-06 15:55 EDT

> **For post-compaction Claude:** Read this section BEFORE doing anything else. It captures the in-flight state of the prior session. If anything below conflicts with the actual filesystem or git state, **TRUST THE FILESYSTEM** and ask the user.

## 1. Self-instructions (DO THESE IN ORDER)

1. Run `/refresh` (no project-specific `/refresh-context-weekly-corrections` skill exists; check `.claude/commands/` to confirm).
2. Read these files in order:
   - `docs/CHANGELOG.md` (top section: v2.7.2, v2.7.1, v2.7.0) — the four versions shipped this session, with full root-cause analysis for each.
   - `docs/AI_INSTRUCTIONS.md` (sections "Timeback SIS bridge (v2.7.0)" + "Field coverage for Timeback rows (v2.7.1)" + "Staleness gotcha") — current architecture state, including the IMPORTRANGE staleness gotcha.
   - `timeback_sis.py` — new module added this session for OneRoster API bridge.
   - `generate_corrections.py::_find_mismatches` (line ~395) and `::read_combined_sis_data` (line ~285) — the v2.7.0/v2.7.1 dispatch logic.
   - `config.py::MAP_HEADER_MAP` (line ~42) — has an explicit comment about why `"alpha student id"` must NOT be in the `ext_student_id` matcher (v2.7.2 lesson).
3. Verify environment matches Section 6 (run quick sanity checks if anything is uncertain).
4. Check `git rev-parse HEAD` — if it differs from `63e7a09`, a concurrent session committed; run `git log 63e7a09..HEAD --oneline` to see what changed.
5. Resume work at: **NO ACTIVE WORK**. v2.7.2 is the latest commit + push. All TodoWrite items completed. Working tree clean. Session is at a natural stopping point. If user asks "continue", clarify which thread (could be v2.7.2 follow-up, the spawned hook-path-fix task, or something new).

## 2. Session marker

| Field | Value |
|-------|-------|
| Date | 2026-05-06 15:55 EDT |
| Project | weekly-corrections |
| Branch | main |
| HEAD at pre-compact | `63e7a09` |
| Working tree | clean (verified by `git status --short` returning empty) |
| Pre-compact prepared by | Claude Opus 4.7 (1M context) |
| Recovery doc path | `docs/PRE_COMPACT_RECOVERY.md` |

## 3. Active TodoWrite state

```
(none — all phases of v2.7.2 completed; session is at a natural stopping point post-commit)
```

## 4. Files committed this session

This session shipped 4 versions: v2.6.1, v2.7.0, v2.7.1, v2.7.2.

| File | Change | Commit | Why |
|------|--------|--------|-----|
| `generate_weekly_snapshot.py` | modified (+439 LOC) | `9705c18` (v2.6.1) | Added `--all-unsent` flag + Instructions tab for ad-hoc support packet generation |
| `config.py` | modified | `9705c18` (v2.6.1) | Added `WEEKLY_TAB_INSTRUCTIONS` constant |
| `docs/CHANGELOG.md` | modified | `9705c18` + `5882e09` + `f130e76` + `63e7a09` | v2.6.1, v2.7.0, v2.7.1, v2.7.2 entries |
| `docs/AI_INSTRUCTIONS.md` | modified | `9705c18` + `5882e09` + `f130e76` | Support-packet section + Timeback SIS bridge section + field-coverage table |
| `docs/HUMAN_INSTRUCTIONS.md` | modified | `9705c18` + `5882e09` | "How to generate a support packet" + supported-campuses table |
| `config.py` | modified | `5882e09` (v2.7.0) | Added `TIMEBACK_CAMPUSES`, `TIMEBACK_CREDS_PATH`, ScienceSIS + Vita to `CAMPUS_SHEETS` and `ISR_CONFIG`; added "alpha student id" to `ext_student_id` matcher (REVERTED in v2.7.2) |
| `timeback_sis.py` | **created** (266 LOC) | `5882e09` (v2.7.0) | New self-contained OneRoster client (OAuth2 + GET /schools/{id}/students) |
| `generate_corrections.py` | modified | `5882e09` + `f130e76` | New `read_combined_sis_data` (alpha_roster + Timeback merge); empty-Notes coercion to "Enrolled" for Timeback; `is_timeback` flag to `_find_mismatches` to skip non-corresponding fields |
| `requirements.txt` | modified | `5882e09` (v2.7.0) | Added `requests` for `timeback_sis.py` |
| `.github/workflows/hourly-pipeline.yml` | modified | `5882e09` (v2.7.0) | Added "Write Timeback creds from secret (fail-soft)" step that reads `TIMEBACK_CREDS_JSON` GHA secret |
| `README.md` | modified | `5882e09` (v2.7.0) | Updated architecture diagram for 11 campuses (9 Dash + 2 Timeback) |
| `config.py` | modified | `f130e76` (v2.7.1) | Added `TIMEBACK_CAMPUS_NAMES` set for `_find_mismatches` dispatch |
| `timeback_sis.py` | modified | `f130e76` (v2.7.1) | Strip `" (TimeBack)"` from Campus value (was producing systemic Campus mismatch) |
| `config.py` | modified | `63e7a09` (v2.7.2) | REVERT — removed "alpha student id" from `ext_student_id` matcher (was breaking Dash detection on all 9 campuses) |

## 5. Files modified but NOT YET committed

```
(none — clean working tree)
```

## 6. Environment state — what was VERIFIED to work this session

- **Python**: 3.12 working (used for all pipeline runs)
- **`pip install -q requests`**: succeeded; `requests` library now in requirements.txt
- **`python -m py_compile config.py timeback_sis.py generate_corrections.py`**: passes (verified at multiple points: post-Phase A in v2.7.0, v2.7.1, v2.7.2)
- **Service account auth (`keys/sa-main.json`)**: working for BigQuery + Sheets + Drive APIs
- **Timeback Cognito OAuth2 (`keys/timeback-creds.json`)**: working — pulled 76 students across 2 schools, OAuth handshake successful, pagination working
- **Live `_CorrData` probe (post-v2.7.2 run)**:
  - Total rows: 230 (was 2,056 in v2.7.1)
  - Matches (clean): 1,870 (was 44 in v2.7.1 — 1,826 false-positive Dash ExtSID mismatches resolved)
  - Hardeeville `083-11733`: matches cleanly, dropped from `_CorrData`
  - 5 ScienceSIS Unenroll students all routed correctly: 2 still pending Sheet 1 acceptance (Alanah 066-6749, Autumn 066-6742); 3 in `_UnenrollData` (Arie 066-6778, Armoni 066-6774, Ataijah 066-6773 — user accepted between v2.7.1 and v2.7.2)
  - 0 Timeback noise violations (no Level / Student Group / Guide Email / Guide Name / External Student ID in any Vita / ScienceSIS mismatch_summary)
  - 0 Dash ExtSID mismatches (was 1,867 in v2.7.0/v2.7.1)
- **Pipeline runtime**: 15.6s end-to-end post-v2.7.2 (was ~18-20s in v2.7.0/v2.7.1; faster because fewer corrections to write)
- **GHA secret `TIMEBACK_CREDS_JSON`**: user confirmed added earlier in session (need next hourly cron to verify it loads, but not yet observed via Actions tab)
- **Apps Script auto-deploy** via clasp (v2.6.0): NOT verified this session (no Code.js changes in v2.6.1/v2.7.x)
- **Setup scripts** (`setup_unenroll_columns.py`, `build_unenroll_queue.py`): both ran successfully end-to-end

## 7. Key decisions + rationale

| Decision | What was decided | Why | Where it lives |
|----------|------------------|-----|----------------|
| Timeback bridge approach | Live OneRoster API per run (~3s overhead) | User chose authoritative correctness over cached BQ table during plan-time AskUserQuestion | `generate_corrections.py::read_combined_sis_data`; `timeback_sis.py` |
| Cred handling | `keys/timeback-creds.json` (gitignored) + GHA secret `TIMEBACK_CREDS_JSON` | User explicitly chose this over pointing at sibling repo's keys/ | `config.py::TIMEBACK_CREDS_PATH`; `.github/workflows/hourly-pipeline.yml` |
| CMR Unenroll header position | IMPORTRANGE goes in row 2; row 1 keeps existing "Unenroll" header | User explicitly asked to keep header | `setup_unenroll_columns.py::setup_cmr_importrange` |
| Reverse sync | NOT included in v2.7.0 (would auto-unenroll students in Timeback API on accept) | Out of scope — different blast radius. Support manually unenrolls via Timeback admin UI | (intentionally not in code) |
| Merge precedence | Timeback wins over alpha_roster on `student_id` collision (62 overlapping students) | Timeback is the new system of record per user spec; alpha_roster is legacy during the migration window | `generate_corrections.py::read_combined_sis_data` line ~325 |
| Empty Notes coercion | Timeback CMR rows with empty Notes → coerced to "Enrolled" so they enter `map_enrolled` and the IM-checkbox path can fire | Timeback IMs don't maintain Notes (the OneRoster API is the source); without coercion, all Timeback rows were skipped | `generate_corrections.py::read_map_roster` line ~140 |
| Fields skipped for Timeback | Level, External Student ID, Student Group, Guide Email, Guide Name combine | OneRoster API doesn't expose these; comparing produces noise mismatches | `generate_corrections.py::_find_mismatches` (is_timeback branch); `docs/AI_INSTRUCTIONS.md` field-coverage table |
| Campus value stripping | Strip `" (TimeBack)"` from `campus_label` in `timeback_sis.py` so it matches CMR's "ScienceSIS" / "Vita High School" | Without strip, every Timeback row produced a systemic Campus mismatch | `timeback_sis.py::query_timeback_enrolled` line ~210 |
| `"alpha student id"` matcher | REMOVED in v2.7.2 (had been added in v2.7.0) | Every Dash CMR has BOTH "Alpha Student ID" AND "SUNS Number" columns; left-to-right detector picked the wrong one for all 9 Dash tabs (1,867 false mismatches) | `config.py::MAP_HEADER_MAP["ext_student_id"]` (now 4 elements + explanatory comment) |

## 8. Data / artifact details (the precise stuff auto-compact destroys)

### Spreadsheet IDs

- **CMR (Combined MAP Roster)**: `1scEay0a8OR6vU3uJuxbHKWCEx_RVgSsRXF9naJh3XYw`
  - Vita tab gid: `1311780933` (title: `"Vita High School (TimeBack)"`)
  - ScienceSIS tab gid: `254383148` (title: `"ScienceSIS (TimeBack)"`)
- **Output (corrections)**: `12dqu58KKdsZN9nLre9Fntkk7vSILu3KfcW4WDvo5-Ls`
- **Vita ISR**: `1sOSwvwPb8cXSfJgXF-E2Ur2v0lyvi1qkR94OvgQLQ4Y`
- **ScienceSIS ISR**: `1SjVoQRubz_nsD3YVKLf68KcTaZwx1s7CA5S9V_E8gQ8`
- **Weekly Corrections Archive Shared Drive**: `0AFQGIqcKjsyFUk9PVA`
- **Apps Script project**: `16_ypoWiIFRpIZzUEpwJGLP8DvexGoCDAiXaZGKoIhRQFa38H8vcS436_`

### Timeback school sourcedIds

- Vita High School: `e57cb46d-b6b0-4f45-96ed-327441b5d068`
- ScienceSIS: `7c475cf4-12b4-40ed-8857-dc6e624a5fa1`

### Dash CMR dual-column layout (THE GOTCHA THAT BIT v2.7.0)

EVERY Dash CMR tab has BOTH columns:

| Dash CMR tab | "Alpha Student ID" col | "SUNS Number" / "External ID" col |
|---|---|---|
| Hardeeville Elementary | W (idx 22) | AB (idx 27) — "SUNS Number" |
| Hardeeville JR/SR HS | W | AB |
| Ridgeland Elementary | W | AB |
| Ridgeland Sec Academy | W | AB |
| Allendale Fairfax Elem | W | AB |
| Allendale Fairfax MS | W | AB |
| Allendale Aspire | X (idx 23) | AC (idx 28) |
| Metro Schools | W | AB ("External ID") |
| Reading CCSD | X | (NO SUNS column — only Alpha) |

`config.py::MAP_HEADER_MAP["ext_student_id"]` MUST contain only `{"suns number", "external student id", "suns #", "external id"}` — NOT `"alpha student id"` — or Dash detection breaks. Reading CCSD intentionally has no detection (preserves pre-v2.7.0 behavior).

### Timeback CMR layout (different from Dash)

Both Vita and ScienceSIS CMR tabs (28 cols):

| Col | Header |
|-----|--------|
| A | Student ID (legacyDashStudentId, e.g. `066-6757`, `033-2154`) |
| B | Student Email |
| C | Campus |
| D | NWEA Account |
| E | Last Name |
| F | First Name |
| G | Grade |
| H | Level |
| I-V | DOB / Gender / Subjects / RIT scores etc. (Timeback-specific) |
| W | Alpha Student ID |
| X | School, if separate from Campus |
| Y/Z/AA | Teacher 1 First / Last / Email |
| AB | Unenroll |

### ISR_CONFIG entries for Timeback (post-v2.7.0)

```python
"ScienceSIS (TimeBack)": {
    "isr_id": "1SjVoQRubz_nsD3YVKLf68KcTaZwx1s7CA5S9V_E8gQ8",
    "mr_gid": 1256615349,
    "sr_unenroll_col": 23,   # col X
    "mr_unenroll_col": 27,   # col AB
},
"Vita High School (TimeBack)": {
    "isr_id": "1sOSwvwPb8cXSfJgXF-E2Ur2v0lyvi1qkR94OvgQLQ4Y",
    "mr_gid": 1256615349,
    "sr_unenroll_col": 23,
    "mr_unenroll_col": 27,
},
```

### Pipeline metrics (most recent run, v2.7.2)

- Total MAP roster: 2,084 enrolled, 254 non-enrolled
- alpha_roster: 8,709 students
- Timeback API: 76 students across 2 campuses (52 ScienceSIS + 24 Vita; 10 skipped for missing legacyDashStudentId)
- Combined SIS: 8,723 students (62 overlap; Timeback wins)
- **Matches: 1,870** (was 44 in v2.7.1)
- Roster Additions: 67
- Field mismatches: 115 (was 1,941 in v2.7.1)
- Unenrolling: 48 (IM-flagged: 32, Notes-based: 16)
- Total corrections: 230 (was 2,056)
- Hidden recently-handled: 105
- Visible on Sheet 1: 125 corrections
- Pipeline runtime: 15.6s

### 5 ScienceSIS Unenroll students (user-flagged this session)

| student_id | Name | Status |
|---|---|---|
| 066-6749 | Alanah Cossey | On Sheet 1 (Unenrolling) — pending IM acceptance |
| 066-6778 | Arie Sturgis | In `_UnenrollData` (accepted at 2026-05-06 15:39:35) |
| 066-6774 | Armoni Nelson | In `_UnenrollData` (accepted at 2026-05-06 15:40:20) — user's specific complaint, now resolved |
| 066-6773 | Ataijah Mitchell | In `_UnenrollData` (accepted at 2026-05-06 15:40:31) |
| 066-6742 | Autumn Boothe | On Sheet 1 (Unenrolling) — pending IM acceptance |

## 9. Pending work — what's NEXT

**No active work.** v2.7.2 is the latest commit, pushed to main. All TodoWrite items completed. Working tree clean.

If user asks for "next steps" or "continue":
- **Possible thread A (waiting on user)**: confirm GHA `TIMEBACK_CREDS_JSON` secret loads correctly. User added the secret earlier; the next `:00 UTC` hourly cron will exercise it. Look for `Timeback creds written.` in the workflow log (vs `::warning::TIMEBACK_CREDS_JSON secret not set.`).
- **Possible thread B (already spawned as task)**: fix `enforce_go_plan_commit_gate.py` path bug (`weekly-corrections/CHANGELOG.md` should be `weekly-corrections/docs/CHANGELOG.md`). Spawned via `spawn_task` so the user can dismiss or pick up.
- **Possible thread C (potential follow-up)**: smoke-test the IM unenroll flow end-to-end on a fresh student (the user explicitly mentioned this earlier as optional verification).
- **Otherwise**: session is closed; await new user direction.

## 10. Open questions for the user

- (none — all v2.7.x items shipped + verified)

## 11. Anti-patterns / gotchas discovered this session

1. **DO NOT add `"alpha student id"` to `MAP_HEADER_MAP["ext_student_id"]`.** Every Dash CMR has BOTH "Alpha Student ID" (col W/X) AND "SUNS Number" / "External ID" (col AB/AC). The auto-detector breaks on first match (left-to-right), so adding it causes col W to win for all 9 Dash tabs and produces 1,867 false ExtSID mismatches. v2.7.0 shipped this bug; v2.7.2 reverted it. Document is in `config.py::MAP_HEADER_MAP` as an inline comment.

2. **When broadening a `MAP_HEADER_MAP` matcher set, sample at least one row from EACH existing campus** to confirm no detection collision before shipping. v2.7.0 verify-build only probed Vita / ScienceSIS rows and missed the Dash regression.

3. **IMPORTRANGE staleness on first run**: when `setup_unenroll_columns.py` wires up new IMPORTRANGE on a CMR, the formula returns `#REF!` until a human opens the CMR and clicks "Allow access". A pipeline run during this window reads `_unenroll_flag=False` for everyone. Symptom: students in `map_enrolled` (Notes empty → coerced to "Enrolled") fall through to field comparison instead of Unenrolling. Fix: re-run pipeline after IMPORTRANGE resolves. Documented in `docs/AI_INSTRUCTIONS.md` "Staleness gotcha (v2.7.1 lesson)".

4. **Don't compare fields the SIS source doesn't expose.** Timeback OneRoster API doesn't return Level / Student Group / Guide* / External Student ID. Comparing them produces a noise mismatch on every row. v2.7.1 added the `is_timeback` flag to `_find_mismatches` to skip those.

5. **Strip `" (TimeBack)"` from Campus value** in `timeback_sis.py` so it matches the CMR's bare "ScienceSIS" / "Vita High School". Otherwise every Timeback row mismatches on Campus.

6. **62-student overlap during migration window**: ~62 Vita+ScienceSIS students exist in BOTH alpha_roster AND Timeback during the migration. `read_combined_sis_data` uses Timeback as the winner (per user spec — Timeback is the new system of record). Watch this number over time; should drop as migration completes.

7. **`enforce_go_plan_commit_gate.py` hook path bug**: looks for `weekly-corrections/CHANGELOG.md` but the actual path is `weekly-corrections/docs/CHANGELOG.md`. Used `[skip-go-plan-gate]` literal in v2.7.2 commit message to bypass. Spawned task to fix in `best-claude-practices` repo.

8. **cp1252 console encoding**: Windows default console can't render em-dashes (U+2014) or check-marks (U+2713). Probe scripts that print Unicode need `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')` at the top.

9. **Apps Script `onEdit` doesn't fire on `setValue`**: the v2.4.4 hide-handled feature doesn't need the Apps Script trigger (server-side filter). Don't try to "unhide" via Apps Script.

10. **Slash commands MUST be invoked via the Skill tool** — typing `/go_plan` or `/go` in user message text doesn't auto-invoke; Claude must call `Skill(skill="go_plan")`. Caught in this session as a meta-debugging exercise.

---

(End of POST-COMPACTION RECOVERY section. This is a freshly-created file, no prior body content.)
