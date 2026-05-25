"""generate_corrections.py — Compare MAP roster against SIS pipeline data
and write mismatched students to the Automated Weekly Corrections sheet.

Usage:
    python generate_corrections.py
"""

import sys
import time

from google.cloud import bigquery
from google.oauth2 import service_account
from googleapiclient.discovery import build

from config import (
    SERVICE_ACCOUNT_KEY,
    BQ_PROJECT,
    BQ_DATASET,
    BQ_TABLE,
    SCOPES,
    MAP_SPREADSHEET_ID,
    OUTPUT_SPREADSHEET_ID,
    CAMPUS_SHEETS,
    MAP_HEADER_MAP,
    OUTPUT_FIELDS,
    TIMEBACK_CAMPUSES,  # v2.7.0
    TIMEBACK_CAMPUS_NAMES,  # v2.7.1
)
from queries import query_alpha_roster
from retry_helper import retry_api  # v2.5.2: shared exponential-backoff retry
from sheets_writer import write_corrections
from timeback_sis import query_timeback_enrolled  # v2.7.0


def print_step(msg):
    print(f"\n{'='*60}\n  {msg}\n{'='*60}")


# ── MAP Roster Reader ──────────────────────────────────────────────────────


def _detect_columns(header_row):
    """Auto-detect column indices from header row using MAP_HEADER_MAP.

    Returns a dict mapping our internal field names to column indices (0-based),
    or None for fields not found in this sheet's headers.
    """
    col_map = {}
    normalized_headers = [str(h).strip().lower() for h in header_row]

    for field_name, possible_headers in MAP_HEADER_MAP.items():
        for idx, header in enumerate(normalized_headers):
            if header in possible_headers:
                col_map[field_name] = idx
                break
        # field stays absent from col_map if not found

    return col_map


def read_map_roster(sheets_service):
    """Read students from all MAP roster campus sheets.

    Auto-detects column layout per sheet from row 1 headers, so sheets with
    different schemas (e.g. Reading CCSD has extra 'Full Name' column) are
    handled automatically.

    Returns:
        (enrolled, non_enrolled, emailonly, all_emails):
        - enrolled / non_enrolled: dicts keyed by student_id (Notes == "Enrolled"
          vs any other value).
        - emailonly (v2.8.4): list of records with a blank Student ID but a valid
          email, kept for email-fallback matching in compare_students.
        - all_emails (v2.8.4): set of every lowercased MAP email (id rows +
          emailonly rows), used to suppress false "Add to MAP Roster".
    """
    print_step("1. READING MAP ROSTER")
    students_enrolled = {}
    students_non_enrolled = {}
    map_emailonly = []
    all_map_emails = set()
    skipped_sheets = []

    for sheet_name in CAMPUS_SHEETS:
        # Read entire sheet including header row.
        # v2.5.2: wrapped in retry_api so transient 5xx/Timeout doesn't drop
        # a whole campus for the run. Permanent errors (404 missing tab, etc.)
        # still raise and get caught by the except, which logs and skips.
        range_str = f"'{sheet_name}'!A1:AE"
        try:
            resp = retry_api(
                lambda: sheets_service.spreadsheets()
                .values()
                .get(
                    spreadsheetId=MAP_SPREADSHEET_ID,
                    range=range_str,
                )
                .execute(),
                label=f"read MAP roster '{sheet_name}'",
            )
        except Exception as e:
            print(f"  WARNING: Could not read '{sheet_name}': {e}")
            skipped_sheets.append(sheet_name)
            continue

        rows = resp.get("values", [])
        if not rows:
            print(f"  {sheet_name}: empty sheet")
            continue

        # Detect columns from header row
        header_row = rows[0]
        col_map = _detect_columns(header_row)
        data_rows = rows[1:]  # skip header by default

        # Fallback: if student_id header is missing but most other columns detected,
        # assume column A is student_id (corrupt header cell, e.g. "4" instead of "Student ID")
        required = ["student_id", "notes"]
        missing = [f for f in required if f not in col_map]
        if "student_id" in missing and "notes" in col_map and len(col_map) >= 5:
            print(
                f"  NOTE: '{sheet_name}' col A header is '{header_row[0]}', assuming Student ID"
            )
            col_map["student_id"] = 0
            missing = [f for f in required if f not in col_map]

        if missing:
            print(f"  WARNING: '{sheet_name}' missing required headers: {missing}")
            print(f"           Found headers: {header_row[:15]}...")
            skipped_sheets.append(sheet_name)
            continue

        notes_col_letter = chr(65 + col_map["notes"]) if col_map["notes"] < 26 else "?"
        num_cols = len(header_row)

        enrolled_count = 0
        non_enrolled_count = 0
        for row in data_rows:
            # Pad row
            while len(row) < num_cols:
                row.append("")

            notes = _safe_get(row, col_map.get("notes")).strip()
            student_id = _safe_get(row, col_map.get("student_id")).strip()
            email = _safe_get(row, col_map.get("email")).strip()

            # v2.8.4: blank Student ID. Previously skipped outright. If the row has
            # an email, keep it for email-fallback matching: the student may be in
            # the SIS under a real id, in which case compare_students surfaces a
            # "Student ID" correction (fill in the MAP id) instead of a false
            # "Add to MAP Roster". With no email we still cannot match, so skip.
            if not student_id:
                if email:
                    map_emailonly.append(_build_map_record(row, col_map))
                    all_map_emails.add(email.lower())
                continue

            if not notes:
                # v2.7.0: Timeback CMR tabs don't maintain the Notes column. The
                # OneRoster API is the SIS source of truth, so empty Notes means
                # "currently rostered". Coerce to "Enrolled" so the student enters
                # map_enrolled and both the IM-checkbox path AND the field-mismatch
                # path can fire.
                if sheet_name in TIMEBACK_CAMPUSES:
                    notes = "Enrolled"
                else:
                    continue

            record = _build_map_record(row, col_map)
            if email:
                all_map_emails.add(email.lower())

            if notes.lower() == "enrolled":
                if student_id in students_enrolled:
                    prev = students_enrolled[student_id].get("Campus", "?")
                    curr = record["Campus"]
                    print(
                        f"    WARNING: Duplicate student_id {student_id} "
                        f"(prev={prev}, now={curr} in {sheet_name})"
                    )
                students_enrolled[student_id] = record
                enrolled_count += 1
            else:
                students_non_enrolled[student_id] = record
                non_enrolled_count += 1

        # v2.8.4: loud warning if a campus sheet yielded ZERO students. Catches
        # whole-campus breakage (corrupt headers that defeat the col-A fallback,
        # all-blank Notes on a non-Timeback sheet, etc.) that would otherwise
        # silently drop the entire campus from the comparison.
        if data_rows and enrolled_count == 0 and non_enrolled_count == 0:
            print(
                f"  *** WARNING: '{sheet_name}' processed 0 students out of "
                f"{len(data_rows)} data rows. Check the Student ID / Notes columns "
                f"(detected fields: {sorted(col_map.keys())})."
            )

        print(
            f"  {sheet_name}: {enrolled_count} enrolled, {non_enrolled_count} non-enrolled "
            f"(Notes=col {notes_col_letter}, {num_cols} cols)"
        )

    if skipped_sheets:
        print(f"\n  WARNING: Skipped sheets: {', '.join(skipped_sheets)}")

    print(
        f"\n  Total MAP roster: {len(students_enrolled):,} enrolled, "
        f"{len(students_non_enrolled):,} non-enrolled, "
        f"{len(map_emailonly):,} email-only (blank Student ID)"
    )
    return students_enrolled, students_non_enrolled, map_emailonly, all_map_emails


def _safe_get(row, idx):
    """Safely get a value from a row, returning empty string if missing.

    idx can be None (field not found in this sheet's headers) — returns "".
    """
    if idx is None or idx >= len(row):
        return ""
    val = row[idx]
    return str(val).strip() if val is not None else ""


def _combine_name(first, last):
    """Combine first and last name into a single string."""
    parts = [p.strip() for p in [first, last] if p and p.strip()]
    return " ".join(parts)


def _build_map_record(row, col_map):
    """Build a normalized MAP student record dict from a sheet row + col_map."""
    guide_first = _safe_get(row, col_map.get("guide_first"))
    guide_last = _safe_get(row, col_map.get("guide_last"))
    unenroll_raw = _safe_get(row, col_map.get("unenroll"))
    return {
        "Campus": _safe_get(row, col_map.get("campus")),
        "Grade": _safe_get(row, col_map.get("grade")),
        "Level": _safe_get(row, col_map.get("level")),
        "First Name": _safe_get(row, col_map.get("first_name")),
        "Last Name": _safe_get(row, col_map.get("last_name")),
        "Email": _safe_get(row, col_map.get("email")),
        "Student Group": _safe_get(row, col_map.get("student_group")),
        "Guide First Name": guide_first,
        "Guide Last Name": guide_last,
        "Guide Email": _safe_get(row, col_map.get("guide_email")),
        "Student_ID": _safe_get(row, col_map.get("student_id")),
        "External Student ID": _safe_get(row, col_map.get("ext_student_id")),
        "Guide Name": _combine_name(guide_first, guide_last),
        "_unenroll_flag": unenroll_raw.strip().upper() == "TRUE",  # IM-driven; option C
    }


# ── SIS Data Reader ────────────────────────────────────────────────────────


def read_sis_data(bq_client):
    """Query SIS data from BigQuery and return dict keyed by student_id.

    Raises:
        SystemExit: if BigQuery query fails. Displays actionable error message
        directing user to check the alpha_roster table exists and the service
        account has bigquery.jobs.create + bigquery.tables.getData permissions.
    """
    print_step("2. QUERYING SIS DATA FROM BIGQUERY")
    try:
        rows = query_alpha_roster(bq_client, BQ_PROJECT, BQ_DATASET, BQ_TABLE)
    except Exception as e:
        print(f"\n  ERROR: BigQuery query failed: {e}")
        print(
            f"  Check that `{BQ_PROJECT}.{BQ_DATASET}.{BQ_TABLE}` exists and the "
            f"service account has bigquery.jobs.create + bigquery.tables.getData. "
            f"If the table is missing, run `run_export.ps1` to recreate it."
        )
        sys.exit(1)

    students = {}
    for row in rows:
        student_id = (row.get("student_id") or "").strip()
        if not student_id:
            continue

        guide_name = (row.get("guide_name") or "").strip()
        guide_first, guide_last = _split_name(guide_name)

        students[student_id] = {
            "Campus": (row.get("campus") or "").strip(),
            "Grade": (row.get("grade") or "").strip(),
            "Level": (row.get("level") or "").strip(),
            "First Name": (row.get("first_name") or "").strip(),
            "Last Name": (row.get("last_name") or "").strip(),
            "Email": (row.get("email") or "").strip(),
            "Student Group": (row.get("student_group") or "").strip(),
            "Guide First Name": guide_first,
            "Guide Last Name": guide_last,
            "Guide Email": (row.get("guide_email") or "").strip(),
            "Student_ID": student_id,
            "External Student ID": (row.get("ext_student_id") or "").strip(),
            "Guide Name": guide_name,
            "admissionstatus": (row.get("admissionstatus") or "").strip(),
        }

    print(f"  Total SIS students: {len(students):,}")
    return students


def _split_name(full_name):
    """Split a full name into (first, last). Handles edge cases."""
    if not full_name or not full_name.strip():
        return "", ""
    parts = full_name.strip().split(" ", 1)
    first = parts[0]
    last = parts[1] if len(parts) > 1 else ""
    return first, last


def read_combined_sis_data(bq_client):
    """Read SIS data from BOTH sources and return a single merged dict.

    v2.7.0: Dash campuses (9) cross-reference against the alpha_roster BQ
    table (existing behavior). Timeback-backed campuses (Vita + ScienceSIS)
    cross-reference against the OneRoster API live each run.

    Both sources produce dicts keyed by `student_id` (Dash format like
    `084-13193` for both — Timeback rows use legacyDashStudentId, which
    matches the CMR Student_ID column for Vita/ScienceSIS rows).

    Skips the Timeback fetch entirely if TIMEBACK_CAMPUSES is empty —
    no extra latency for Dash-only deployments.

    On Timeback API failure: logs the error and continues with just the
    Dash sis_students dict. Vita/ScienceSIS rows will then surface as
    "Roster Addition" mismatches (because they're absent from the SIS
    dict) instead of correct "Unenrolling" detection. Better to surface
    noise than to crash the whole pipeline run.
    """
    sis_students = read_sis_data(bq_client)

    if not TIMEBACK_CAMPUSES:
        return sis_students

    print_step("2b. QUERYING TIMEBACK SIS DATA (OneRoster API)")
    try:
        timeback_students = query_timeback_enrolled(TIMEBACK_CAMPUSES)
    except Exception as e:
        print(f"\n  WARNING: Timeback API fetch failed: {e}")
        print(
            "  Vita/ScienceSIS students will surface as Roster Additions until "
            "the Timeback API recovers. Pipeline continuing with Dash data only."
        )
        return sis_students

    # Merge: Timeback wins on key collision. The Vita/ScienceSIS migration
    # window has ~62 students existing in BOTH alpha_roster (legacy Dash row)
    # AND Timeback (new system of record). Per user spec, Timeback is the
    # source of truth for these campuses, so it must override Dash on overlap.
    combined = dict(sis_students)
    overlap = set(sis_students.keys()) & set(timeback_students.keys())
    if overlap:
        print(f"  NOTE: {len(overlap)} student_id(s) appear in both Dash + Timeback")
        print(
            f"        sources. Timeback entries take precedence: {sorted(overlap)[:5]}"
        )
    combined.update(timeback_students)
    print(
        f"  Total combined SIS: {len(combined):,} students "
        f"({len(sis_students):,} Dash + {len(timeback_students):,} Timeback)"
    )
    return combined


# ── Comparison Engine ──────────────────────────────────────────────────────


def compare_students(
    map_enrolled,
    map_non_enrolled,
    sis_students,
    map_emailonly=None,
    all_map_emails=None,
):
    """Compare MAP roster against SIS data, return mismatched students.

    Detection paths (option-C precedence):
    1. IM-flagged Unenroll checkbox TRUE + SIS still Enrolled -> "Unenrolling"
       (highest priority; takes precedence over field mismatches)
    2. "Roster Addition" - enrolled in MAP, not found in SIS (by id OR email)
    3. Field mismatches - enrolled in both, specific fields differ
    4. v2.8.4 "Student ID" - MAP row has a blank Student ID but matches a SIS
       student by email (the IM just needs to fill in the MAP id)
    5. Notes-based "Unenrolling" - not Enrolled in MAP (via Notes col), but SIS
       admissionstatus=Enrolled
    6. v2.8.0 "Add to MAP Roster" - enrolled in SIS, no MAP row (by id OR email)

    Returns:
        (corrections_map, corrections_sis): parallel lists of dicts.
    """
    print_step("3. COMPARING MAP ROSTER vs SIS DATA")

    map_emailonly = map_emailonly or []
    all_map_emails = all_map_emails or set()
    # v2.8.4: SIS email index for fallback matching when Student IDs are blank/wrong.
    sis_by_email = {}
    for _rec in sis_students.values():
        _em = (_rec.get("Email") or "").strip().lower()
        if _em and _em not in sis_by_email:
            sis_by_email[_em] = _rec

    corrections_map = []
    corrections_sis = []
    match_count = 0
    roster_addition_count = 0
    field_mismatch_count = 0
    unenroll_count = 0
    im_flagged_unenroll_count = 0
    add_studentid_count = 0

    # ── Enrolled MAP students vs SIS ──────────────────────────────────
    for student_id, map_rec in sorted(map_enrolled.items()):
        sis_rec = sis_students.get(student_id)

        # Option-C: IM-flagged Unenroll trumps field mismatches.
        # If Unenroll=TRUE on CMR AND SIS still shows Enrolled → flag as Unenrolling.
        if map_rec.get("_unenroll_flag"):
            if (
                sis_rec
                and sis_rec.get("admissionstatus", "").strip().lower() == "enrolled"
            ):
                map_rec_copy = dict(map_rec)
                map_rec_copy["mismatch_summary"] = "Unenrolling"
                corrections_map.append(map_rec_copy)
                corrections_sis.append(dict(sis_rec))
                im_flagged_unenroll_count += 1
                unenroll_count += 1
                continue
            # If IM flagged but SIS already not-enrolled, nothing to flag — SIS matches.

        if sis_rec is None:
            # v2.8.4: before declaring Roster Addition, try matching by email first.
            # Catches a MAP row whose Student ID is wrong/typo'd but whose student
            # IS in the SIS. If matched, fall through to field comparison (the id
            # difference surfaces as a "Student_ID" mismatch).
            _em = (map_rec.get("Email") or "").strip().lower()
            if _em:
                sis_rec = sis_by_email.get(_em)
            if sis_rec is None:
                # Genuinely not in SIS by id OR email -> Roster Addition
                map_rec_copy = dict(map_rec)
                map_rec_copy["mismatch_summary"] = "Roster Addition"
                corrections_map.append(map_rec_copy)

                sis_placeholder = {field: "NOT FOUND IN SIS" for field in OUTPUT_FIELDS}
                sis_placeholder["Student_ID"] = student_id
                corrections_sis.append(sis_placeholder)
                roster_addition_count += 1
                continue

        # Compare fields. v2.7.1: skip noise mismatches for Timeback campuses
        # whose OneRoster API doesn't expose Level / External Student ID /
        # Student Group / Guide*.
        is_timeback = map_rec.get("Campus", "").strip() in TIMEBACK_CAMPUS_NAMES
        mismatches = _find_mismatches(map_rec, sis_rec, is_timeback=is_timeback)

        if mismatches:
            map_rec_copy = dict(map_rec)
            map_rec_copy["mismatch_summary"] = ", ".join(mismatches)
            corrections_map.append(map_rec_copy)
            corrections_sis.append(dict(sis_rec))
            field_mismatch_count += 1
        else:
            match_count += 1

    # ── Email-only MAP rows (blank Student ID) -> match by email (v2.8.4) ──
    # read_map_roster keeps blank-Student-ID rows that have an email instead of
    # dropping them. If the email matches a SIS student, the student IS in both
    # systems, just missing their MAP id. Flag a "Student ID" correction (whatever
    # the MAP enrollment status) and stamp it with the SIS Student_ID so (a) the IM
    # sees exactly which id to enter, and (b) each correction has a DISTINCT
    # (Student_ID, mismatch) key. Otherwise every blank-id row would share the
    # ("", "Student ID") tuple and accepting/hiding one would hide them all. No
    # email match -> we cannot place them, so leave as-is (needs a manual id).
    for map_rec in map_emailonly:
        _em = (map_rec.get("Email") or "").strip().lower()
        sis_rec = sis_by_email.get(_em) if _em else None
        if sis_rec is None:
            continue
        map_rec_copy = dict(map_rec)
        map_rec_copy["Student_ID"] = sis_rec.get("Student_ID", "")
        map_rec_copy["mismatch_summary"] = "Student ID"
        corrections_map.append(map_rec_copy)
        corrections_sis.append(dict(sis_rec))
        add_studentid_count += 1

    # ── Non-enrolled MAP students vs SIS (unenrolling detection) ──────
    # Skip student_ids already processed in the enrolled loop to avoid double-counting
    already_processed = {rec["Student_ID"] for rec in corrections_map}
    for student_id, map_rec in sorted(map_non_enrolled.items()):
        if student_id in map_enrolled or student_id in already_processed:
            print(
                f"    NOTE: student_id {student_id} is both enrolled and non-enrolled "
                f"across campus sheets — skipping unenrolling check"
            )
            continue
        sis_rec = sis_students.get(student_id)
        if sis_rec is None:
            continue  # not in SIS either — nothing to flag
        if sis_rec.get("admissionstatus", "").strip().lower() != "enrolled":
            continue  # not enrolled in SIS — no conflict

        # Student not enrolled in MAP but enrolled in SIS → Unenrolling
        map_rec_copy = dict(map_rec)
        map_rec_copy["mismatch_summary"] = "Unenrolling"
        corrections_map.append(map_rec_copy)
        corrections_sis.append(dict(sis_rec))
        unenroll_count += 1

    # ── SIS-only students (in SIS, not in MAP) → "Add to MAP Roster" ──────
    # v2.8.0: the reverse of "Roster Addition". A student enrolled in the SIS
    # with NO MAP row is invisible to the two loops above (both iterate MAP).
    # Scope to MANAGED campuses only: the alpha_roster BQ table is a global
    # Alpha export (~9,400 students incl. hundreds from unmanaged schools).
    # The set of Campus values present in the MAP roster IS the managed set,
    # and MAP Campus values exactly match SIS campus values, so membership
    # cleanly isolates the managed campuses. Obvious test accounts skipped.
    add_to_map_count = 0
    managed_campuses = {
        (rec.get("Campus") or "").strip()
        for rec in list(map_enrolled.values())
        + list(map_non_enrolled.values())
        + map_emailonly
        if (rec.get("Campus") or "").strip()
    }
    map_ids = set(map_enrolled) | set(map_non_enrolled)
    for student_id, sis_rec in sorted(sis_students.items()):
        if student_id in map_ids:
            continue  # has a MAP row by id; handled by the loops above
        _em = (sis_rec.get("Email") or "").strip().lower()
        if _em and _em in all_map_emails:
            continue  # v2.8.4: in MAP via an email-only (blank-id) row, not a real add
        if sis_rec.get("admissionstatus", "").strip().lower() != "enrolled":
            continue  # only currently-enrolled SIS students
        if (sis_rec.get("Campus") or "").strip() not in managed_campuses:
            continue  # unmanaged Alpha school (TSA, Colearn, etc.)
        if _is_test_account(sis_rec):
            continue  # skip obvious test accounts per user spec

        rec = dict(sis_rec)
        rec["mismatch_summary"] = "Add to MAP Roster"
        corrections_map.append(rec)
        corrections_sis.append(dict(sis_rec))
        add_to_map_count += 1

    print(f"  Matches (no correction needed): {match_count:,}")
    print(f"  Roster Additions (not in SIS): {roster_addition_count:,}")
    print(f"  Field mismatches: {field_mismatch_count:,}")
    print(f"  Student ID (in SIS by email, blank MAP id): {add_studentid_count:,}")
    print(
        f"  Unenrolling: {unenroll_count:,} "
        f"(IM-flagged: {im_flagged_unenroll_count:,}, Notes-based: {unenroll_count - im_flagged_unenroll_count:,})"
    )
    print(f"  Add to MAP Roster (in SIS, not in MAP): {add_to_map_count:,}")
    print(f"  Total corrections: {len(corrections_map):,}")

    return corrections_map, corrections_sis


def _is_test_account(rec):
    """True if the record looks like a test account (name contains 'test')."""
    name = f"{rec.get('First Name', '')} {rec.get('Last Name', '')}".lower()
    return "test" in name


def _find_mismatches(map_rec, sis_rec, is_timeback=False):
    """Compare two student records and return list of mismatched field names.

    v2.7.1: when `is_timeback=True`, skip fields the Timeback OneRoster API
    doesn't expose (Level, Student Group, Guide Email, Guide Name combine,
    External Student ID). Without this, every Vita / ScienceSIS row that
    doesn't hit the Unenrolling path generated a noise mismatch chain
    "Level, External Student ID" because MAP had values and SIS returned "".
    """
    mismatches = []

    # Fields compared on ALL campuses
    simple_fields = {
        "Campus": ("Campus", "Campus"),
        "Grade": ("Grade", "Grade"),
        "First Name": ("First Name", "First Name"),
        "Last Name": ("Last Name", "Last Name"),
        "Email": ("Email", "Email"),
    }
    # Fields compared on Dash campuses only (Timeback OneRoster API doesn't
    # expose these — would always show as noise mismatches)
    if not is_timeback:
        simple_fields.update(
            {
                "Level": ("Level", "Level"),
                "Student Group": ("Student Group", "Student Group"),
                "Guide Email": ("Guide Email", "Guide Email"),
                "External Student ID": (
                    "External Student ID",
                    "External Student ID",
                ),
            }
        )

    for label, (map_key, sis_key) in simple_fields.items():
        map_val = _normalize(map_rec.get(map_key, ""))
        sis_val = _normalize(sis_rec.get(sis_key, ""))
        if map_val != sis_val:
            mismatches.append(label)

    # Guide name comparison (Dash only — Timeback uses class-level enrollment,
    # no per-student guide on the user record).
    if not is_timeback:
        map_guide = _normalize(map_rec.get("Guide Name", ""))
        sis_guide = _normalize(sis_rec.get("Guide Name", ""))
        if map_guide != sis_guide:
            mismatches.append("Guide Name")

    return mismatches


def _normalize(val):
    """Normalize a value for comparison: strip, lowercase, collapse whitespace."""
    if val is None:
        return ""
    return " ".join(str(val).strip().lower().split())


# ── Recently-handled filter ────────────────────────────────────────────────


def read_handled_student_keys(sheets_service):
    """Return set of (student_id, mismatch_summary) tuples for all rows in the
    4 cumulative tabs (`_ApprovedData`, `_AdditionsData`, `_UnenrollData`,
    `_RejectedData`). Used to hide previously-actioned student-mismatch pairs
    from Sheet 1 so IMs don't re-review what they already handled.

    v2.7.5: removed the `days_back` cutoff. Once a (student_id, mismatch)
    pair has been actioned, it stays hidden from Sheet 1 indefinitely. If a
    NEW *different* mismatch type arises for the same student, the new tuple
    is not in handled_keys and the student resurfaces on Sheet 1.

    Reads col B (mismatch_summary, idx 1) and col M (student_id, idx 12)
    from each cumulative tab. Date column (col A) is no longer read.
    """
    handled = set()

    for tab in [
        "_ApprovedData",
        "_AdditionsData",
        "_UnenrollData",
        "_RejectedData",
        "_MapAdditionsData",  # v2.8.0: "Add to MAP Roster" accepted rows
    ]:
        # v2.5.2: wrapped in retry_api. If transient errors exhaust all
        # retries, fall through to silent skip (the next hourly run will catch up).
        try:
            resp = retry_api(
                lambda t=tab: sheets_service.spreadsheets()
                .values()
                .get(spreadsheetId=OUTPUT_SPREADSHEET_ID, range=f"'{t}'!A:M")
                .execute(),
                label=f"read handled keys from '{tab}'",
            )
        except Exception:
            continue

        for row in resp.get("values", []):
            if len(row) < 13:
                continue
            mismatch = str(row[1] or "").strip()
            sid = str(row[12] or "").strip()
            if sid:
                handled.add((sid, mismatch))

    return handled


def _hide_handled(corrections_map, corrections_sis, handled_keys):
    """Filter parallel lists to drop students whose (sid, mismatch_summary)
    tuple is in handled_keys."""
    if not handled_keys:
        return corrections_map, corrections_sis
    kept_map = []
    kept_sis = []
    for m, s in zip(corrections_map, corrections_sis):
        key = (m.get("Student_ID", ""), m.get("mismatch_summary", ""))
        if key in handled_keys:
            continue
        kept_map.append(m)
        kept_sis.append(s)
    return kept_map, kept_sis


# ── Main ───────────────────────────────────────────────────────────────────


def main():
    start_time = time.time()
    print("=" * 60)
    print("  WEEKLY CORRECTIONS — MAP Roster vs SIS Comparison")
    print("=" * 60)

    # Authenticate
    print("\n  Authenticating with service account...")
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEY, scopes=SCOPES
    )
    bq_client = bigquery.Client(project=BQ_PROJECT, credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)

    # Read data from both sources (v2.7.0: Dash via BQ + Timeback via OneRoster API)
    map_enrolled, map_non_enrolled, map_emailonly, all_map_emails = read_map_roster(
        sheets_service
    )
    sis_students = read_combined_sis_data(bq_client)

    # Compare and find mismatches
    corrections_map, corrections_sis = compare_students(
        map_enrolled, map_non_enrolled, sis_students, map_emailonly, all_map_emails
    )

    # v2.7.5: Hide ALL previously-handled (student_id, mismatch_summary) tuples
    # from Sheet 1 — no time cutoff. Once IM clicks Accept or Reject for a given
    # student-mismatch pair, that pair stays hidden indefinitely. If a NEW
    # different mismatch type arises for the same student later, it surfaces.
    handled_keys = read_handled_student_keys(sheets_service)
    if handled_keys:
        before = len(corrections_map)
        corrections_map, corrections_sis = _hide_handled(
            corrections_map, corrections_sis, handled_keys
        )
        hidden = before - len(corrections_map)
        print(
            f"  Hidden {hidden} previously-handled student-mismatch pair(s) "
            f"({len(handled_keys):,} total handled keys). "
            f"{len(corrections_map):,} corrections remain on Sheet 1."
        )

    # Write to output spreadsheet
    print_step("4. WRITING TO CORRECTIONS SPREADSHEET")
    write_corrections(sheets_service, corrections_map, corrections_sis)

    elapsed = time.time() - start_time
    print(f"\n  Completed in {elapsed:.1f}s")
    print(f"  Output: https://docs.google.com/spreadsheets/d/{OUTPUT_SPREADSHEET_ID}")


if __name__ == "__main__":
    main()
