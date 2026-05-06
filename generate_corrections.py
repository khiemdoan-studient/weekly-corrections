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
    HIDE_HANDLED_DAYS,
    TIMEBACK_CAMPUSES,  # v2.7.0
)
from datetime import datetime, timedelta
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
        (enrolled, non_enrolled) — two dicts keyed by student_id.
        enrolled = students with Notes == "Enrolled"
        non_enrolled = students with any other Notes value
    """
    print_step("1. READING MAP ROSTER")
    students_enrolled = {}
    students_non_enrolled = {}
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
            if not student_id:
                continue
            if not notes:
                # v2.7.0: Timeback CMR tabs don't maintain the Notes column —
                # the OneRoster API is the SIS source of truth, so empty
                # Notes means "currently rostered". Coerce to "Enrolled"
                # so the student enters map_enrolled and both the
                # IM-checkbox path AND field-mismatch path can fire.
                if sheet_name in TIMEBACK_CAMPUSES:
                    notes = "Enrolled"
                else:
                    continue

            guide_first = _safe_get(row, col_map.get("guide_first"))
            guide_last = _safe_get(row, col_map.get("guide_last"))

            # IM-driven Unenroll checkbox (TRUE string, Python bool, or "TRUE")
            unenroll_raw = _safe_get(row, col_map.get("unenroll"))
            unenroll_flag = unenroll_raw.strip().upper() == "TRUE"

            record = {
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
                "Student_ID": student_id,
                "External Student ID": _safe_get(row, col_map.get("ext_student_id")),
                "Guide Name": _combine_name(guide_first, guide_last),
                "_unenroll_flag": unenroll_flag,  # IM-driven; option C precedence
            }

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

        print(
            f"  {sheet_name}: {enrolled_count} enrolled, {non_enrolled_count} non-enrolled "
            f"(Notes=col {notes_col_letter}, {num_cols} cols)"
        )

    if skipped_sheets:
        print(f"\n  WARNING: Skipped sheets: {', '.join(skipped_sheets)}")

    print(
        f"\n  Total MAP roster: {len(students_enrolled):,} enrolled, {len(students_non_enrolled):,} non-enrolled"
    )
    return students_enrolled, students_non_enrolled


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


def compare_students(map_enrolled, map_non_enrolled, sis_students):
    """Compare MAP roster against SIS data, return mismatched students.

    Four detection paths (option-C precedence):
    1. IM-flagged Unenroll checkbox TRUE + SIS still Enrolled → "Unenrolling"
       (highest priority — takes precedence over field mismatches)
    2. "Roster Addition" — enrolled in MAP, not found in SIS
    3. Field mismatches — enrolled in both, specific fields differ
    4. Notes-based "Unenrolling" — not Enrolled in MAP (via Notes col),
       but SIS admissionstatus=Enrolled

    Returns:
        (corrections_map, corrections_sis) — parallel lists of dicts.
    """
    print_step("3. COMPARING MAP ROSTER vs SIS DATA")

    corrections_map = []
    corrections_sis = []
    match_count = 0
    roster_addition_count = 0
    field_mismatch_count = 0
    unenroll_count = 0
    im_flagged_unenroll_count = 0

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
            # Student enrolled in MAP but not in SIS → Roster Addition
            map_rec_copy = dict(map_rec)
            map_rec_copy["mismatch_summary"] = "Roster Addition"
            corrections_map.append(map_rec_copy)

            sis_placeholder = {field: "NOT FOUND IN SIS" for field in OUTPUT_FIELDS}
            sis_placeholder["Student_ID"] = student_id
            corrections_sis.append(sis_placeholder)
            roster_addition_count += 1
            continue

        # Compare fields
        mismatches = _find_mismatches(map_rec, sis_rec)

        if mismatches:
            map_rec_copy = dict(map_rec)
            map_rec_copy["mismatch_summary"] = ", ".join(mismatches)
            corrections_map.append(map_rec_copy)
            corrections_sis.append(dict(sis_rec))
            field_mismatch_count += 1
        else:
            match_count += 1

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

    print(f"  Matches (no correction needed): {match_count:,}")
    print(f"  Roster Additions (not in SIS): {roster_addition_count:,}")
    print(f"  Field mismatches: {field_mismatch_count:,}")
    print(
        f"  Unenrolling: {unenroll_count:,} "
        f"(IM-flagged: {im_flagged_unenroll_count:,}, Notes-based: {unenroll_count - im_flagged_unenroll_count:,})"
    )
    print(f"  Total corrections: {len(corrections_map):,}")

    return corrections_map, corrections_sis


def _find_mismatches(map_rec, sis_rec):
    """Compare two student records and return list of mismatched field names."""
    mismatches = []

    # Direct field comparisons (case-insensitive, stripped)
    simple_fields = {
        "Campus": ("Campus", "Campus"),
        "Grade": ("Grade", "Grade"),
        "Level": ("Level", "Level"),
        "First Name": ("First Name", "First Name"),
        "Last Name": ("Last Name", "Last Name"),
        "Email": ("Email", "Email"),
        "Student Group": ("Student Group", "Student Group"),
        "Guide Email": ("Guide Email", "Guide Email"),
        "External Student ID": ("External Student ID", "External Student ID"),
    }

    for label, (map_key, sis_key) in simple_fields.items():
        map_val = _normalize(map_rec.get(map_key, ""))
        sis_val = _normalize(sis_rec.get(sis_key, ""))
        if map_val != sis_val:
            mismatches.append(label)

    # Guide name comparison: combine MAP first+last → compare to SIS combined
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


def read_handled_student_ids(sheets_service, days_back):
    """Return set of student_ids handled (accepted or rejected) within the last
    `days_back` days. Used to hide recently-actioned students from Sheet 1 so
    IMs don't re-review the same correction while the data team processes it.

    Reads date (col A) and student_id (col M) from all 4 cumulative tabs.
    Supports canonical "yyyy-MM-dd HH:mm:ss" format — rows with unparseable
    dates are silently skipped (treated as unhandled).
    """
    if days_back <= 0:
        return set()

    cutoff = datetime.now() - timedelta(days=days_back)
    handled = set()

    for tab in ["_ApprovedData", "_AdditionsData", "_UnenrollData", "_RejectedData"]:
        # v2.5.2: wrapped in retry_api. If transient errors exhaust all
        # retries, fall through to silent skip (existing behavior — the
        # next hourly run will catch up).
        try:
            resp = retry_api(
                lambda t=tab: sheets_service.spreadsheets()
                .values()
                .get(spreadsheetId=OUTPUT_SPREADSHEET_ID, range=f"'{t}'!A:M")
                .execute(),
                label=f"read handled ids from '{tab}'",
            )
        except Exception:
            continue

        for row in resp.get("values", []):
            if len(row) < 13:
                continue
            date_str = str(row[0] or "").strip()
            sid = str(row[12] or "").strip()
            if not date_str or not sid:
                continue
            try:
                dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue
            if dt >= cutoff:
                handled.add(sid)

    return handled


def _hide_recently_handled(corrections_map, corrections_sis, handled_ids):
    """Filter parallel lists to drop students whose student_id is in handled_ids."""
    if not handled_ids:
        return corrections_map, corrections_sis
    kept_map = []
    kept_sis = []
    for m, s in zip(corrections_map, corrections_sis):
        if m.get("Student_ID", "") in handled_ids:
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
    map_enrolled, map_non_enrolled = read_map_roster(sheets_service)
    sis_students = read_combined_sis_data(bq_client)

    # Compare and find mismatches
    corrections_map, corrections_sis = compare_students(
        map_enrolled, map_non_enrolled, sis_students
    )

    # Hide recently-handled students (checked Accept or Reject in the last N days)
    # so IMs don't re-review the same correction while the data team processes it.
    if HIDE_HANDLED_DAYS > 0:
        handled_ids = read_handled_student_ids(sheets_service, HIDE_HANDLED_DAYS)
        if handled_ids:
            before = len(corrections_map)
            corrections_map, corrections_sis = _hide_recently_handled(
                corrections_map, corrections_sis, handled_ids
            )
            hidden = before - len(corrections_map)
            print(
                f"  Hidden {hidden} recently-handled students "
                f"(within last {HIDE_HANDLED_DAYS} days). "
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
