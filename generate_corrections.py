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
)
from queries import query_alpha_roster
from sheets_writer import write_corrections


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
        # Read entire sheet including header row
        range_str = f"'{sheet_name}'!A1:AC"
        try:
            resp = (
                sheets_service.spreadsheets()
                .values()
                .get(
                    spreadsheetId=MAP_SPREADSHEET_ID,
                    range=range_str,
                )
                .execute()
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
                continue

            guide_first = _safe_get(row, col_map.get("guide_first"))
            guide_last = _safe_get(row, col_map.get("guide_last"))

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


# ── Comparison Engine ──────────────────────────────────────────────────────


def compare_students(map_enrolled, map_non_enrolled, sis_students):
    """Compare MAP roster against SIS data, return mismatched students.

    Three mismatch categories:
    1. "Roster Addition" — student Enrolled in MAP, not found in SIS at all
    2. Field mismatches — student Enrolled in both, fields differ (e.g. "Grade, Email")
    3. "Unenrolling" — student NOT Enrolled in MAP, but Enrolled in SIS

    Returns:
        (corrections_map, corrections_sis) — parallel lists of dicts for
        Sheet 1 (MAP data) and Sheet 2 (SIS data), only for mismatched students.
    """
    print_step("3. COMPARING MAP ROSTER vs SIS DATA")

    corrections_map = []
    corrections_sis = []
    match_count = 0
    roster_addition_count = 0
    field_mismatch_count = 0
    unenroll_count = 0

    # ── Enrolled MAP students vs SIS ──────────────────────────────────
    for student_id, map_rec in sorted(map_enrolled.items()):
        sis_rec = sis_students.get(student_id)

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
    print(f"  Unenrolling: {unenroll_count:,}")
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

    # Read data from both sources
    map_enrolled, map_non_enrolled = read_map_roster(sheets_service)
    sis_students = read_sis_data(bq_client)

    # Compare and find mismatches
    corrections_map, corrections_sis = compare_students(
        map_enrolled, map_non_enrolled, sis_students
    )

    # Write to output spreadsheet
    print_step("4. WRITING TO CORRECTIONS SPREADSHEET")
    write_corrections(sheets_service, corrections_map, corrections_sis)

    elapsed = time.time() - start_time
    print(f"\n  Completed in {elapsed:.1f}s")
    print(f"  Output: https://docs.google.com/spreadsheets/d/{OUTPUT_SPREADSHEET_ID}")


if __name__ == "__main__":
    main()
