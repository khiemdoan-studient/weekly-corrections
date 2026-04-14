"""Configuration constants for the Weekly Corrections tool."""

import os

# ── Paths ───────────────────────────────────────────────────────────────────
KEYS_DIR = os.path.join(os.path.dirname(__file__), "keys")
SERVICE_ACCOUNT_KEY = os.path.join(KEYS_DIR, "sa-main.json")

# ── Google Cloud / BigQuery ─────────────────────────────────────────────────
BQ_PROJECT = "studient-flat-exports-doan"
BQ_DATASET = "studient_analytics"
BQ_TABLE = "alpha_roster"

SCOPES = [
    "https://www.googleapis.com/auth/bigquery",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

# ── Spreadsheet IDs ─────────────────────────────────────────────────────────
MAP_SPREADSHEET_ID = "1scEay0a8OR6vU3uJuxbHKWCEx_RVgSsRXF9naJh3XYw"
OUTPUT_SPREADSHEET_ID = "12dqu58KKdsZN9nLre9Fntkk7vSILu3KfcW4WDvo5-Ls"

# ── MAP Roster campus sheets ────────────────────────────────────────────────
CAMPUS_SHEETS = [
    "Ridgeland Secondary Academy of Excellence (Dash)",
    "Ridgeland Elementary School (Dash)",
    "Hardeeville Junior & Senior High School (Dash)",
    "Hardeeville Elementary School (Dash)",
    "Allendale Aspire Academy (Dash)",
    "Allendale Fairfax Middle School (Dash)",
    "Allendale Fairfax Elementary School (Dash)",
    "Metro Schools (Dash)",
    "Reading CCSD (Dash)",
]

# ── MAP Roster header-based column mapping ──────────────────────────────────
# Maps our internal field names → set of possible header strings (lowercase).
# The reader auto-detects columns from row 1 headers, so this works even when
# campus sheets have different layouts (e.g. Reading CCSD has an extra "Full Name"
# column that shifts indices, and Notes is in column O instead of N).
MAP_HEADER_MAP = {
    "student_id": {"student id"},
    "email": {"student email"},
    "campus": {"campus"},
    "last_name": {"last name"},
    "first_name": {"first name"},
    "grade": {"grade"},
    "level": {"level"},
    "notes": {"notes"},
    "student_group": {
        "school, if separate from campus (e.g., in case the campus is a school district)",
        "school_name",
        "school",
    },
    "guide_first": {"teacher 1 first name", "teacher_first"},
    "guide_last": {"teacher 1 last name", "teacher_last"},
    "guide_email": {"teacher 1 email", "teacher_email"},
    "ext_student_id": {"suns number"},
}

# ── Output sheet tab names ──────────────────────────────────────────────────
TAB_CORRECTED = "Corrected Roster Info"
TAB_SIS = "Current Roster Info in SIS"
TAB_APPROVED = "Automated Correction List"

# ── Output column headers ──────────────────────────────────────────────────
OUTPUT_FIELDS = [
    "Campus",
    "Grade",
    "Level",
    "First Name",
    "Last Name",
    "Email",
    "Student Group",
    "Guide First Name",
    "Guide Last Name",
    "Guide Email",
    "Student_ID",
    "External Student ID",
]

# Comparison fields (subset of OUTPUT_FIELDS used for mismatch detection)
COMPARE_FIELDS = [
    "Campus",
    "Grade",
    "Level",
    "First Name",
    "Last Name",
    "Email",
    "Student Group",
    "Guide Name",
    "Guide Email",
    "External Student ID",
]
