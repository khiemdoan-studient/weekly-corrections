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
    "ext_student_id": {"suns number", "external student id", "suns #", "external id"},
    "unenroll": {"unenroll", "unenrolled"},
}

# ── ISR (Individual Student Roster) configuration ──────────────────────────
# Each campus has a dedicated Google Sheet ("ISR") containing:
#   - "Student Roster" tab (SR) — IMs edit directly
#   - "MAP Roster" tab (MR) — formula-derived view that feeds the CMR
#   - "Auto Synced from SIS" — SIS admission status dump
# The ISR's MR tab is pulled into the Combined MAP Roster (CMR) via IMPORTRANGE.
#
# Per-ISR column positions (as of v2.3.0 audit). `sr_unenroll_col` is 0-based
# index. `mr_unenroll_col` is where we write the new Unenroll mirror column.
ISR_CONFIG = {
    "Reading CCSD (Dash)": {
        "isr_id": "1b28bgPy9mysb31Op01DPL6IS5jhMZy51VKVSsS0feII",
        "mr_gid": 1256615349,
        "sr_unenroll_col": 23,  # col X — already exists
        "mr_unenroll_col": 30,  # col AE — new
    },
    "Metro Schools (Dash)": {
        "isr_id": "1Eri0B_WMmjJxPs6SYszK2F1jJt08rjPFRwgrpuKdakU",
        "mr_gid": 438424229,
        "sr_unenroll_col": 24,  # col Y — already exists as "Unenrolled"
        "mr_unenroll_col": 27,  # col AB — new
    },
    "Allendale Fairfax Elementary School (Dash)": {
        "isr_id": "1zhWCgoJB9WXA9sDxnHj0uZbHQaWTiMFk3bLP61rUKWo",
        "mr_gid": 1256615349,
        "sr_unenroll_col": 25,  # col Z — already exists
        "mr_unenroll_col": 27,  # col AB — new
    },
    "Allendale Fairfax Middle School (Dash)": {
        "isr_id": "1r6o0j8ENz01gt9L5ygJLBZAtwCD-L9H2SrutgZyTfQc",
        "mr_gid": 1256615349,
        "sr_unenroll_col": 25,  # col Z — already exists
        "mr_unenroll_col": 27,  # col AB — new
    },
    "Allendale Aspire Academy (Dash)": {
        "isr_id": "10H5y0Z3_QAH9wYH5V80yLSLuqWPKpDXf5k7wDMz7hww",
        "mr_gid": 1256615349,
        "sr_unenroll_col": 27,  # col AB — already exists
        "mr_unenroll_col": 27,  # col AB — new
    },
    "Hardeeville Elementary School (Dash)": {
        "isr_id": "1waahGamoiMb5DkLF1_IlO5kEhpcc9g7NZr3WIeiAfFw",
        "mr_gid": 1256615349,
        "sr_unenroll_col": 27,  # col AB — NEEDS TO BE ADDED
        "mr_unenroll_col": 29,  # col AD — new
    },
    "Hardeeville Junior & Senior High School (Dash)": {
        "isr_id": "1g8KUreiGlBd2NM5huZjSSDA30YdDD8kL0geUJb0Ajww",
        "mr_gid": 1256615349,
        "sr_unenroll_col": 27,  # col AB — NEEDS TO BE ADDED
        "mr_unenroll_col": 29,  # col AD — new
    },
    "Ridgeland Elementary School (Dash)": {
        "isr_id": "1IwGsdtThjQJmcfbh_eR5ZrFKQiCj9FL5GR02geRofWQ",
        "mr_gid": 1256615349,
        "sr_unenroll_col": 27,  # col AB — NEEDS TO BE ADDED
        "mr_unenroll_col": 29,  # col AD — new
    },
    "Ridgeland Secondary Academy of Excellence (Dash)": {
        "isr_id": "1AT4jEZPbaYdFJUI8OTAVIgHOs4cjCY6zh96I7uMvZZM",
        "mr_gid": 1256615349,
        "sr_unenroll_col": 27,  # col AB — NEEDS TO BE ADDED
        "mr_unenroll_col": 29,  # col AD — new
    },
}

# ── Output sheet tab names ──────────────────────────────────────────────────
# ── Row-hiding behavior ────────────────────────────────────────────────────
# Students who've been Accept'd or Reject'd within this many days are hidden
# from Sheet 1 "Corrected Roster Info" on the next pipeline run. This prevents
# IMs from re-reviewing the same correction while the data team processes it.
# After HIDE_HANDLED_DAYS elapses, if the mismatch still exists in MAP vs SIS,
# the student reappears on Sheet 1 (signal that the correction hasn't been
# processed yet).
HIDE_HANDLED_DAYS = 7

# ── Output sheet tab names ──────────────────────────────────────────────────
TAB_CORRECTED = "Corrected Roster Info"
TAB_SIS = "Current Roster Info in SIS"
TAB_APPROVED = "Automated Correction List"
TAB_ADDITIONS = "Roster Additions"
TAB_UNENROLL = "Roster Unenrollments"
TAB_REJECTED = "Rejected Changes"

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
