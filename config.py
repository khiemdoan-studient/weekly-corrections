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
    # v2.7.0: Timeback-backed campuses. SIS-side enrollment lookup goes via the
    # OneRoster API (see TIMEBACK_CAMPUSES below + timeback_sis.py), NOT the
    # alpha_roster BQ table. Students still appear on the same Sheet 1 with
    # the same Mismatch Summary semantics.
    "ScienceSIS (TimeBack)",
    "Vita High School (TimeBack)",
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
    "ext_student_id": {
        "suns number",
        "external student id",
        "suns #",
        "external id",
        "alpha student id",  # v2.7.0: Vita + ScienceSIS CMR header
    },
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
    # ── v2.7.0: Timeback-backed ISRs ─────────────────────────────────────
    # Layout differs from Dash ISRs: SR has 23 cols, MR has 40 cols.
    # No Notes/Unenroll column on SR yet; setup_unenroll_columns.py adds
    # them at SR col X (24th, idx 23) + MR col AB (28th, idx 27, mirroring
    # the CMR Unenroll column position). The CMR's "Unenroll" header
    # already exists at col AB on both Timeback campus tabs.
    "ScienceSIS (TimeBack)": {
        "isr_id": "1SjVoQRubz_nsD3YVKLf68KcTaZwx1s7CA5S9V_E8gQ8",
        "mr_gid": 1256615349,
        "sr_unenroll_col": 23,  # col X — NEW (no header yet)
        "mr_unenroll_col": 27,  # col AB — NEW, matches CMR layout
    },
    "Vita High School (TimeBack)": {
        "isr_id": "1sOSwvwPb8cXSfJgXF-E2Ur2v0lyvi1qkR94OvgQLQ4Y",
        "mr_gid": 1256615349,
        "sr_unenroll_col": 23,
        "mr_unenroll_col": 27,
    },
}

# ── Timeback (OneRoster) SIS bridge — v2.7.0 ───────────────────────────────
# For Vita + ScienceSIS, the source-of-truth for "currently enrolled" is the
# Timeback OneRoster API, NOT the alpha_roster BQ table. Each pipeline run
# calls oneroster_client.get_students(school_id) for every Timeback campus
# and merges the result into the SIS dict that compare_students consumes.
#
# Mapping is CMR tab name (must match CAMPUS_SHEETS) → Timeback school
# sourcedId (UUID). Sourced from timeback-data-pipeline/oneroster_client.py
# SCHOOL_IDS constant.
TIMEBACK_CAMPUSES = {
    "ScienceSIS (TimeBack)": "7c475cf4-12b4-40ed-8857-dc6e624a5fa1",
    "Vita High School (TimeBack)": "e57cb46d-b6b0-4f45-96ed-327441b5d068",
}

# v2.7.1: bare campus values (no " (TimeBack)" suffix) — what actually appears
# in the CMR Campus column for these rows. Used by `_find_mismatches` to
# detect Timeback rows and skip fields the OneRoster API doesn't expose
# (Level, External Student ID, Student Group, Guide*). Without this skip,
# every Timeback row produces noise mismatches like "Level, External Student
# ID" because MAP has those fields populated and Timeback returns "".
TIMEBACK_CAMPUS_NAMES = {"ScienceSIS", "Vita High School"}

# Path to Timeback API credentials JSON. Contains:
#   {"client_id": "...", "client_secret": "..."}
# Locally: file at this path. In GHA: workflow writes the TIMEBACK_CREDS_JSON
# secret to this path before invoking generate_corrections.py.
TIMEBACK_CREDS_PATH = os.path.join(KEYS_DIR, "timeback-creds.json")

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

# ── Weekly Snapshot (v2.5.0) ───────────────────────────────────────────────
# Each Monday a new Google Sheet is created named "M/D Corrections" (e.g.
# "4/20 Corrections"), containing up to 3 tabs — Correction List, Roster
# Additions, Roster Unenrollments — populated with rows from the hidden
# cumulative tabs that have NOT yet been included in a prior weekly snapshot.
#
# Mechanism: a "Sent Week" column (col O, 0-indexed 14) on _ApprovedData /
# _AdditionsData / _UnenrollData. When the weekly snapshot runs, any row
# with blank Sent Week OR Sent Week == current-Monday ISO date is included;
# the row's Sent Week is then stamped with the current Monday. Re-running
# the same week updates the sheet in place.
#
# _RejectedData does NOT get a Sent Week column (rejected rows don't go
# to support).

# Shared Drive ID — owned org-level, so no user quota issues. SA is added as
# Content Manager of this drive in Google Drive UI. Files created here are
# owned by the drive itself; every drive member automatically has access,
# so no per-file sharing is needed.
WEEKLY_SHARED_DRIVE_ID = "0AFQGIqcKjsyFUk9PVA"
WEEKLY_SHARED_DRIVE_NAME = "Weekly Corrections Archive"
WEEKLY_TIMEZONE = "America/New_York"

SENT_WEEK_COL = 14  # 0-based column index for "Sent Week" on cumulative tabs
SENT_WEEK_HEADER = "Sent Week"

# Tab names inside the weekly snapshot file (user-facing)
WEEKLY_TAB_CORRECTIONS = "Correction List"
WEEKLY_TAB_ADDITIONS = "Roster Additions"
WEEKLY_TAB_UNENROLLMENTS = "Roster Unenrollments"
# v2.6.1: support-packet Instructions tab (added when --all-unsent flag is used).
# Pinned to sheet index 0 (first tab) so support sees it on open.
WEEKLY_TAB_INSTRUCTIONS = "Instructions"

# Weekly-tab → source cumulative tab mapping
WEEKLY_SOURCE_TABS = {
    WEEKLY_TAB_CORRECTIONS: "_ApprovedData",
    WEEKLY_TAB_ADDITIONS: "_AdditionsData",
    WEEKLY_TAB_UNENROLLMENTS: "_UnenrollData",
}

# Column headers for the weekly snapshot tabs (14 cols, matches approval sheets)
WEEKLY_HEADERS = [
    "Date Approved",
    "Mismatch Summary",
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
