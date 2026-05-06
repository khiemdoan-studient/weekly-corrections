"""timeback_sis.py — OneRoster API bridge for Timeback-backed campuses.

Source-of-truth for "currently enrolled" for Vita + ScienceSIS students.
Wraps just the OAuth2 + GET /schools/{id}/students endpoint of the OneRoster
API at api.alpha-1edtech.ai. Returns rows shaped like the alpha_roster BQ
table so generate_corrections.py can merge both sources into a single
SIS dict.

The full OneRoster client lives in the sibling repo
`timeback-data-pipeline/oneroster_client.py`. This module is a deliberate
narrow extraction — only what weekly-corrections needs for the
"is this student currently rostered at this school?" check.

Credentials
-----------
Read from `keys/timeback-creds.json` by default (path = TIMEBACK_CREDS_PATH
in config.py). File contains:
    {"client_id": "...", "client_secret": "..."}

In GHA, the deploy workflow writes the TIMEBACK_CREDS_JSON secret to this
path before invoking the pipeline. Falls back to env vars
COGNITO_CLIENT_ID + COGNITO_CLIENT_SECRET if the file is missing.

Identifier bridge
-----------------
The CMR's `Student_ID` column for Vita + ScienceSIS rows holds the
Timeback `metadata.legacyDashStudentId` (e.g. "066-6757", "033-2154"),
NOT the Timeback `sourcedId` UUID. So we key the returned dict by
legacyDashStudentId to match what compare_students looks up.

Students whose `legacyDashStudentId` is blank are skipped — they can't
be matched to the CMR row anyway, and they're typically test accounts.
"""

import json
import os
import time

import requests

from config import TIMEBACK_CAMPUSES, TIMEBACK_CREDS_PATH

# ── OneRoster API endpoints (must stay in sync with sibling repo) ──────────
BASE_URL = "https://api.alpha-1edtech.ai"
TOKEN_URL = (
    "https://prod-beyond-timeback-api-2-idp.auth.us-east-1.amazoncognito.com"
    "/oauth2/token"
)


# ── Credentials ────────────────────────────────────────────────────────────


def _load_credentials():
    """Return (client_id, client_secret). File path → env var fallback."""
    if os.path.exists(TIMEBACK_CREDS_PATH):
        with open(TIMEBACK_CREDS_PATH) as f:
            data = json.load(f)
        cid = data.get("client_id", "").strip()
        cs = data.get("client_secret", "").strip()
        if cid and cs:
            return cid, cs

    cid = os.environ.get("COGNITO_CLIENT_ID", "").strip()
    cs = os.environ.get("COGNITO_CLIENT_SECRET", "").strip()
    if cid and cs:
        return cid, cs

    raise RuntimeError(
        f"Timeback credentials not found. Either:\n"
        f"  (a) place a JSON file at {TIMEBACK_CREDS_PATH} with "
        f'{{"client_id": "...", "client_secret": "..."}}, or\n'
        f"  (b) set env vars COGNITO_CLIENT_ID + COGNITO_CLIENT_SECRET."
    )


# ── Minimal OneRoster client ───────────────────────────────────────────────


class _OneRosterClient:
    """Self-contained OneRoster client. OAuth2 + paginated school students.

    Intentionally narrow — this is NOT the sibling repo's full client; it
    only does what weekly-corrections needs. Keep it that way: any other
    OneRoster lookup goes in timeback-data-pipeline.
    """

    def __init__(self, client_id, client_secret):
        self.client_id = client_id
        self.client_secret = client_secret
        self._token = None
        self._token_expires_at = 0

    def _authenticate(self):
        resp = requests.post(
            TOKEN_URL,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            data={
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        self._token = data["access_token"]
        # Subtract 60s safety margin
        self._token_expires_at = time.time() + data.get("expires_in", 3600) - 60

    def _headers(self):
        if not self._token or time.time() >= self._token_expires_at:
            self._authenticate()
        return {"Authorization": f"Bearer {self._token}"}

    def _get(self, path, params=None, retries=3):
        for attempt in range(retries):
            try:
                resp = requests.get(
                    f"{BASE_URL}{path}",
                    headers=self._headers(),
                    params=params,
                    timeout=120,
                )
                if resp.status_code == 401:
                    self._authenticate()
                    continue
                if resp.status_code == 429:
                    wait = (2**attempt) * 5
                    print(f"  [Timeback] Rate limited, waiting {wait}s")
                    time.sleep(wait)
                    continue
                resp.raise_for_status()
                return resp.json()
            except requests.exceptions.RequestException as e:
                if attempt == retries - 1:
                    raise
                wait = 2**attempt
                print(f"  [Timeback] retry {attempt + 1} after {wait}s: {e}")
                time.sleep(wait)
        return None

    def get_students(self, school_id):
        """Paginate through all currently-rostered students at a school.

        Returns a list of OneRoster user records. Membership in this list
        IS the "currently enrolled" signal — the OneRoster server only
        returns active rosterings.
        """
        all_users = []
        offset = 0
        limit = 1000
        while True:
            data = self._get(
                f"/ims/oneroster/rostering/v1p2/schools/{school_id}/students",
                params={"offset": offset, "limit": limit},
            )
            if not data:
                break
            users = data.get("users", [])
            all_users.extend(users)
            if len(users) < limit:
                break
            offset += limit
        return all_users


# ── Public API ─────────────────────────────────────────────────────────────


def query_timeback_enrolled(timeback_campuses=None):
    """For each Timeback campus, fetch currently-rostered students and
    return a dict shaped like alpha_roster output.

    Args:
        timeback_campuses: dict of {campus_label: school_sourcedId}.
            Defaults to config.TIMEBACK_CAMPUSES.

    Returns:
        Dict keyed by legacyDashStudentId. Each value is a dict with the
        same keys read_sis_data populates from alpha_roster:
            Campus, Grade, Level, First Name, Last Name, Email,
            Student Group, Guide First Name, Guide Last Name, Guide Email,
            Student_ID, External Student ID, Guide Name, admissionstatus.

        admissionstatus is always "Enrolled" — only currently-rostered
        students appear in the OneRoster response. A student missing from
        the dict (because the OneRoster API didn't return them) is treated
        as not-enrolled by compare_students, which is exactly the signal
        we want for the "Unenrolling" mismatch.

    Raises:
        RuntimeError: if credentials are missing.
        requests.exceptions.RequestException: after all retries exhausted.
    """
    if timeback_campuses is None:
        timeback_campuses = TIMEBACK_CAMPUSES

    if not timeback_campuses:
        return {}

    cid, cs = _load_credentials()
    client = _OneRosterClient(cid, cs)

    out = {}
    skipped_no_sid = 0
    for campus_label, school_id in timeback_campuses.items():
        # v2.7.1: strip " (TimeBack)" suffix so the Campus value matches what's
        # actually in the CMR Campus column (e.g. "ScienceSIS", not
        # "ScienceSIS (TimeBack)"). Without this, every Timeback row produced
        # a noise Campus mismatch on every comparison.
        campus_short = campus_label.replace(" (TimeBack)", "")
        print(f"  Timeback: fetching '{campus_label}' (school_id={school_id[:8]}…)")
        users = client.get_students(school_id)
        kept = 0
        for u in users:
            meta = u.get("metadata", {}) or {}
            legacy_sid = (meta.get("legacyDashStudentId") or "").strip()
            if not legacy_sid:
                skipped_no_sid += 1
                continue

            given = (u.get("givenName") or "").strip()
            family = (u.get("familyName") or "").strip()
            email = (u.get("email") or u.get("userMasterIdentifier") or "").strip()
            grades = u.get("grades") or []
            grade_str = ""
            if isinstance(grades, list) and grades:
                grade_str = str(grades[0]).strip()

            # Student Group — meta key spelling varies by tenant
            student_group = (
                meta.get("studentGroup")
                or meta.get("Student Group")
                or meta.get("StudentGroup")
                or ""
            ).strip()

            # Guide / teacher names not exposed on the user record itself.
            # Leave blank — Vita/ScienceSIS rows won't trigger "Guide" mismatches
            # because the comparison lower-cases empty == empty.
            out[legacy_sid] = {
                "Campus": (meta.get("Campus") or campus_short).strip(),
                "Grade": grade_str,
                "Level": (meta.get("level") or "").strip(),
                "First Name": given,
                "Last Name": family,
                "Email": email,
                "Student Group": student_group,
                "Guide First Name": "",
                "Guide Last Name": "",
                "Guide Email": "",
                "Student_ID": legacy_sid,
                "External Student ID": "",  # Timeback schools don't expose this
                "Guide Name": "",
                "admissionstatus": "Enrolled",
            }
            kept += 1
        print(f"    -> {kept} students with legacyDashStudentId")

    if skipped_no_sid:
        print(f"  Timeback: {skipped_no_sid} students skipped (no legacyDashStudentId)")
    print(
        f"  Timeback total: {len(out):,} students across {len(timeback_campuses)} campus(es)"
    )
    return out
