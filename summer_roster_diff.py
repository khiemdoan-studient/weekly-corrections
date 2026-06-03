"""summer_roster_diff.py - Diff the combined Summer School Roster against a past
Drive revision, or against the per-ISR source of truth, and optionally mark the
delta red.

Why this exists: the combined "Summer School Roster" tab on the CMR is a live
QUERY over the per-campus tabs, which IMPORTRANGE the ISRs. IMPORTRANGE can go
STALE (on 2026-06-03 the JRES per-campus tab showed the old wrong students for
hours/days until a rebuild refreshed it, so the tech team configured 20 wrong
students that morning). This tool:
  (a) finds who was added/removed since a past point via Drive revision history
      ("who is new since this morning?"),
  (b) detects staleness by comparing the VISIBLE roster vs the source of truth
      (the per-ISR MAP Roster summer flags), and
  (c) marks new students red (float-to-top) by appending to the _Highlight tab.

PII-free: operates on live data; no student data is hardcoded. Diff output goes
to stdout and an optional --csv path (keep that path gitignored, e.g. _scratch_*).

Usage:
  # who was added/removed since a past revision or date
  python summer_roster_diff.py --since-rev 3286
  python summer_roster_diff.py --since-date 2026-06-03
  python summer_roster_diff.py --since-date 2026-06-03 --csv _scratch_added.csv

  # staleness check: does the visible roster match the source? (exit 1 on drift)
  python summer_roster_diff.py --vs-source

  # mark everyone added since a baseline red (float to top of the roster)
  python summer_roster_diff.py --since-date 2026-06-03 --mark-red
"""

import argparse
import csv
import io
import sys

from google.oauth2 import service_account
from google.auth.transport.requests import AuthorizedSession
from googleapiclient.discovery import build

from config import SERVICE_ACCOUNT_KEY, SCOPES, MAP_SPREADSHEET_ID, ISR_CONFIG
from retry_helper import retry_api
import setup_summer_school_columns as S

ODS_MIME = "application/x-vnd.oasis.opendocument.spreadsheet"
# Summer School Roster column layout (CORE_HEADERS + SUMMER_HEADERS):
# A=Student ID, B=Email, C=Campus, D=NWEA, E=Last, F=First, G=Grade, ...
COL_EMAIL, COL_CAMPUS, COL_LAST, COL_FIRST, COL_GRADE = 1, 2, 4, 5, 6


def build_services():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEY, scopes=SCOPES
    )
    sheets = build("sheets", "v4", credentials=creds)
    drive = build("drive", "v3", credentials=creds)
    return sheets, drive, AuthorizedSession(creds)


def roster_tab_gid(sheets):
    meta = retry_api(
        lambda: sheets.spreadsheets()
        .get(
            spreadsheetId=MAP_SPREADSHEET_ID, fields="sheets.properties(title,sheetId)"
        )
        .execute(),
        label="get CMR tab gids",
    )
    for s in meta["sheets"]:
        if s["properties"]["title"] == S.ROSTER_TAB:
            return s["properties"]["sheetId"]
    raise ValueError(f"'{S.ROSTER_TAB}' tab not found on the CMR")


def _info(row):
    """Pull (campus, last, first, grade) from a roster row (list)."""
    g = lambda i: (str(row[i]).strip() if len(row) > i else "")
    return (g(COL_CAMPUS).split(" - ")[0], g(COL_LAST), g(COL_FIRST), g(COL_GRADE))


def current_roster(sheets):
    """{email_lower: (campus, last, first, grade)} from the live roster, in order."""
    vals = retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .get(
            spreadsheetId=MAP_SPREADSHEET_ID,
            range=f"'{S.ROSTER_TAB}'!A2:G",
            valueRenderOption="FORMATTED_VALUE",
        )
        .execute(),
        label="read live roster",
    ).get("values", [])
    out = {}
    for r in vals:
        em = (r[COL_EMAIL] if len(r) > COL_EMAIL else "").strip().lower()
        if em:
            out[em] = _info(r)
    return out


def _rows_from_rev(session, rev, gid):
    """{email_lower: (campus, last, first, grade)} from the roster tab at a revision."""
    link = (rev.get("exportLinks") or {}).get(ODS_MIME)
    if not link:
        raise ValueError(f"revision {rev['id']} has no ODS export link")
    url = link.replace("exportFormat=ods", "exportFormat=csv") + f"&gid={gid}"
    resp = session.get(url)
    resp.raise_for_status()
    rows = list(csv.reader(io.StringIO(resp.content.decode("utf-8", "ignore"))))
    out = {}
    for r in rows[1:]:
        em = (r[COL_EMAIL] if len(r) > COL_EMAIL else "").strip().lower()
        if em:
            out[em] = _info(r)
    return out


def list_revisions(drive):
    return retry_api(
        lambda: drive.revisions()
        .list(
            fileId=MAP_SPREADSHEET_ID,
            fields="revisions(id,modifiedTime,exportLinks)",
            pageSize=1000,
        )
        .execute(),
        label="list revisions",
    ).get("revisions", [])


def pick_revision(drive, rev_id=None, before_date=None):
    revs = list_revisions(drive)
    if rev_id:
        rev = next((r for r in revs if r["id"] == rev_id), None)
        if not rev:
            raise ValueError(f"revision {rev_id} not found")
        return rev
    before = [r for r in revs if r["modifiedTime"] < before_date]
    if not before:
        raise ValueError(f"no revision before {before_date}")
    return max(before, key=lambda r: r["modifiedTime"])


def source_true_emails(sheets):
    """Union of each ISR's MAP Roster computed-TRUE summer emails = the precise
    "who should be visible" source of truth (accounts for email matching)."""
    out = set()
    for tab in S.SUMMER_TABS:
        isr = ISR_CONFIG[tab]["isr_id"]
        pos = S.read_summer_positions(sheets, isr, "MAP Roster")
        flag = S.col_letter(min(pos.values()))
        vr = retry_api(
            lambda isr=isr, flag=flag: sheets.spreadsheets()
            .values()
            .batchGet(
                spreadsheetId=isr,
                ranges=["'MAP Roster'!B2:B", f"'MAP Roster'!{flag}2:{flag}"],
                valueRenderOption="FORMATTED_VALUE",
            )
            .execute(),
            label=f"read MR TRUE set {tab}",
        ).get("valueRanges", [])
        em = vr[0].get("values", []) if vr else []
        fl = vr[1].get("values", []) if len(vr) > 1 else []
        for i in range(len(em)):
            if (
                em[i]
                and i < len(fl)
                and fl[i]
                and str(fl[i][0]).strip().upper() == "TRUE"
            ):
                out.add(em[i][0].strip().lower())
    return out


def highlight_emails(sheets):
    vals = retry_api(
        lambda: sheets.spreadsheets()
        .values()
        .get(spreadsheetId=MAP_SPREADSHEET_ID, range=f"'{S.HIGHLIGHT_TAB}'!A2:A")
        .execute(),
        label="read _Highlight",
    ).get("values", [])
    return set(r[0].strip().lower() for r in vals if r and r[0].strip())


def _print_group(title, emails, info):
    print(f"\n{title} ({len(emails)}):")
    for e in sorted(emails, key=lambda x: (info.get(x, ("",))[0], x)):
        c, last, first, grade = info.get(e, ("?", "?", "?", "?"))
        print(f"   {c:<6} {last:<20} {first:<14} g={grade:<3} {e}")


def run_vs_source(sheets, cur):
    cur_set = set(cur)
    src = source_true_emails(sheets)
    missing = src - cur_set  # flagged at source but NOT visible (roster behind = stale)
    extra = cur_set - src  # visible but NOT flagged at source (roster showing old)
    print(f"visible roster={len(cur_set)} | source (ISR MAP Roster TRUE)={len(src)}")
    if not missing and not extra:
        print("CONSISTENT - the visible roster matches the source. No stale drift.")
        return 0
    print("STALE DRIFT DETECTED - the visible roster does NOT match the source:")
    if missing:
        print(
            f"  flagged at source but MISSING from the roster (roster behind) ({len(missing)}): {sorted(missing)}"
        )
    if extra:
        _print_group(
            "  visible but NOT flagged at source (roster showing removed/old)",
            extra,
            cur,
        )
    print(
        "\nFix: re-run setup_summer_school_columns.py (rebuilds the roster QUERY and "
        "re-writes the per-campus IMPORTRANGE, forcing a fresh pull)."
    )
    return 1


def run_diff(sheets, drive, session, cur, args):
    gid = roster_tab_gid(sheets)
    rev = pick_revision(drive, rev_id=args.since_rev, before_date=args.since_date)
    base = _rows_from_rev(session, rev, gid)
    print(f"baseline revision {rev['id']} @ {rev['modifiedTime']}")
    cur_set, base_set = set(cur), set(base)
    added = cur_set - base_set
    removed = base_set - cur_set
    print(
        f"baseline={len(base_set)} | current={len(cur_set)} | +{len(added)} added / -{len(removed)} removed"
    )
    _print_group("ADDED since baseline (on the list now, not at baseline)", added, cur)
    _print_group(
        "REMOVED since baseline (on the list at baseline, gone now)", removed, base
    )

    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["status", "campus", "last", "first", "grade", "email"])
            for e in sorted(added, key=lambda x: (cur.get(x, ("",))[0], x)):
                w.writerow(["added", *cur.get(e, ("", "", "", "")), e])
            for e in sorted(removed, key=lambda x: (base.get(x, ("",))[0], x)):
                w.writerow(["removed", *base.get(e, ("", "", "", "")), e])
        print(f"\nwrote {args.csv} ({len(added)} added + {len(removed)} removed rows)")

    if args.mark_red:
        hl = highlight_emails(sheets)
        ordered = [e for e in cur if e in added and e not in hl]  # roster order
        if not ordered:
            print("\n--mark-red: nothing to add (all added students are already red).")
        else:
            label = f"added since {args.since_rev or args.since_date}"
            retry_api(
                lambda: sheets.spreadsheets()
                .values()
                .append(
                    spreadsheetId=MAP_SPREADSHEET_ID,
                    range=f"'{S.HIGHLIGHT_TAB}'!A:B",
                    valueInputOption="RAW",
                    insertDataOption="INSERT_ROWS",
                    body={"values": [[e, label] for e in ordered]},
                )
                .execute(),
                label="append _Highlight",
            )
            print(
                f"\n--mark-red: appended {len(ordered)} email(s) to {S.HIGHLIGHT_TAB} (now red + floated to top)."
            )
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Diff the combined Summer School Roster vs a past Drive revision or the source of truth.",
        epilog=(
            "examples:\n"
            "  python summer_roster_diff.py --since-rev 3286\n"
            "  python summer_roster_diff.py --since-date 2026-06-03 --csv _scratch_added.csv\n"
            "  python summer_roster_diff.py --vs-source\n"
            "  python summer_roster_diff.py --since-date 2026-06-03 --mark-red"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument(
        "--since-rev",
        metavar="ID",
        help="diff current roster vs this Drive revision id",
    )
    g.add_argument(
        "--since-date",
        metavar="YYYY-MM-DD",
        help="diff vs the last revision before this date",
    )
    g.add_argument(
        "--vs-source",
        action="store_true",
        help="compare the visible roster vs the source (ISR MAP Roster TRUE); exit 1 on drift",
    )
    ap.add_argument(
        "--mark-red",
        action="store_true",
        help="append added-not-already-red emails to _Highlight (requires --since-rev/--since-date)",
    )
    ap.add_argument(
        "--csv",
        metavar="PATH",
        help="also write added+removed to this CSV (keep gitignored)",
    )
    args = ap.parse_args()
    if args.mark_red and args.vs_source:
        ap.error("--mark-red cannot be combined with --vs-source")
    if (args.csv or args.mark_red) and args.vs_source:
        ap.error("--csv / --mark-red apply only to --since-rev / --since-date")

    sheets, drive, session = build_services()
    cur = current_roster(sheets)
    if args.vs_source:
        return run_vs_source(sheets, cur)
    return run_diff(sheets, drive, session, cur, args)


if __name__ == "__main__":
    sys.exit(main())
