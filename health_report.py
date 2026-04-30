"""health_report.py — Pipeline health summary.

Queries recent GitHub Actions workflow runs via the gh CLI and prints a
Markdown summary suitable for posting as a tracking Issue or reading
locally.

Output covers both pipelines (hourly-pipeline.yml + weekly-snapshot.yml):
- success rate over the window
- failure count + max consecutive failure streak
- last failure timestamp
- median run duration

Usage:
    python health_report.py                       # last 30 days, stdout
    python health_report.py --days 7               # last 7 days
    python health_report.py --output /tmp/h.md     # write to file
    python health_report.py --repo OWNER/NAME      # custom repo

Requires the `gh` CLI authenticated against the repo.
"""

import argparse
import json
import statistics
import subprocess
import sys
from collections import Counter
from datetime import datetime, timedelta, timezone

DEFAULT_REPO = "khiemdoan-studient/weekly-corrections"
WORKFLOWS = ["hourly-pipeline.yml", "weekly-snapshot.yml"]

# How many runs to fetch per workflow. Hourly cron = ~720/month, so 1000 is
# enough for a 30-day window. Bump if days >> 30.
PER_PAGE = 1000


def fetch_runs(repo, workflow, since):
    """Fetch workflow runs newer than `since` (datetime). Returns list of dicts."""
    cmd = [
        "gh",
        "run",
        "list",
        "--repo",
        repo,
        "--workflow",
        workflow,
        "--limit",
        str(PER_PAGE),
        "--json",
        "databaseId,createdAt,updatedAt,conclusion,status,event,workflowName",
    ]
    try:
        out = subprocess.check_output(cmd, stderr=subprocess.PIPE).decode()
    except subprocess.CalledProcessError as e:
        sys.stderr.write(f"gh CLI failed for {workflow}: {e.stderr.decode()}\n")
        return []
    runs = json.loads(out)
    cutoff = since.isoformat()
    return [r for r in runs if r["createdAt"] >= cutoff]


def parse_iso(ts):
    """Parse a GH-format ISO timestamp; returns timezone-aware datetime."""
    # gh returns "2026-04-29T08:09:19Z"
    return datetime.fromisoformat(ts.replace("Z", "+00:00"))


def consecutive_failure_streak(runs):
    """Return the max length of consecutive failures across runs.

    `runs` is expected sorted by createdAt ascending.
    """
    max_streak = 0
    current = 0
    for r in runs:
        if r["conclusion"] == "failure":
            current += 1
            max_streak = max(max_streak, current)
        elif r["conclusion"] == "success":
            current = 0
    return max_streak


def current_consecutive_failures(runs):
    """Return the number of consecutive failures at the END of the window.

    Useful for the smart-notify threshold check. `runs` sorted asc.
    """
    streak = 0
    for r in reversed(runs):
        if r["conclusion"] == "failure":
            streak += 1
        elif r["conclusion"] == "success":
            break
        # ignore cancelled / null
    return streak


def median_duration_seconds(runs):
    """Return median (createdAt → updatedAt) duration across runs in seconds."""
    durations = []
    for r in runs:
        try:
            start = parse_iso(r["createdAt"])
            end = parse_iso(r["updatedAt"])
            durations.append((end - start).total_seconds())
        except Exception:
            continue
    if not durations:
        return None
    return statistics.median(durations)


def summarize(runs, workflow_name, days):
    """Build a Markdown section for one workflow."""
    if not runs:
        return f"### {workflow_name}\n\n_No runs in the last {days} days._\n"

    runs_sorted = sorted(runs, key=lambda r: r["createdAt"])
    n = len(runs_sorted)
    by_concl = Counter(r["conclusion"] for r in runs_sorted)
    success = by_concl.get("success", 0)
    failure = by_concl.get("failure", 0)
    cancelled = by_concl.get("cancelled", 0)
    other = n - success - failure - cancelled
    rate = (success / n) * 100 if n > 0 else 0

    max_streak = consecutive_failure_streak(runs_sorted)
    current_streak = current_consecutive_failures(runs_sorted)

    last_failure = next(
        (r for r in reversed(runs_sorted) if r["conclusion"] == "failure"),
        None,
    )
    median_dur = median_duration_seconds(runs_sorted)

    lines = [
        f"### {workflow_name}",
        "",
        f"- **Total runs**: {n}",
        f"- **Success rate**: {rate:.1f}% ({success} / {n})",
        f"- **Failures**: {failure}",
    ]
    if cancelled:
        lines.append(f"- **Cancelled**: {cancelled}")
    if other:
        lines.append(f"- **Other (in-progress / unknown)**: {other}")
    lines.append(f"- **Max consecutive failure streak**: {max_streak}")
    if current_streak > 0:
        lines.append(
            f"- **Currently failing**: {current_streak} consecutive failures "
            f"(triggers smart-notify if >=3)"
        )
    if last_failure:
        lines.append(f"- **Last failure**: {last_failure['createdAt']}")
    if median_dur is not None:
        lines.append(f"- **Median run duration**: {median_dur:.1f}s")
    lines.append("")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--days", type=int, default=30, help="Lookback window in days")
    parser.add_argument("--repo", default=DEFAULT_REPO, help="GitHub repo owner/name")
    parser.add_argument(
        "--output", default=None, help="Write to file instead of stdout"
    )
    args = parser.parse_args()

    since = datetime.now(timezone.utc) - timedelta(days=args.days)
    today = datetime.now(timezone.utc).date().isoformat()

    sections = [f"# Pipeline Health Report — {today}", ""]
    sections.append(f"_Last {args.days} days. Repo: `{args.repo}`._")
    sections.append("")

    for wf in WORKFLOWS:
        runs = fetch_runs(args.repo, wf, since)
        sections.append(summarize(runs, wf, args.days))

    body = "\n".join(sections).rstrip() + "\n"

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(body)
        print(f"Wrote {len(body)} bytes to {args.output}")
    else:
        print(body)


if __name__ == "__main__":
    main()
