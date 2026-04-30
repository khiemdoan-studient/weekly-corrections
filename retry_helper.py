"""retry_helper.py — Shared retry helper for transient Google API errors.

Replaces the per-file `_retry_api` / `_retry` helpers that were too tight
(3 attempts, linear backoff) for sustained API outages.

The 2026-04-29 incident showed why: a transient `HttpError 500` followed
by `TimeoutError` on the same endpoint can span ≥3 minutes. The old
3-attempt × linear-backoff helper exhausted retries inside that window
and crashed the hourly pipeline. The next hourly run self-recovered, but
that data window was empty for IMs.

Strategy
--------
- 5 attempts (vs old 3) — covers ~5 min of API hiccups
- Exponential backoff: 1s, 2s, 4s, 8s, 16s = ~31s of sleeps
- 25% random jitter on each sleep — avoids synchronized retries when
  multiple workflows hit the same outage
- Transient-only catch: HttpError 5xx/429/408 + TimeoutError +
  socket.timeout + ConnectionError. Programming bugs (KeyError, etc.)
  raise immediately instead of being masked by retries.

Usage
-----
    from retry_helper import retry_api

    resp = retry_api(
        lambda: sheets.spreadsheets()
        .get(spreadsheetId=sid, fields="sheets.properties")
        .execute(),
        label="get sheet properties",  # optional, for log clarity
    )
"""

import random
import socket
import time

from googleapiclient.errors import HttpError

# HTTP status codes that should trigger a retry.
# 408 Request Timeout, 429 Too Many Requests, 500/502/503/504 server errors.
TRANSIENT_HTTP_STATUSES = {408, 429, 500, 502, 503, 504}

# Connection-level exceptions that warrant retry.
TRANSIENT_EXCEPTIONS = (TimeoutError, socket.timeout, ConnectionError)

DEFAULT_MAX_ATTEMPTS = 5
DEFAULT_BASE_DELAY = 1.0
DEFAULT_MAX_DELAY = 30.0


def _is_transient(exc):
    """Return True if exc is a transient API error worth retrying."""
    if isinstance(exc, TRANSIENT_EXCEPTIONS):
        return True
    if isinstance(exc, HttpError):
        try:
            return int(exc.resp.status) in TRANSIENT_HTTP_STATUSES
        except (AttributeError, ValueError):
            return False
    return False


def _summarize(exc):
    """Short label for an exception, suitable for retry logs."""
    if isinstance(exc, HttpError):
        try:
            return f"HttpError {exc.resp.status}"
        except AttributeError:
            return "HttpError"
    return type(exc).__name__


def retry_api(
    fn,
    max_attempts=DEFAULT_MAX_ATTEMPTS,
    base_delay=DEFAULT_BASE_DELAY,
    max_delay=DEFAULT_MAX_DELAY,
    label="",
):
    """Run fn() with exponential backoff + jitter on transient API errors.

    Parameters
    ----------
    fn : callable
        Zero-arg callable wrapping the API call (e.g. `lambda: req.execute()`).
    max_attempts : int
        Total attempts including the first try. Default 5.
    base_delay : float
        Seconds for the first retry sleep. Doubles each subsequent retry.
        Default 1.0.
    max_delay : float
        Cap per-attempt sleep at this many seconds. Default 30.0.
    label : str
        Optional human-readable name for the operation. Appears in retry
        log messages so you can tell which call site is retrying.

    Returns
    -------
    Whatever fn() returns.

    Raises
    ------
    The last transient exception if max_attempts is exhausted, OR any
    non-transient exception immediately on first encounter (no retries).
    """
    last_exc = None
    for attempt in range(max_attempts):
        try:
            return fn()
        except Exception as e:
            if not _is_transient(e):
                # Permanent error — fail fast, don't waste retry budget.
                raise
            last_exc = e
            if attempt == max_attempts - 1:
                break  # exhausted; raise outside the loop
            wait = min(base_delay * (2**attempt), max_delay)
            wait += random.uniform(0, wait * 0.25)  # 0-25% jitter
            label_str = f" [{label}]" if label else ""
            print(
                f"     [retry{label_str}] attempt {attempt + 1}/{max_attempts} "
                f"failed ({_summarize(e)}); waiting {wait:.1f}s"
            )
            time.sleep(wait)
    # Exhausted retries — raise the final transient exception.
    raise last_exc
