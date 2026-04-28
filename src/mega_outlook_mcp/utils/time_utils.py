"""Time helpers: UTC/local conversion, 65-min window, AppleScript date literals."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

from ..constants import DEFAULT_LOOKBACK_MINUTES


def now_utc() -> datetime:
    return datetime.now(timezone.utc)


def now_local() -> datetime:
    return datetime.now().astimezone()


def lookback_window(minutes: int = DEFAULT_LOOKBACK_MINUTES) -> tuple[datetime, datetime]:
    """Return `(start_utc, end_utc)` with `end_utc` at now and `start_utc` earlier."""
    end = now_utc()
    start = end - timedelta(minutes=minutes)
    return start, end


def to_utc(dt: datetime) -> datetime:
    if dt.tzinfo is None:
        raise ValueError("Naive datetime passed where timezone-aware was required.")
    return dt.astimezone(timezone.utc)


def to_local(dt: datetime) -> datetime:
    if dt.tzinfo is None:
        raise ValueError("Naive datetime passed where timezone-aware was required.")
    return dt.astimezone()


def is_within_window(ts: datetime, start: datetime, end: datetime) -> bool:
    ts = to_utc(ts)
    return to_utc(start) <= ts < to_utc(end)


def outlook_restrict_format(dt: datetime) -> str:
    """Outlook `Restrict`/`Items.Find` demands local time in this format."""
    return to_local(dt).strftime("%m/%d/%Y %I:%M %p")


def applescript_date_literal(dt: datetime) -> str:
    """Locale-independent AppleScript date literal usable in `whose`-clauses.

    Returns an expression of the form
        (current date) - (current date) + (TIMESTAMP * seconds)
    where TIMESTAMP is a Unix epoch seconds value. This avoids the
    locale-dependence of `date "…"`.
    """
    epoch = int(to_utc(dt).timestamp())
    # AppleScript has no direct "epoch to date" but supports arithmetic on
    # dates. We anchor from the known AppleScript epoch (1904-01-01 UTC is
    # AppleScript's internal zero in some contexts; but `current date`
    # avoids that trap). Using (current date) minus its seconds-since-epoch
    # equivalent + the target seconds-since-epoch yields the target date.
    return (
        "((current date) - ((do shell script \"date +%s\") as integer) "
        f"+ {epoch})"
    )
