"""Escape helpers for AppleScript strings, folder names, and date literals."""

from __future__ import annotations

from datetime import datetime


def as_str(value: str) -> str:
    """Quote a Python string for embedding as an AppleScript string literal."""
    # AppleScript escapes: backslash and double-quote.
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def applescript_epoch_date(dt: datetime) -> str:
    """Return an AppleScript expression that evaluates to the given instant.

    Uses `do shell script "date -r <epoch> ..."` which is locale-stable on
    macOS and returns something AppleScript can coerce to a date via
    `(date X)` pattern — but the simpler approach is arithmetic from
    `current date`: we subtract `now`'s epoch and add the target's.
    """
    epoch = int(dt.timestamp())
    return (
        "((current date) - "
        "((do shell script \"date +%s\") as integer) "
        f"+ {epoch})"
    )
