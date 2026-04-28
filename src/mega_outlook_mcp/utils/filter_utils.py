"""Outlook Restrict/DASL filter string builders (Windows)."""

from __future__ import annotations

from datetime import datetime

from .time_utils import outlook_restrict_format


def escape_restrict(value: str) -> str:
    """Escape a value for inclusion in an Outlook Restrict string."""
    return value.replace("'", "''")


def escape_dasl(value: str) -> str:
    """Escape a value for inclusion in a DASL Items.Restrict query."""
    return value.replace("'", "''").replace('"', '""')


def build_time_range_restrict(
    field: str, start: datetime, end: datetime
) -> str:
    start_str = outlook_restrict_format(start)
    end_str = outlook_restrict_format(end)
    return f"[{field}] >= '{start_str}' AND [{field}] < '{end_str}'"
