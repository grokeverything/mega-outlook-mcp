"""outlook_get_emails_in_time_range, outlook_search_emails."""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime
from typing import Any

from ..backends.base import Backend, EmailSummary
from ..models.inputs import GetEmailsInTimeRangeInput, SearchEmailsInput


def _summary_to_json(s: EmailSummary) -> dict[str, Any]:
    raw = asdict(s)
    for key in ("received_utc", "sent_utc"):
        value = raw.get(key)
        if isinstance(value, datetime):
            raw[key] = value.isoformat()
    return raw


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_get_emails_in_time_range",
        description=(
            "Return emails received or sent within [start_utc, end_utc). Folders default "
            "to Inbox and Sent Items. This is the primary extraction tool."
        ),
    )
    async def outlook_get_emails_in_time_range(
        start_utc: str,
        end_utc: str,
        folders: list[str] | None = None,
        max_results: int = 500,
    ) -> dict[str, Any]:
        params = GetEmailsInTimeRangeInput(
            start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc),
            folders=folders or ["Inbox", "Sent Items"],
            max_results=max_results,
        )
        results = await backend.get_emails_in_time_range(
            params.start_utc, params.end_utc, params.folders, params.max_results
        )
        return {"emails": [_summary_to_json(r) for r in results]}

    @mcp.tool(
        name="outlook_search_emails",
        description="Search emails by subject, sender, or body text.",
    )
    async def outlook_search_emails(
        query: str,
        folders: list[str] | None = None,
        field: str = "any",
        max_results: int = 100,
    ) -> dict[str, Any]:
        params = SearchEmailsInput(
            query=query,
            folders=folders or ["Inbox"],
            field=field,  # type: ignore[arg-type]
            max_results=max_results,
        )
        results = await backend.search_emails(
            params.query, params.folders, params.field, params.max_results
        )
        return {"emails": [_summary_to_json(r) for r in results]}
