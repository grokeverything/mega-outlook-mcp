"""outlook_list_calendar_events, outlook_get_calendar_event."""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime
from typing import Any

from ..backends.base import Backend, CalendarEvent
from ..models.inputs import GetCalendarEventInput, ListCalendarEventsInput


def _event_to_json(e: CalendarEvent) -> dict[str, Any]:
    raw = asdict(e)
    for key in ("start_utc", "end_utc"):
        value = raw.get(key)
        if isinstance(value, datetime):
            raw[key] = value.isoformat()
    return raw


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_list_calendar_events",
        description="List calendar events whose start falls within [start_utc, end_utc).",
    )
    async def outlook_list_calendar_events(
        start_utc: str,
        end_utc: str,
        calendar_id: str | None = None,
        max_results: int = 200,
    ) -> dict[str, Any]:
        params = ListCalendarEventsInput(
            start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc),
            calendar_id=calendar_id,
            max_results=max_results,
        )
        events = await backend.list_calendar_events(
            params.start_utc, params.end_utc, params.calendar_id, params.max_results
        )
        return {"events": [_event_to_json(e) for e in events]}

    @mcp.tool(
        name="outlook_get_calendar_event",
        description="Return a single calendar event including attendees and body.",
    )
    async def outlook_get_calendar_event(event_id: str) -> dict[str, Any]:
        params = GetCalendarEventInput(event_id=event_id)
        event = await backend.get_calendar_event(params.event_id)
        return _event_to_json(event)
