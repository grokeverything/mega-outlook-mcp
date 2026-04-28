"""Phase 4 (Exchange-specific) tool handlers.

Most tools require Exchange/Microsoft 365 + the Windows COM backend. The
macOS backend returns the MAC_UNAVAILABLE sentinel for fields it cannot
reach.
"""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime
from typing import Any

from ..backends.base import (
    Backend,
    FreeBusyResponse,
    OutOfOfficeStatus,
)
from ..models.inputs import (
    CalendarFreeBusyInput,
    GalSearchInput,
    GetSignatureInput,
    MeetingRoomFinderInput,
    SetOutOfOfficeInput,
    SetSignatureInput,
    ToggleRuleInput,
)


def _freebusy_to_json(fb: FreeBusyResponse) -> dict[str, Any]:
    return {
        "smtp": fb.smtp,
        "slots": [
            {
                "start_utc": s.start_utc.isoformat(),
                "end_utc": s.end_utc.isoformat(),
                "status": s.status,
            }
            for s in fb.slots
        ],
    }


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(name="outlook_get_out_of_office", description="Read the current automatic-replies (OOO) state.")
    async def outlook_get_out_of_office() -> dict[str, Any]:
        result = await backend.get_out_of_office()
        raw = asdict(result)
        for key in ("start_utc", "end_utc"):
            value = raw.get(key)
            if isinstance(value, datetime):
                raw[key] = value.isoformat()
        return raw

    @mcp.tool(
        name="outlook_set_out_of_office",
        description="Enable/disable automatic replies and set internal/external messages.",
    )
    async def outlook_set_out_of_office(
        enabled: bool,
        internal_message: str = "",
        external_message: str = "",
        start_utc: str | None = None,
        end_utc: str | None = None,
        external_audience: str = "all",
    ) -> dict[str, Any]:
        params = SetOutOfOfficeInput(
            enabled=enabled, internal_message=internal_message,
            external_message=external_message,
            start_utc=datetime.fromisoformat(start_utc) if start_utc else None,
            end_utc=datetime.fromisoformat(end_utc) if end_utc else None,
            external_audience=external_audience,  # type: ignore[arg-type]
        )
        status = OutOfOfficeStatus(
            enabled=params.enabled,
            internal_message=params.internal_message,
            external_message=params.external_message,
            start_utc=params.start_utc,
            end_utc=params.end_utc,
            external_audience=params.external_audience,
        )
        return asdict(await backend.set_out_of_office(status))

    @mcp.tool(name="outlook_get_signature", description="Read the configured email signature for an account (or default).")
    async def outlook_get_signature(account_id: str | None = None) -> dict[str, Any]:
        params = GetSignatureInput(account_id=account_id)
        return asdict(await backend.get_signature(params.account_id))

    @mcp.tool(name="outlook_set_signature", description="Replace the email signature.")
    async def outlook_set_signature(
        body_html: str, body_plain: str = "", account_id: str | None = None
    ) -> dict[str, Any]:
        params = SetSignatureInput(body_html=body_html, body_plain=body_plain, account_id=account_id)
        return asdict(await backend.set_signature(params.account_id, params.body_html, params.body_plain))

    @mcp.tool(name="outlook_list_rules", description="List inbox rules. Names + enabled state.")
    async def outlook_list_rules() -> dict[str, Any]:
        rules = await backend.list_rules()
        return {"rules": [asdict(r) for r in rules]}

    @mcp.tool(name="outlook_toggle_rule", description="Enable or disable a rule by id (or name).")
    async def outlook_toggle_rule(rule_id: str, enabled: bool = True) -> dict[str, Any]:
        params = ToggleRuleInput(rule_id=rule_id, enabled=enabled)
        return asdict(await backend.toggle_rule(params.rule_id, params.enabled))

    @mcp.tool(
        name="outlook_calendar_freebusy",
        description=(
            "Free/busy lookup for one or more SMTP addresses across a window. "
            "Slots are returned at the requested granularity. Mac returns empty slots."
        ),
    )
    async def outlook_calendar_freebusy(
        smtps: list[str], start_utc: str, end_utc: str, slot_minutes: int = 30
    ) -> dict[str, Any]:
        params = CalendarFreeBusyInput(
            smtps=smtps,
            start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc),
            slot_minutes=slot_minutes,
        )
        results = await backend.calendar_freebusy(
            params.smtps, params.start_utc, params.end_utc, params.slot_minutes
        )
        return {"freebusy": [_freebusy_to_json(r) for r in results]}

    @mcp.tool(
        name="outlook_meeting_room_finder",
        description=(
            "Find conference rooms free in a window. Capacity hints which to prefer. "
            "Mac returns []. Windows uses the Resource address list + free/busy."
        ),
    )
    async def outlook_meeting_room_finder(
        start_utc: str, end_utc: str, capacity: int = 10, location_hint: str | None = None
    ) -> dict[str, Any]:
        params = MeetingRoomFinderInput(
            start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc),
            capacity=capacity, location_hint=location_hint,
        )
        rooms = await backend.meeting_room_finder(
            params.start_utc, params.end_utc, params.capacity, params.location_hint
        )
        return {"rooms": [asdict(r) for r in rooms]}

    @mcp.tool(
        name="outlook_gal_search",
        description=(
            "Search the Exchange Global Address List (GAL). Mac falls back to local "
            "contacts search since AppleScript has no GAL access."
        ),
    )
    async def outlook_gal_search(query: str, limit: int = 25) -> dict[str, Any]:
        params = GalSearchInput(query=query, limit=limit)
        results = await backend.gal_search(params.query, params.limit)
        return {"results": [asdict(c) for c in results]}

    @mcp.tool(
        name="outlook_list_delegated_mailboxes",
        description=(
            "List shared/delegated mailboxes you have access to via Exchange "
            "delegation. Mac returns []."
        ),
    )
    async def outlook_list_delegated_mailboxes() -> dict[str, Any]:
        results = await backend.list_delegated_mailboxes()
        return {"mailboxes": [asdict(m) for m in results]}

    @mcp.tool(
        name="outlook_list_public_folders",
        description="List Exchange public folders attached to this profile. Mac returns [].",
    )
    async def outlook_list_public_folders() -> dict[str, Any]:
        results = await backend.list_public_folders()
        return {"folders": [asdict(f) for f in results]}

    @mcp.tool(
        name="outlook_get_mailbox_quota",
        description="Mailbox usage in bytes plus the warning / prohibit-send thresholds.",
    )
    async def outlook_get_mailbox_quota() -> dict[str, Any]:
        return asdict(await backend.get_mailbox_quota())
