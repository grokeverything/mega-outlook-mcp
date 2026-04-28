"""outlook_get_email_full_metadata."""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime
from typing import Any

from ..backends.base import Backend
from ..models.inputs import GetEmailFullMetadataInput


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_get_email_full_metadata",
        description=(
            "Return deep per-email metadata: Message-ID, In-Reply-To, References, "
            "full transport headers, and delegation info. On macOS, fields that "
            "cannot be resolved are returned as 'ERROR-MAC-Support-Unavailable'."
        ),
    )
    async def outlook_get_email_full_metadata(entry_id: str) -> dict[str, Any]:
        params = GetEmailFullMetadataInput(entry_id=entry_id)
        result = await backend.get_email_full_metadata(params.entry_id)
        raw = asdict(result)
        summary = raw.get("summary", {})
        for key in ("received_utc", "sent_utc"):
            value = summary.get(key)
            if isinstance(value, datetime):
                summary[key] = value.isoformat()
        return raw
