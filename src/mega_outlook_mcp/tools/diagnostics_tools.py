"""outlook_diagnostics — self-test against the baseline manifest."""

from __future__ import annotations

from dataclasses import asdict
from typing import Any

from ..backends.base import Backend


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_diagnostics",
        description=(
            "Self-test the MCP against the currently-installed Outlook. Probes every "
            "field this server depends on against a sample message (or auto-picks the "
            "newest Inbox item) and reports HEALTHY / DEGRADED:<n>_fields / BROKEN with "
            "the list of affected tools. Run this after Outlook updates."
        ),
    )
    async def outlook_diagnostics(sample_message_id: str | None = None) -> dict[str, Any]:
        result = await backend.diagnostics(sample_message_id)
        return asdict(result)
