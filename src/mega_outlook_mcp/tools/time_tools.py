"""outlook_get_current_time."""

from __future__ import annotations

from typing import Any

from ..models.inputs import GetCurrentTimeInput
from ..utils.time_utils import lookback_window, now_local, now_utc


def register(mcp: Any) -> None:
    @mcp.tool(
        name="outlook_get_current_time",
        description=(
            "Return the current time plus a default lookback window. Use this before "
            "calling outlook_get_emails_in_time_range so the window is grounded in the "
            "server's clock."
        ),
    )
    async def outlook_get_current_time(lookback_minutes: int = 65) -> dict[str, Any]:
        params = GetCurrentTimeInput(lookback_minutes=lookback_minutes)
        start, end = lookback_window(params.lookback_minutes)
        local = now_local()
        return {
            "now_utc": now_utc().isoformat(),
            "now_local": local.isoformat(),
            "local_tz": str(local.tzinfo),
            "lookback_window": {
                "start_utc": start.isoformat(),
                "end_utc": end.isoformat(),
                "minutes": params.lookback_minutes,
            },
        }
