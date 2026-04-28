"""outlook_get_mailbox_info."""

from __future__ import annotations

from dataclasses import asdict
from typing import Any

from ..backends.base import Backend
from ..models.inputs import GetMailboxInfoInput


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_get_mailbox_info",
        description=(
            "Return the mailbox owner's display name, SMTP address, and domain. Used "
            "to classify internal vs external participants."
        ),
    )
    async def outlook_get_mailbox_info() -> dict[str, Any]:
        GetMailboxInfoInput()  # validation: no inputs
        info = await backend.get_mailbox_info()
        return asdict(info)
