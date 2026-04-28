"""Tool registration. Each module exposes one or more tools via @mcp.tool."""

from __future__ import annotations

from typing import Any

from ..backends.base import Backend
from . import (
    attachment_tools,
    calendar_tools,
    composite_tools,
    contact_tools,
    diagnostics_tools,
    email_tools,
    exchange_tools,
    file_tools,
    folder_tools,
    mailbox_tools,
    metadata_tools,
    parity_tools,
    thread_tools,
    time_tools,
)


def register_all(mcp: Any, backend: Backend) -> None:
    time_tools.register(mcp)
    folder_tools.register(mcp, backend)
    email_tools.register(mcp, backend)
    thread_tools.register(mcp, backend)
    metadata_tools.register(mcp, backend)
    attachment_tools.register(mcp, backend)
    mailbox_tools.register(mcp, backend)
    file_tools.register(mcp)
    calendar_tools.register(mcp, backend)
    contact_tools.register(mcp, backend)
    diagnostics_tools.register(mcp, backend)
    composite_tools.register(mcp, backend)
    parity_tools.register(mcp, backend)
    exchange_tools.register(mcp, backend)


__all__ = ["register_all"]
