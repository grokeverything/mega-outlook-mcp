"""outlook_list_folders."""

from __future__ import annotations

from dataclasses import asdict
from typing import Any

from ..backends.base import Backend
from ..models.inputs import ListFoldersInput


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_list_folders",
        description="List all mail folders visible to the logged-in Outlook profile.",
    )
    async def outlook_list_folders(include_subfolders: bool = True) -> dict[str, Any]:
        params = ListFoldersInput(include_subfolders=include_subfolders)
        folders = await backend.list_folders(params.include_subfolders)
        return {"folders": [asdict(f) for f in folders]}
