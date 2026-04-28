"""outlook_save_attachment."""

from __future__ import annotations

from dataclasses import asdict
from typing import Any

from ..backends.base import Backend
from ..models.inputs import SaveAttachmentInput


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_save_attachment",
        description=(
            "Save a real (non-inline) attachment from an email to an absolute path on "
            "disk. Refuses to save inline signature images."
        ),
    )
    async def outlook_save_attachment(
        entry_id: str, attachment_index: int, save_path: str
    ) -> dict[str, Any]:
        params = SaveAttachmentInput(
            entry_id=entry_id,
            attachment_index=attachment_index,
            save_path=save_path,
        )
        info = await backend.save_attachment(
            params.entry_id, params.attachment_index, params.save_path
        )
        return {"saved": asdict(info), "path": params.save_path}
