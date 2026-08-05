"""outlook_save_attachment."""

from __future__ import annotations

import os
from dataclasses import asdict
from typing import Any

from ..backends.base import Backend
from ..errors import ValidationError
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
        # AppleScript's `POSIX file` (and COM's SaveAsFile) won't expand `~`
        # or create missing directories; both fail with opaque errors.
        resolved = os.path.expanduser(params.save_path)
        if not os.path.isabs(resolved):
            raise ValidationError(
                f"save_path must be absolute (got {params.save_path!r})."
            )
        os.makedirs(os.path.dirname(resolved), exist_ok=True)
        info = await backend.save_attachment(
            params.entry_id, params.attachment_index, resolved
        )
        return {"saved": asdict(info), "path": resolved}
