"""outlook_write_file — the one non-Outlook write tool."""

from __future__ import annotations

import os
from typing import Any

from ..errors import ValidationError
from ..models.inputs import WriteFileInput


def register(mcp: Any) -> None:
    @mcp.tool(
        name="outlook_write_file",
        description=(
            "Write UTF-8 text to an absolute file path. Used by the agent to write its "
            "final markdown extraction. By default, refuses to overwrite existing "
            "files; set overwrite=true to replace."
        ),
    )
    async def outlook_write_file(
        path: str, content: str, overwrite: bool = False
    ) -> dict[str, Any]:
        params = WriteFileInput(path=path, content=content, overwrite=overwrite)
        if not os.path.isabs(params.path):
            raise ValidationError("path must be absolute.")
        if os.path.exists(params.path) and not params.overwrite:
            raise ValidationError(
                f"Refusing to overwrite existing file {params.path!r}; pass overwrite=true to replace."
            )
        os.makedirs(os.path.dirname(params.path) or ".", exist_ok=True)
        with open(params.path, "w", encoding="utf-8") as fh:
            fh.write(params.content)
        return {"path": params.path, "bytes_written": len(params.content.encode("utf-8"))}
