"""Parse Message-ID / In-Reply-To / References from a raw RFC 5322 message.

Used by the Mac backend to fill MAPI-equivalent header fields by exporting
a message's `source` via AppleScript and parsing the result with stdlib
`email.parser`.
"""

from __future__ import annotations

from email import message_from_bytes
from email.parser import BytesParser


def parse_headers(raw_source: bytes) -> dict[str, str | list[str]]:
    """Return a dict with keys: message_id, in_reply_to, references, headers_raw."""
    msg = BytesParser().parsebytes(raw_source, headersonly=True)
    references_field = msg.get("References")
    references_list = _split_msg_ids(references_field) if references_field else []
    return {
        "message_id": (msg.get("Message-ID") or "").strip() or None,
        "in_reply_to": (msg.get("In-Reply-To") or "").strip() or None,
        "references": references_list,
        "headers_raw": _serialize_headers(msg),
    }


def _split_msg_ids(value: str) -> list[str]:
    # Message-IDs are whitespace-separated, each wrapped in angle brackets.
    ids: list[str] = []
    buf = ""
    depth = 0
    for ch in value:
        if ch == "<":
            depth = 1
            buf = "<"
        elif ch == ">" and depth:
            buf += ">"
            ids.append(buf)
            buf = ""
            depth = 0
        elif depth:
            buf += ch
    return ids


def _serialize_headers(msg) -> str:  # type: ignore[no-untyped-def]
    lines = []
    for key, value in msg.items():
        lines.append(f"{key}: {value}")
    return "\n".join(lines)


__all__ = ["parse_headers", "message_from_bytes"]
