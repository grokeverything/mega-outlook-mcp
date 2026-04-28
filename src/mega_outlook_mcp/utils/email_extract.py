"""Backend-agnostic helpers for building EmailSummary / EmailFullMetadata."""

from __future__ import annotations


def preview_text(plain_body: str, length: int = 200) -> str:
    text = " ".join((plain_body or "").split())
    if len(text) <= length:
        return text
    return text[: length - 1].rstrip() + "…"


def detect_importance(value: object) -> str:
    from ..constants import IMPORTANCE_MAP

    try:
        idx = int(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return "normal"
    return IMPORTANCE_MAP.get(idx, "normal")


def domain_of(smtp: str) -> str:
    if not smtp or "@" not in smtp:
        return ""
    return smtp.rsplit("@", 1)[-1].lower()
