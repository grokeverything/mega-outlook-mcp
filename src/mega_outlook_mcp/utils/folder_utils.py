"""Shared folder-name resolution helpers.

Each backend implements its own `_resolve_folder`; this module only holds
name-normalization helpers so the two implementations stay consistent.
"""

from __future__ import annotations

_DEFAULT_ALIASES = {
    "inbox": "Inbox",
    "sent": "Sent Items",
    "sent items": "Sent Items",
    "sent mail": "Sent Items",
    "drafts": "Drafts",
    "deleted": "Deleted Items",
    "deleted items": "Deleted Items",
    "trash": "Deleted Items",
    "junk": "Junk Email",
    "junk email": "Junk Email",
    "spam": "Junk Email",
    "outbox": "Outbox",
    "archive": "Archive",
}


def canonicalize_folder_name(name: str) -> str:
    """Map casual folder names ('sent', 'trash') to canonical Outlook names."""
    key = (name or "").strip().lower()
    return _DEFAULT_ALIASES.get(key, name.strip())


def split_folder_path(path: str) -> list[str]:
    """Split `Inbox/Projects/Q3` into `['Inbox', 'Projects', 'Q3']`."""
    return [p for p in (path or "").split("/") if p]
