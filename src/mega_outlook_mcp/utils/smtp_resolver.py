"""Resolve SMTP addresses from Outlook address entries.

Exchange addresses can come through as a distinguished name
(`/O=ExchangeLabs/...`) instead of SMTP. This resolver walks the usual
fallback chain:

1. `PR_SMTP_ADDRESS` via the PropertyAccessor
2. `GetExchangeUser().PrimarySmtpAddress`
3. `Recipient.Address` (only if it already looks like SMTP)

Mac address lookups skip this resolver entirely — AppleScript returns SMTP.
"""

from __future__ import annotations

from typing import Any

from .mapi_props import PR_SMTP_ADDRESS, safe_get


def _looks_like_smtp(value: str | None) -> bool:
    return bool(value) and "@" in value and "/" not in value


def resolve_address_entry(entry: Any) -> str:
    """Best-effort SMTP extraction from an Outlook AddressEntry."""
    smtp = safe_get(entry, PR_SMTP_ADDRESS)
    if _looks_like_smtp(smtp):
        return smtp
    try:
        user = entry.GetExchangeUser()
        if user is not None and _looks_like_smtp(user.PrimarySmtpAddress):
            return user.PrimarySmtpAddress
    except Exception:
        pass
    try:
        addr = entry.Address
        if _looks_like_smtp(addr):
            return addr
    except Exception:
        pass
    return ""


def resolve_recipient(recipient: Any) -> str:
    """Resolve an Outlook Recipient to SMTP."""
    try:
        entry = recipient.AddressEntry
    except Exception:
        entry = None
    if entry is not None:
        smtp = resolve_address_entry(entry)
        if smtp:
            return smtp
    try:
        return recipient.Address if _looks_like_smtp(recipient.Address) else ""
    except Exception:
        return ""


def resolve_sender(mail_item: Any) -> str:
    """Resolve the sender of a MailItem to SMTP."""
    try:
        entry = mail_item.Sender
    except Exception:
        entry = None
    if entry is not None:
        smtp = resolve_address_entry(entry)
        if smtp:
            return smtp
    # Fallback: SenderEmailAddress
    try:
        addr = mail_item.SenderEmailAddress
        if _looks_like_smtp(addr):
            return addr
    except Exception:
        pass
    return ""
