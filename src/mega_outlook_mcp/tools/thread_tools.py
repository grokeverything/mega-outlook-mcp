"""outlook_get_conversation_thread, outlook_get_thread_metadata."""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime, timezone
from typing import Any

from ..backends.base import Backend, EmailSummary
from ..models.inputs import GetConversationThreadInput, GetThreadMetadataInput
from ..utils.email_extract import domain_of
from ..utils.subject_utils import normalize_subject


def _summary_to_json(s: EmailSummary) -> dict[str, Any]:
    raw = asdict(s)
    for key in ("received_utc", "sent_utc"):
        value = raw.get(key)
        if isinstance(value, datetime):
            raw[key] = value.isoformat()
    return raw


_IMPORTANCE_RANK = {"low": 0, "normal": 1, "high": 2}


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_get_conversation_thread",
        description="Return every message in a conversation, sorted oldest → newest.",
    )
    async def outlook_get_conversation_thread(
        conversation_id: str,
        conversation_topic: str | None = None,
        max_messages: int = 200,
    ) -> dict[str, Any]:
        params = GetConversationThreadInput(
            conversation_id=conversation_id,
            conversation_topic=conversation_topic,
            max_messages=max_messages,
        )
        msgs = await backend.get_conversation_thread(
            params.conversation_id, params.conversation_topic, params.max_messages
        )
        return {"messages": [_summary_to_json(m) for m in msgs]}

    @mcp.tool(
        name="outlook_get_thread_metadata",
        description=(
            "Composite tool: return the `thread-meta` YAML block for a conversation. "
            "Includes normalized subject, participants, escalated importance, and "
            "whether the thread is internal-only."
        ),
    )
    async def outlook_get_thread_metadata(
        conversation_id: str,
        conversation_topic: str | None = None,
        mailbox_domain: str | None = None,
    ) -> dict[str, Any]:
        params = GetThreadMetadataInput(
            conversation_id=conversation_id,
            conversation_topic=conversation_topic,
            mailbox_domain=mailbox_domain,
        )
        msgs = await backend.get_conversation_thread(
            params.conversation_id, params.conversation_topic, max_messages=1000
        )

        if params.mailbox_domain:
            owner_domain = params.mailbox_domain.lower()
        else:
            info = await backend.get_mailbox_info()
            owner_domain = info.domain.lower()

        participants: set[str] = set()
        domains: set[str] = set()
        unread = 0
        highest_imp = 0
        first_ts = None
        last_ts = None
        normalized = ""

        for m in msgs:
            if m.sender_smtp:
                participants.add(m.sender_smtp.lower())
                domains.add(domain_of(m.sender_smtp))
            for addr in (*m.to_smtp, *m.cc_smtp):
                participants.add(addr.lower())
                domains.add(domain_of(addr))
            if not m.is_read:
                unread += 1
            rank = _IMPORTANCE_RANK.get(m.importance, 1)
            if rank > highest_imp:
                highest_imp = rank
            ts = m.sent_utc or m.received_utc
            if ts is not None:
                if first_ts is None or ts < first_ts:
                    first_ts = ts
                if last_ts is None or ts > last_ts:
                    last_ts = ts
            if not normalized and m.subject:
                normalized = normalize_subject(m.subject)

        highest_importance = next(
            (k for k, v in _IMPORTANCE_RANK.items() if v == highest_imp), "normal"
        )
        internal_only = (
            bool(domains)
            and all(d == owner_domain for d in domains if d)
        )

        def _iso(dt: datetime | None) -> str | None:
            if dt is None:
                return None
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.isoformat()

        return {
            "thread_meta": {
                "conversation_id": params.conversation_id,
                "normalized_subject": normalized,
                "message_count": len(msgs),
                "unread_count": unread,
                "participants_smtp": sorted(participants),
                "participant_domains": sorted(d for d in domains if d),
                "first_message_utc": _iso(first_ts),
                "last_message_utc": _iso(last_ts),
                "escalated_importance": highest_importance,
                "internal_only": internal_only,
            }
        }
