"""Tier C composites: agent-friendly aggregations over backend primitives.

Each tool calls existing backend methods and assembles a higher-level view.
No new backend methods are required for these.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import asdict
from datetime import datetime, timedelta, timezone
from typing import Any

from ..backends.base import Backend, EmailSummary
from ..models.inputs import (
    ExtractActionItemsInput,
    FindPromisedActionsInput,
    FindUnansweredInput,
    MeetingPrepInput,
    RelationshipGraphInput,
    SummarizeInboxInput,
    ThreadifyInput,
)
from ..utils.email_extract import domain_of
from ..utils.subject_utils import normalize_subject

_IMPORTANCE_RANK = {"low": 0, "normal": 1, "high": 2}

_ACTION_PHRASES = (
    "action required",
    "please review",
    "please approve",
    "please confirm",
    "please respond",
    "needs your attention",
    "by eod",
    "by end of day",
    "by end of week",
    "deadline",
    "urgent",
    "asap",
    "due ",
    "please complete",
    "kindly",
)

_PROMISE_PHRASES = (
    "i will get back",
    "i'll get back",
    "i will send",
    "i'll send",
    "i will follow up",
    "i'll follow up",
    "i will check",
    "i'll check",
    "i will review",
    "i'll review",
    "by eod",
    "by end of day",
    "by end of week",
    "circle back",
    "i'll have",
    "i will have",
)


def _summary_to_json(s: EmailSummary) -> dict[str, Any]:
    raw = asdict(s)
    for key in ("received_utc", "sent_utc"):
        value = raw.get(key)
        if isinstance(value, datetime):
            raw[key] = value.isoformat()
    return raw


def _ts(s: EmailSummary) -> datetime | None:
    return s.sent_utc or s.received_utc


def _utc(dt: datetime) -> datetime:
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_summarize_inbox",
        description=(
            "Composite. Bucket recent mail into priority tiers using [Action Required] "
            "tags, importance, sender domain (internal/external), and unread state. "
            "Returns counts plus the top items per bucket."
        ),
    )
    async def outlook_summarize_inbox(
        start_utc: str,
        end_utc: str,
        folders: list[str] | None = None,
        max_results: int = 300,
    ) -> dict[str, Any]:
        params = SummarizeInboxInput(
            start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc),
            folders=folders or ["Inbox"],
            max_results=max_results,
        )
        emails = await backend.get_emails_in_time_range(
            params.start_utc, params.end_utc, params.folders, params.max_results
        )
        info = await backend.get_mailbox_info()
        owner_domain = (info.domain or "").lower()

        buckets: dict[str, list[EmailSummary]] = {
            "action_required": [],
            "high_importance": [],
            "external_unread": [],
            "internal_unread": [],
            "read_external": [],
            "read_internal": [],
        }
        for e in emails:
            tag_action = "[action required]" in (e.subject or "").lower()
            high = e.importance == "high"
            d = domain_of(e.sender_smtp)
            external = bool(d) and d != owner_domain
            unread = not e.is_read
            if tag_action:
                buckets["action_required"].append(e)
            elif high:
                buckets["high_importance"].append(e)
            elif unread and external:
                buckets["external_unread"].append(e)
            elif unread:
                buckets["internal_unread"].append(e)
            elif external:
                buckets["read_external"].append(e)
            else:
                buckets["read_internal"].append(e)

        return {
            "window": {"start_utc": params.start_utc.isoformat(), "end_utc": params.end_utc.isoformat()},
            "owner_domain": owner_domain,
            "counts": {k: len(v) for k, v in buckets.items()},
            "top_per_bucket": {
                k: [_summary_to_json(s) for s in sorted(v, key=lambda x: _ts(x) or datetime.min.replace(tzinfo=timezone.utc), reverse=True)[:10]]
                for k, v in buckets.items()
            },
        }

    @mcp.tool(
        name="outlook_extract_action_items",
        description=(
            "Composite. Find emails with action-required signals (subject tags, "
            "importance flags, or known commitment phrases in the preview). Returns "
            "structured items with sender, subject, signal type, and conversation_id."
        ),
    )
    async def outlook_extract_action_items(
        start_utc: str,
        end_utc: str,
        folders: list[str] | None = None,
        max_results: int = 300,
    ) -> dict[str, Any]:
        params = ExtractActionItemsInput(
            start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc),
            folders=folders or ["Inbox"],
            max_results=max_results,
        )
        emails = await backend.get_emails_in_time_range(
            params.start_utc, params.end_utc, params.folders, params.max_results
        )
        items: list[dict[str, Any]] = []
        for e in emails:
            signals: list[str] = []
            subj_lower = (e.subject or "").lower()
            preview_lower = (e.preview or "").lower()
            if "[action required]" in subj_lower:
                signals.append("subject_tag:action_required")
            if e.importance == "high":
                signals.append("importance:high")
            for phrase in _ACTION_PHRASES:
                if phrase in preview_lower:
                    signals.append(f"phrase:{phrase}")
            if signals:
                items.append(
                    {
                        "entry_id": e.entry_id,
                        "conversation_id": e.conversation_id,
                        "subject": e.subject,
                        "sender_smtp": e.sender_smtp,
                        "received_utc": e.received_utc.isoformat() if e.received_utc else None,
                        "signals": signals,
                        "preview": e.preview,
                    }
                )
        return {"action_items": items, "count": len(items)}

    @mcp.tool(
        name="outlook_find_unanswered",
        description=(
            "Composite. Sent items where you're waiting on a reply: sent in the window "
            "with no inbound reply in the same conversation since `waiting_hours` after "
            "send. Returns one entry per stalled thread."
        ),
    )
    async def outlook_find_unanswered(
        start_utc: str,
        end_utc: str,
        waiting_hours: int = 24,
        max_results: int = 200,
    ) -> dict[str, Any]:
        params = FindUnansweredInput(
            start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc),
            waiting_hours=waiting_hours,
            max_results=max_results,
        )
        info = await backend.get_mailbox_info()
        owner_smtp = (info.smtp_address or "").lower()

        sent = await backend.get_emails_in_time_range(
            params.start_utc, params.end_utc, ["Sent Items"], params.max_results
        )
        # Latest sent message per conversation_id
        last_sent: dict[str, EmailSummary] = {}
        for s in sent:
            cid = s.conversation_id
            if not cid:
                continue
            ts = _ts(s)
            if not ts:
                continue
            cur = last_sent.get(cid)
            if cur is None or ((_ts(cur) or datetime.min.replace(tzinfo=timezone.utc)) < ts):
                last_sent[cid] = s

        threshold = timedelta(hours=params.waiting_hours)
        now = datetime.now(timezone.utc)
        stalled: list[dict[str, Any]] = []
        for cid, s in last_sent.items():
            sent_at = _utc(_ts(s))
            if now - sent_at < threshold:
                continue
            try:
                thread = await backend.get_conversation_thread(cid, None, max_messages=200)
            except Exception:
                continue
            inbound_after = [
                m for m in thread
                if (m.sender_smtp or "").lower() != owner_smtp
                and (_ts(m) or datetime.min.replace(tzinfo=timezone.utc)) > sent_at
            ]
            if inbound_after:
                continue
            stalled.append(
                {
                    "conversation_id": cid,
                    "subject": s.subject,
                    "last_sent_to": s.to_smtp,
                    "last_sent_utc": sent_at.isoformat(),
                    "hours_waiting": int((now - sent_at).total_seconds() / 3600),
                }
            )
        stalled.sort(key=lambda r: r["hours_waiting"], reverse=True)
        return {"stalled_threads": stalled, "count": len(stalled)}

    @mcp.tool(
        name="outlook_find_promised_actions",
        description=(
            "Composite. Search your Sent items for commitment phrases (\"I'll send\", "
            "\"by EOD\", etc.) so you can spot promises you may have forgotten. "
            "Returns one entry per matching sent message."
        ),
    )
    async def outlook_find_promised_actions(
        start_utc: str,
        end_utc: str,
        max_results: int = 200,
        extra_phrases: list[str] | None = None,
    ) -> dict[str, Any]:
        params = FindPromisedActionsInput(
            start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc),
            max_results=max_results,
            extra_phrases=extra_phrases or [],
        )
        phrases = list(_PROMISE_PHRASES) + [p.lower() for p in params.extra_phrases]
        sent = await backend.get_emails_in_time_range(
            params.start_utc, params.end_utc, ["Sent Items"], params.max_results
        )
        results: list[dict[str, Any]] = []
        for s in sent:
            preview_lower = (s.preview or "").lower()
            hits = [p for p in phrases if p in preview_lower]
            if hits:
                results.append(
                    {
                        "entry_id": s.entry_id,
                        "conversation_id": s.conversation_id,
                        "subject": s.subject,
                        "to_smtp": s.to_smtp,
                        "sent_utc": s.sent_utc.isoformat() if s.sent_utc else None,
                        "phrases": hits,
                        "preview": s.preview,
                    }
                )
        return {"promises": results, "count": len(results)}

    @mcp.tool(
        name="outlook_meeting_prep",
        description=(
            "Composite. For an upcoming calendar event, return the event details plus "
            "the most recent threads exchanged with each attendee in the past N days. "
            "Saves the agent multiple round trips."
        ),
    )
    async def outlook_meeting_prep(
        event_id: str,
        history_window_days: int = 30,
        max_threads_per_attendee: int = 3,
    ) -> dict[str, Any]:
        params = MeetingPrepInput(
            event_id=event_id,
            history_window_days=history_window_days,
            max_threads_per_attendee=max_threads_per_attendee,
        )
        event = await backend.get_calendar_event(params.event_id)
        end = datetime.now(timezone.utc)
        start = end - timedelta(days=params.history_window_days)
        # Pull all recent mail once; bucket by attendee.
        recent = await backend.get_emails_in_time_range(
            start, end, ["Inbox", "Sent Items"], max_results=2000
        )
        per_attendee: dict[str, dict[str, EmailSummary]] = defaultdict(dict)
        owner = (await backend.get_mailbox_info()).smtp_address.lower()
        attendees = [a.lower() for a in event.attendees_smtp if a]
        for m in recent:
            counterparts = {addr.lower() for addr in (m.sender_smtp, *m.to_smtp, *m.cc_smtp) if addr}
            counterparts.discard(owner)
            for a in attendees:
                if a in counterparts and m.conversation_id:
                    cur = per_attendee[a].get(m.conversation_id)
                    if cur is None or ((_ts(cur) or datetime.min.replace(tzinfo=timezone.utc)) < (_ts(m) or datetime.min.replace(tzinfo=timezone.utc))):
                        per_attendee[a][m.conversation_id] = m

        return {
            "event": {
                "event_id": event.event_id,
                "subject": event.subject,
                "organizer_smtp": event.organizer_smtp,
                "start_utc": event.start_utc.isoformat() if event.start_utc else None,
                "end_utc": event.end_utc.isoformat() if event.end_utc else None,
                "location": event.location,
                "attendees_smtp": event.attendees_smtp,
                "body_plain": event.body_plain,
            },
            "context_per_attendee": {
                a: [
                    _summary_to_json(s)
                    for s in sorted(
                        threads.values(),
                        key=lambda x: _ts(x) or datetime.min.replace(tzinfo=timezone.utc),
                        reverse=True,
                    )[: params.max_threads_per_attendee]
                ]
                for a, threads in per_attendee.items()
            },
        }

    @mcp.tool(
        name="outlook_relationship_graph",
        description=(
            "Composite. Per-correspondent communication frequency over a window: "
            "messages exchanged, last contact time, internal vs external. Useful for "
            "the agent to judge relationship intensity at a glance."
        ),
    )
    async def outlook_relationship_graph(
        start_utc: str,
        end_utc: str,
        folders: list[str] | None = None,
        max_results: int = 2000,
        top_n: int = 25,
    ) -> dict[str, Any]:
        params = RelationshipGraphInput(
            start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc),
            folders=folders or ["Inbox", "Sent Items"],
            max_results=max_results,
            top_n=top_n,
        )
        info = await backend.get_mailbox_info()
        owner_smtp = (info.smtp_address or "").lower()
        owner_domain = (info.domain or "").lower()
        emails = await backend.get_emails_in_time_range(
            params.start_utc, params.end_utc, params.folders, params.max_results
        )
        stats: dict[str, dict[str, Any]] = defaultdict(
            lambda: {"sent_to": 0, "received_from": 0, "last_contact_utc": None}
        )
        for e in emails:
            ts = _ts(e)
            sender = (e.sender_smtp or "").lower()
            if sender and sender == owner_smtp:
                # Outbound
                for to in (*e.to_smtp, *e.cc_smtp):
                    addr = to.lower()
                    if not addr or addr == owner_smtp:
                        continue
                    s = stats[addr]
                    s["sent_to"] += 1
                    if ts and (s["last_contact_utc"] is None or ts > s["last_contact_utc"]):
                        s["last_contact_utc"] = ts
            elif sender:
                s = stats[sender]
                s["received_from"] += 1
                if ts and (s["last_contact_utc"] is None or ts > s["last_contact_utc"]):
                    s["last_contact_utc"] = ts

        top: list[dict[str, Any]] = []
        for addr, s in stats.items():
            total = s["sent_to"] + s["received_from"]
            top.append(
                {
                    "smtp": addr,
                    "domain": domain_of(addr),
                    "internal": domain_of(addr) == owner_domain,
                    "total": total,
                    "sent_to": s["sent_to"],
                    "received_from": s["received_from"],
                    "last_contact_utc": s["last_contact_utc"].isoformat() if s["last_contact_utc"] else None,
                }
            )
        top.sort(key=lambda r: r["total"], reverse=True)
        return {
            "owner_smtp": owner_smtp,
            "owner_domain": owner_domain,
            "correspondents": top[: params.top_n],
            "total_correspondents": len(top),
        }

    @mcp.tool(
        name="outlook_threadify",
        description=(
            "Composite. Group a flat list of email entry_ids into threads using the "
            "shared ConversationID and our normalized subject. Returns one entry per "
            "thread with messages sorted oldest -> newest."
        ),
    )
    async def outlook_threadify(entry_ids: list[str]) -> dict[str, Any]:
        params = ThreadifyInput(entry_ids=entry_ids)
        # Pull metadata for each entry.
        emails: list[EmailSummary] = []
        for eid in params.entry_ids:
            try:
                meta = await backend.get_email_full_metadata(eid)
                emails.append(meta.summary)
            except Exception:
                continue
        threads: dict[str, list[EmailSummary]] = defaultdict(list)
        for e in emails:
            cid = e.conversation_id or normalize_subject(e.subject) or e.entry_id
            threads[cid].append(e)
        out = []
        for cid, msgs in threads.items():
            msgs.sort(key=lambda m: _ts(m) or datetime.min.replace(tzinfo=timezone.utc))
            out.append(
                {
                    "conversation_id": cid,
                    "normalized_subject": normalize_subject(msgs[0].subject) if msgs else "",
                    "message_count": len(msgs),
                    "messages": [_summary_to_json(m) for m in msgs],
                }
            )
        out.sort(key=lambda t: t["message_count"], reverse=True)
        return {"threads": out, "thread_count": len(out)}
