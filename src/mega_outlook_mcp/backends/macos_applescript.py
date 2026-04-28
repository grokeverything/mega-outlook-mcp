"""macOS AppleScript backend.

Parses the sentinel-delimited output produced by `applescript/templates.py`.
Fields that cannot be resolved on Mac (e.g. RFC 5322 headers when `source`
access is blocked by "New Outlook") are returned as the literal string
`MAC_UNAVAILABLE` rather than None.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from ..applescript import (
    LARGE_SCAN_TIMEOUT_SECONDS,
    run_osascript,
)
from ..applescript import templates as tpl
from ..applescript.templates import FLD, REC
from ..baseline import load_baseline
from ..constants import MAC_UNAVAILABLE
from ..errors import (
    AttachmentError,
    ConversationNotFoundError,
    EmailNotFoundError,
    FolderNotFoundError,
)
from ..utils.email_extract import domain_of, preview_text
from ..utils.folder_utils import canonicalize_folder_name
from ..utils.rfc5322 import parse_headers
from .base import (
    AccountInfo,
    AttachmentInfo,
    Backend,
    CalendarEvent,
    Contact,
    DiagnosticsResult,
    EmailBody,
    EmailFullMetadata,
    EmailSummary,
    FolderInfo,
    FreeBusyResponse,
    MailboxInfo,
    MailboxQuota,
    NoteInfo,
    OperationResult,
    OutOfOfficeStatus,
    RuleInfo,
    SignatureInfo,
    TaskInfo,
)


class MacOSAppleScriptBackend(Backend):
    # ------------------------------------------------------------------
    # Mailbox / folders
    # ------------------------------------------------------------------
    async def get_mailbox_info(self) -> MailboxInfo:
        raw = await run_osascript(tpl.mailbox_info())
        fields = _parse_single(raw)
        smtp = fields.get("smtp", "")
        return MailboxInfo(
            display_name=fields.get("name", ""),
            smtp_address=smtp,
            domain=domain_of(smtp),
        )

    async def list_folders(self, include_subfolders: bool) -> list[FolderInfo]:
        raw = await run_osascript(tpl.list_folders(include_subfolders))
        records = _parse_records(raw)
        return [
            FolderInfo(
                name=r.get("name", ""),
                full_path=r.get("path", r.get("name", "")),
                item_count=_int(r.get("count", "0")),
                unread_count=_int(r.get("unread", "0")),
                parent_path=_parent_of(r.get("path", "")),
            )
            for r in records
        ]

    # ------------------------------------------------------------------
    # Emails
    # ------------------------------------------------------------------
    async def get_emails_in_time_range(
        self,
        start_utc: datetime,
        end_utc: datetime,
        folders: list[str],
        max_results: int,
    ) -> list[EmailSummary]:
        resolved = [canonicalize_folder_name(f) for f in folders]
        raw = await run_osascript(
            tpl.emails_in_time_range(start_utc, end_utc, resolved, max_results),
            timeout=LARGE_SCAN_TIMEOUT_SECONDS,
        )
        return [_record_to_summary(r) for r in _parse_records(raw)]

    async def get_conversation_thread(
        self,
        conversation_id: str,
        conversation_topic: str | None,
        max_messages: int,
    ) -> list[EmailSummary]:
        raw = await run_osascript(
            tpl.conversation_thread(conversation_id, max_messages),
            timeout=LARGE_SCAN_TIMEOUT_SECONDS,
        )
        if raw.strip().startswith("ERR:"):
            raise ConversationNotFoundError(
                f"No messages for conversation_id {conversation_id!r}."
            )
        summaries = [_record_to_summary(r) for r in _parse_records(raw)]
        if not summaries:
            raise ConversationNotFoundError(
                f"No messages for conversation_id {conversation_id!r}."
            )
        summaries.sort(
            key=lambda s: s.sent_utc
            or s.received_utc
            or datetime.min.replace(tzinfo=timezone.utc)
        )
        return summaries

    async def get_email_full_metadata(self, entry_id: str) -> EmailFullMetadata:
        raw = await run_osascript(
            tpl.email_metadata(entry_id), timeout=LARGE_SCAN_TIMEOUT_SECONDS
        )
        records = _parse_records(raw)
        if not records:
            raise EmailNotFoundError(f"No email with id {entry_id!r}.")
        head = records[0]
        summary = _record_to_summary(head)
        html = head.get("html") or None
        source_blob = head.get("source", "")
        if source_blob == "__UNAVAILABLE__" or not source_blob:
            msg_id = MAC_UNAVAILABLE
            in_reply_to = MAC_UNAVAILABLE
            references: list[str] | str | None = MAC_UNAVAILABLE
            headers_raw: str | None = MAC_UNAVAILABLE
        else:
            try:
                parsed = parse_headers(source_blob.encode("utf-8", errors="replace"))
                msg_id = parsed.get("message_id") or MAC_UNAVAILABLE
                in_reply_to = parsed.get("in_reply_to") or MAC_UNAVAILABLE
                references_val = parsed.get("references") or []
                references = references_val if references_val else MAC_UNAVAILABLE
                headers_raw = parsed.get("headers_raw") or MAC_UNAVAILABLE
            except Exception:
                msg_id = MAC_UNAVAILABLE
                in_reply_to = MAC_UNAVAILABLE
                references = MAC_UNAVAILABLE
                headers_raw = MAC_UNAVAILABLE
        attachments: list[AttachmentInfo] = []
        for idx, att in enumerate(records[1:], start=1):
            cid = att.get("attachCid", "")
            attachments.append(
                AttachmentInfo(
                    index=idx,
                    filename=att.get("attachName", ""),
                    size_bytes=_int(att.get("attachSize", "0")),
                    is_inline=bool(cid),
                )
            )
        return EmailFullMetadata(
            summary=summary,
            body=EmailBody(
                plain_text=head.get("preview", ""),
                html=html,
            ),
            attachments=attachments,
            internet_message_id=msg_id,
            in_reply_to=in_reply_to,
            references=references,
            mapi_headers_raw=headers_raw,
            delegation_sender_smtp=summary.sender_smtp or None,
            delegation_representing_smtp=MAC_UNAVAILABLE,
        )

    async def save_attachment(
        self, entry_id: str, attachment_index: int, save_path: str
    ) -> AttachmentInfo:
        try:
            raw = await run_osascript(
                tpl.save_attachment(entry_id, attachment_index, save_path)
            )
        except Exception as exc:
            raise AttachmentError(
                f"Failed to save attachment {attachment_index} of {entry_id!r} to {save_path!r}: {exc}"
            ) from exc
        fields = _parse_single(raw)
        return AttachmentInfo(
            index=attachment_index,
            filename=fields.get("name", ""),
            size_bytes=_int(fields.get("size", "0")),
            is_inline=False,
        )

    async def search_emails(
        self, query: str, folders: list[str], field: str, max_results: int
    ) -> list[EmailSummary]:
        resolved = [canonicalize_folder_name(f) for f in folders]
        raw = await run_osascript(
            tpl.search_emails(query, resolved, max_results),
            timeout=LARGE_SCAN_TIMEOUT_SECONDS,
        )
        return [_record_to_summary(r) for r in _parse_records(raw)]

    # ------------------------------------------------------------------
    # Calendar
    # ------------------------------------------------------------------
    async def list_calendar_events(
        self,
        start_utc: datetime,
        end_utc: datetime,
        calendar_id: str | None,
        max_results: int,
    ) -> list[CalendarEvent]:
        raw = await run_osascript(
            tpl.list_calendar_events(start_utc, end_utc, max_results),
            timeout=LARGE_SCAN_TIMEOUT_SECONDS,
        )
        return [_record_to_calendar_event(r) for r in _parse_records(raw)]

    async def get_calendar_event(self, event_id: str) -> CalendarEvent:
        raw = await run_osascript(tpl.get_calendar_event(event_id))
        records = _parse_records(raw)
        head = _parse_single(raw)
        attendees: list[str] = []
        for rec in records:
            if "attendee" in rec:
                attendees.append(rec["attendee"])
        return CalendarEvent(
            event_id=event_id,
            subject=head.get("subject", ""),
            organizer_smtp=head.get("organizer", ""),
            start_utc=_epoch_to_utc(head.get("start")),
            end_utc=_epoch_to_utc(head.get("end")),
            location=head.get("location", ""),
            is_all_day=_bool(head.get("allDay", "false")),
            attendees_smtp=attendees,
            body_plain=head.get("body", ""),
        )

    # ------------------------------------------------------------------
    # Contacts
    # ------------------------------------------------------------------
    async def list_contacts(self, limit: int, offset: int) -> list[Contact]:
        raw = await run_osascript(tpl.list_contacts(limit, offset))
        return [_record_to_contact(r) for r in _parse_records(raw)]

    async def search_contacts(self, query: str, limit: int) -> list[Contact]:
        raw = await run_osascript(tpl.search_contacts(query, limit))
        return [_record_to_contact(r) for r in _parse_records(raw)]

    async def get_contact(self, contact_id: str) -> Contact:
        raw = await run_osascript(tpl.get_contact(contact_id))
        head = _parse_single(raw)
        records = _parse_records(raw)
        emails = [r["email"] for r in records if "email" in r]
        primary = emails[0] if emails else ""
        phones: dict[str, str] = {}
        for r in records:
            if "phone" in r:
                label = r.get("phoneLabel", "other") or "other"
                phones[label] = r["phone"]
        return Contact(
            contact_id=contact_id,
            display_name=head.get("name", ""),
            primary_smtp=primary,
            other_smtps=emails[1:],
            company=head.get("company", ""),
            job_title=head.get("title", ""),
            phone_numbers=phones,
        )

    async def diagnostics(self, sample_message_id: str | None = None) -> DiagnosticsResult:
        baseline = load_baseline()
        message_props = baseline.get("macos_applescript", {}).get("message_properties", [])
        try:
            raw = await run_osascript(tpl.diagnostics(message_props, sample_message_id))
        except Exception as exc:
            return DiagnosticsResult(
                platform="macos",
                outlook_version="unknown",
                baseline_version=baseline.get("baseline_outlook_versions", {}).get("macos", "unknown"),
                status="BROKEN",
                probed_fields={},
                affected_tools=["all"],
                notes=[
                    f"osascript failed: {exc}. Common cause: 'New Outlook' has no AppleScript support. "
                    "Switch to classic Outlook for Mac."
                ],
            )
        head = _parse_single(raw)
        version = head.pop("version", "unknown")
        sample = head.pop("sample", "missing")
        notes: list[str] = []
        if sample == "missing":
            notes.append("No sample message available; per-message field probes skipped.")
        probed: dict[str, str] = {}
        affected: set[str] = set()
        for key, value in head.items():
            field_name = key.replace("_", " ")
            probed[f"AppleScript.{field_name}"] = value
            if value != "ok":
                for t in _MAC_FIELD_TO_TOOLS.get(field_name, []):
                    affected.add(t)
        missing = [k for k, v in probed.items() if v != "ok"]
        if not missing:
            status = "HEALTHY"
        elif len(missing) >= len(probed) // 2:
            status = "BROKEN"
        else:
            status = f"DEGRADED:{len(missing)}_fields"
        baseline_version = baseline.get("baseline_outlook_versions", {}).get("macos", "unknown")
        return DiagnosticsResult(
            platform="macos",
            outlook_version=version,
            baseline_version=baseline_version,
            status=status,
            probed_fields=probed,
            affected_tools=sorted(affected),
            notes=notes,
        )

    # ------------------------------------------------------------------
    # Phase 3 parity: mail write
    # ------------------------------------------------------------------
    async def send_email(self, to, subject, body, body_type, cc, bcc, attachments) -> OperationResult:
        return await self._run_op(tpl.send_email(to, subject, body, body_type, cc, bcc, attachments, send=True), "sent")

    async def create_draft(self, to, subject, body, body_type, cc, bcc, attachments) -> OperationResult:
        return await self._run_op(tpl.send_email(to, subject, body, body_type, cc, bcc, attachments, send=False), "draft saved")

    async def reply_email(self, entry_id, body, body_type, reply_all) -> OperationResult:
        return await self._run_op(tpl.reply_email(entry_id, body, body_type, reply_all), "replied")

    async def forward_email(self, entry_id, to, body, body_type) -> OperationResult:
        return await self._run_op(tpl.forward_email(entry_id, to, body, body_type), "forwarded")

    # ------------------------------------------------------------------
    # Phase 3 parity: mail organize
    # ------------------------------------------------------------------
    async def mark_email_read(self, entry_id, is_read) -> OperationResult:
        return await self._run_op(tpl.mark_email_read(entry_id, is_read), f"marked {'read' if is_read else 'unread'}")

    async def set_email_flag(self, entry_id, flag_status, due_date_utc) -> OperationResult:
        # AppleScript's outgoing/incoming message flag model is limited.
        # Categories serves as the closest equivalent for reliable flagging.
        return OperationResult(
            ok=False,
            message=f"set_email_flag: {MAC_UNAVAILABLE} (use set_email_categories instead)",
        )

    async def set_email_categories(self, entry_id, categories) -> OperationResult:
        return await self._run_op(tpl.set_email_categories(entry_id, categories), "categories set")

    # ------------------------------------------------------------------
    # Phase 3 parity: mail destructive
    # ------------------------------------------------------------------
    async def move_email(self, entry_id, destination_folder) -> OperationResult:
        return await self._run_op(tpl.move_email(entry_id, destination_folder), f"moved to {destination_folder}")

    async def archive_email(self, entry_id) -> OperationResult:
        return await self._run_op(tpl.move_email(entry_id, "Archive"), "archived")

    async def delete_email(self, entry_id, permanent) -> OperationResult:
        return await self._run_op(tpl.delete_email(entry_id, permanent), "deleted" if permanent else "moved to Deleted Items")

    async def junk_email(self, entry_id) -> OperationResult:
        return await self._run_op(tpl.junk_email(entry_id), "moved to Junk")

    # ------------------------------------------------------------------
    # Phase 3 parity: folder management
    # ------------------------------------------------------------------
    async def create_folder(self, parent_path, name) -> OperationResult:
        return await self._run_op(tpl.create_folder(parent_path, name), "created")

    async def rename_folder(self, folder_path, new_name) -> OperationResult:
        return await self._run_op(tpl.rename_folder(folder_path, new_name), "renamed")

    async def move_folder(self, folder_path, new_parent_path) -> OperationResult:
        # AppleScript Outlook does not expose a direct "move folder" verb.
        return OperationResult(
            ok=False,
            message=f"move_folder: {MAC_UNAVAILABLE}",
        )

    async def delete_folder(self, folder_path) -> OperationResult:
        return await self._run_op(tpl.delete_folder(folder_path), "deleted folder")

    async def empty_folder(self, folder_path) -> OperationResult:
        return await self._run_op(tpl.empty_folder(folder_path), "emptied")

    # ------------------------------------------------------------------
    # Phase 3 parity: calendar write
    # ------------------------------------------------------------------
    async def create_calendar_event(self, subject, start_utc, end_utc, attendees, location, body, is_all_day) -> OperationResult:
        return await self._run_op(
            tpl.create_calendar_event(subject, start_utc, end_utc, attendees, location, body, is_all_day),
            "event created",
        )

    async def update_calendar_event(self, event_id, subject, start_utc, end_utc, location, body) -> OperationResult:
        return await self._run_op(
            tpl.update_calendar_event(event_id, subject, start_utc, end_utc, location, body),
            "updated",
        )

    async def delete_calendar_event(self, event_id) -> OperationResult:
        return await self._run_op(tpl.delete_calendar_event(event_id), "deleted")

    async def respond_to_event(self, event_id, response, send_response) -> OperationResult:
        # AppleScript Outlook's respond verb is unreliable across versions.
        return OperationResult(
            ok=False,
            message=f"respond_to_event: {MAC_UNAVAILABLE}",
        )

    # ------------------------------------------------------------------
    # Phase 3 parity: tasks
    # ------------------------------------------------------------------
    async def list_tasks(self, limit, offset) -> list[TaskInfo]:
        raw = await run_osascript(tpl.list_tasks(limit, offset))
        return [_record_to_task(r) for r in _parse_records(raw)]

    async def search_tasks(self, query, limit) -> list[TaskInfo]:
        raw = await run_osascript(tpl.search_tasks(query, limit))
        return [_record_to_task(r) for r in _parse_records(raw)]

    async def get_task(self, task_id) -> TaskInfo:
        raw = await run_osascript(tpl.get_task(task_id))
        head = _parse_single(raw)
        return TaskInfo(
            task_id=task_id,
            subject=head.get("subject", ""),
            due_date_utc=_epoch_to_utc(head.get("due")),
            is_complete=_bool(head.get("done", "false")),
            importance="normal",
            body_plain=head.get("body", ""),
        )

    # ------------------------------------------------------------------
    # Phase 3 parity: notes
    # ------------------------------------------------------------------
    async def list_notes(self, limit, offset) -> list[NoteInfo]:
        raw = await run_osascript(tpl.list_notes(limit, offset))
        return [_record_to_note(r) for r in _parse_records(raw)]

    async def search_notes(self, query, limit) -> list[NoteInfo]:
        raw = await run_osascript(tpl.search_notes(query, limit))
        return [_record_to_note(r) for r in _parse_records(raw)]

    async def get_note(self, note_id) -> NoteInfo:
        raw = await run_osascript(tpl.get_note(note_id))
        head = _parse_single(raw)
        return NoteInfo(
            note_id=note_id,
            subject=head.get("subject", ""),
            body_plain=head.get("body", ""),
            last_modified_utc=_epoch_to_utc(head.get("modified")),
        )

    # ------------------------------------------------------------------
    # Phase 3 parity: accounts
    # ------------------------------------------------------------------
    async def list_accounts(self) -> list[AccountInfo]:
        raw = await run_osascript(tpl.list_accounts())
        out: list[AccountInfo] = []
        for rec in _parse_records(raw):
            out.append(
                AccountInfo(
                    account_id=rec.get("id", ""),
                    display_name=rec.get("name", ""),
                    smtp_address=rec.get("smtp", ""),
                    account_type=rec.get("type", "other"),
                )
            )
        return out

    async def get_unread_count(self, folder) -> int:
        raw = await run_osascript(tpl.get_unread_count(folder))
        head = _parse_single(raw)
        return _int(head.get("count", "0"))

    # ------------------------------------------------------------------
    # Phase 4: Exchange-specific. AppleScript exposes very little of the
    # Exchange surface; most return MAC_UNAVAILABLE.
    # ------------------------------------------------------------------
    async def get_out_of_office(self) -> OutOfOfficeStatus:
        return OutOfOfficeStatus(
            enabled=False,
            internal_message=MAC_UNAVAILABLE,
            external_message=MAC_UNAVAILABLE,
        )

    async def set_out_of_office(self, status) -> OperationResult:
        return OperationResult(ok=False, message=f"set_out_of_office: {MAC_UNAVAILABLE}")

    async def get_signature(self, account_id) -> SignatureInfo:
        # Mac stores signatures inside the Outlook profile DB which we don't parse.
        return SignatureInfo(name=MAC_UNAVAILABLE, body_html=MAC_UNAVAILABLE, body_plain=MAC_UNAVAILABLE)

    async def set_signature(self, account_id, body_html, body_plain) -> OperationResult:
        return OperationResult(ok=False, message=f"set_signature: {MAC_UNAVAILABLE}")

    async def list_rules(self) -> list[RuleInfo]:
        return [RuleInfo(rule_id=MAC_UNAVAILABLE, name=MAC_UNAVAILABLE, enabled=False, description=MAC_UNAVAILABLE)]

    async def toggle_rule(self, rule_id, enabled) -> OperationResult:
        return OperationResult(ok=False, message=f"toggle_rule: {MAC_UNAVAILABLE}")

    async def calendar_freebusy(self, smtps, start_utc, end_utc, slot_minutes) -> list[FreeBusyResponse]:
        return [FreeBusyResponse(smtp=s, slots=[]) for s in smtps]

    async def meeting_room_finder(self, start_utc, end_utc, capacity, location_hint) -> list:
        return []

    async def gal_search(self, query, limit) -> list:
        # Fall back to local contacts search.
        return await self.search_contacts(query, limit)

    async def list_delegated_mailboxes(self) -> list[MailboxInfo]:
        return []

    async def list_public_folders(self) -> list[FolderInfo]:
        return []

    async def get_mailbox_quota(self) -> MailboxQuota:
        return MailboxQuota(total_bytes=0, used_bytes=0, warning_bytes=None, prohibit_send_bytes=None, item_count=0)

    # ------------------------------------------------------------------
    # Internal: shared run-and-format helper for write/destructive ops.
    # ------------------------------------------------------------------
    async def _run_op(self, script: str, success_message: str) -> OperationResult:
        try:
            raw = await run_osascript(script, timeout=LARGE_SCAN_TIMEOUT_SECONDS)
        except Exception as exc:
            return OperationResult(ok=False, message=f"{success_message} failed: {exc}")
        head = _parse_single(raw)
        return OperationResult(
            ok=True,
            message=success_message,
            affected_id=head.get("id"),
        )

    async def close(self) -> None:
        return None


def _record_to_task(rec: dict[str, str]) -> TaskInfo:
    return TaskInfo(
        task_id=rec.get("id", ""),
        subject=rec.get("subject", ""),
        due_date_utc=_epoch_to_utc(rec.get("due")),
        is_complete=_bool(rec.get("done", "false")),
        importance="normal",
        body_plain=rec.get("body", ""),
    )


def _record_to_note(rec: dict[str, str]) -> NoteInfo:
    return NoteInfo(
        note_id=rec.get("id", ""),
        subject=rec.get("subject", ""),
        body_plain=rec.get("body", ""),
        last_modified_utc=_epoch_to_utc(rec.get("modified")),
    )


# Mac field → tools that depend on it.
_MAC_FIELD_TO_TOOLS: dict[str, list[str]] = {
    "conversation": [
        "outlook_get_conversation_thread",
        "outlook_get_thread_metadata",
        "outlook_get_emails_in_time_range",
    ],
    "source": ["outlook_get_email_full_metadata"],
    "time received": ["outlook_get_emails_in_time_range", "outlook_search_emails"],
    "time sent": ["outlook_get_emails_in_time_range"],
    "plain text content": ["outlook_get_emails_in_time_range", "outlook_search_emails"],
    "content": ["outlook_get_email_full_metadata"],
    "has attachment": ["outlook_save_attachment"],
    "mail folder": ["outlook_list_folders", "outlook_get_emails_in_time_range"],
}


# ---------------------------------------------------------------------------
# Parser helpers
# ---------------------------------------------------------------------------


def _parse_records(raw: str) -> list[dict[str, str]]:
    """Split raw osascript stdout into a list of field dicts per record."""
    if not raw:
        return []
    parts = raw.split(REC)
    records: list[dict[str, str]] = []
    for part in parts:
        if not part.strip():
            continue
        fields = _parse_fields(part)
        if fields:
            records.append(fields)
    return records


def _parse_single(raw: str) -> dict[str, str]:
    """Parse a single (non-recordized) field block."""
    return _parse_fields(raw)


def _parse_fields(block: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    for segment in block.split(FLD):
        if "=" not in segment:
            continue
        key, _, value = segment.partition("=")
        key = key.strip()
        if not key:
            continue
        # Trim trailing whitespace introduced by AppleScript.
        fields[key] = value.rstrip("\n\r")
    return fields


def _int(value: Any) -> int:
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return 0


def _bool(value: Any) -> bool:
    return str(value).strip().lower() == "true"


def _epoch_to_utc(value: Any) -> datetime | None:
    if value is None or value == "":
        return None
    try:
        secs = int(value)
    except (TypeError, ValueError):
        return None
    try:
        return datetime.fromtimestamp(secs, tz=timezone.utc)
    except (OverflowError, OSError, ValueError):
        return None


def _record_to_summary(rec: dict[str, str]) -> EmailSummary:
    return EmailSummary(
        entry_id=rec.get("id", ""),
        conversation_id=rec.get("convId", ""),
        subject=rec.get("subject", ""),
        sender_name=rec.get("senderName", ""),
        sender_smtp=rec.get("senderAddr", ""),
        to_smtp=[],
        cc_smtp=[],
        received_utc=_epoch_to_utc(rec.get("received")),
        sent_utc=_epoch_to_utc(rec.get("sent")),
        is_read=_bool(rec.get("isRead", "false")),
        importance="normal",
        has_real_attachments=_bool(rec.get("hasAttach", "false")),
        folder_path=rec.get("folder", ""),
        preview=preview_text(rec.get("preview", "")),
    )


def _record_to_calendar_event(rec: dict[str, str]) -> CalendarEvent:
    return CalendarEvent(
        event_id=rec.get("id", ""),
        subject=rec.get("subject", ""),
        organizer_smtp=rec.get("organizer", ""),
        start_utc=_epoch_to_utc(rec.get("start")),
        end_utc=_epoch_to_utc(rec.get("end")),
        location=rec.get("location", ""),
        is_all_day=_bool(rec.get("allDay", "false")),
    )


def _record_to_contact(rec: dict[str, str]) -> Contact:
    return Contact(
        contact_id=rec.get("id", ""),
        display_name=rec.get("name", ""),
        primary_smtp=rec.get("email", ""),
        company=rec.get("company", ""),
        job_title=rec.get("title", ""),
    )


def _parent_of(path: str) -> str | None:
    if "/" not in path:
        return None
    return path.rsplit("/", 1)[0]


__all__ = ["MacOSAppleScriptBackend"]
