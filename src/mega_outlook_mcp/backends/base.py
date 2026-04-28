"""Backend Protocol and shared result dataclasses.

Both the Windows COM backend and the macOS AppleScript backend implement
this Protocol. Tool handlers depend only on this interface.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Protocol, runtime_checkable

# ---------------------------------------------------------------------------
# Shared result dataclasses
# ---------------------------------------------------------------------------


@dataclass
class MailboxInfo:
    display_name: str
    smtp_address: str
    domain: str


@dataclass
class FolderInfo:
    name: str
    full_path: str
    item_count: int
    unread_count: int
    parent_path: str | None = None


@dataclass
class AttachmentInfo:
    index: int
    filename: str
    size_bytes: int
    is_inline: bool


@dataclass
class EmailSummary:
    entry_id: str
    conversation_id: str
    subject: str
    sender_name: str
    sender_smtp: str
    to_smtp: list[str]
    cc_smtp: list[str]
    received_utc: datetime | None
    sent_utc: datetime | None
    is_read: bool
    importance: str
    has_real_attachments: bool
    folder_path: str
    preview: str


@dataclass
class EmailBody:
    plain_text: str
    html: str | None


@dataclass
class EmailFullMetadata:
    summary: EmailSummary
    body: EmailBody
    attachments: list[AttachmentInfo]
    # Fields below may be the literal MAC_UNAVAILABLE string on Mac when the
    # backend cannot resolve them (e.g. sandboxed "New Outlook").
    internet_message_id: str | None
    in_reply_to: str | None
    references: list[str] | str | None
    mapi_headers_raw: str | None
    delegation_sender_smtp: str | None
    delegation_representing_smtp: str | None


@dataclass
class ThreadMetadata:
    conversation_id: str
    normalized_subject: str
    message_count: int
    unread_count: int
    participants_smtp: list[str]
    participant_domains: list[str]
    first_message_utc: datetime | None
    last_message_utc: datetime | None
    escalated_importance: str
    internal_only: bool


@dataclass
class CalendarEvent:
    event_id: str
    subject: str
    organizer_smtp: str
    start_utc: datetime | None
    end_utc: datetime | None
    location: str
    is_all_day: bool
    attendees_smtp: list[str] = field(default_factory=list)
    body_plain: str = ""
    is_recurring: bool = False


@dataclass
class Contact:
    contact_id: str
    display_name: str
    primary_smtp: str
    other_smtps: list[str] = field(default_factory=list)
    company: str = ""
    job_title: str = ""
    phone_numbers: dict[str, str] = field(default_factory=dict)


@dataclass
class DiagnosticsResult:
    """Output of a backend self-test against the baseline manifest."""

    platform: str
    outlook_version: str
    baseline_version: str
    status: str  # HEALTHY | DEGRADED | BROKEN
    probed_fields: dict[str, str]  # field_name -> "ok" | "missing" | "error:..."
    affected_tools: list[str]
    notes: list[str]


# ---------------------------------------------------------------------------
# Phase 3 (parity) result dataclasses
# ---------------------------------------------------------------------------


@dataclass
class TaskInfo:
    task_id: str
    subject: str
    due_date_utc: datetime | None
    is_complete: bool
    importance: str
    body_plain: str = ""


@dataclass
class NoteInfo:
    note_id: str
    subject: str
    body_plain: str
    last_modified_utc: datetime | None


@dataclass
class AccountInfo:
    account_id: str
    display_name: str
    smtp_address: str
    account_type: str  # "exchange" | "imap" | "pop" | "other"


@dataclass
class OperationResult:
    """Generic ack for write/destructive ops."""

    ok: bool
    message: str = ""
    affected_id: str | None = None


# ---------------------------------------------------------------------------
# Phase 4 (Exchange-specific) result dataclasses
# ---------------------------------------------------------------------------


@dataclass
class OutOfOfficeStatus:
    enabled: bool
    internal_message: str
    external_message: str
    start_utc: datetime | None = None
    end_utc: datetime | None = None
    external_audience: str = "all"  # "all" | "contacts_only" | "none"


@dataclass
class SignatureInfo:
    name: str
    body_html: str
    body_plain: str


@dataclass
class RuleInfo:
    rule_id: str
    name: str
    enabled: bool
    description: str = ""


@dataclass
class FreeBusySlot:
    start_utc: datetime
    end_utc: datetime
    status: str  # "free" | "tentative" | "busy" | "oof" | "unknown"


@dataclass
class FreeBusyResponse:
    smtp: str
    slots: list[FreeBusySlot] = field(default_factory=list)


@dataclass
class MailboxQuota:
    total_bytes: int
    used_bytes: int
    warning_bytes: int | None
    prohibit_send_bytes: int | None
    item_count: int


# ---------------------------------------------------------------------------
# Backend Protocol
# ---------------------------------------------------------------------------


@runtime_checkable
class Backend(Protocol):
    """Operations every backend must provide.

    All methods are `async` so the FastMCP event loop stays responsive.
    Synchronous COM or `osascript` work happens off-loop inside the
    implementation (STA thread on Windows, `asyncio.to_thread` on Mac).
    """

    async def get_mailbox_info(self) -> MailboxInfo: ...

    async def list_folders(self, include_subfolders: bool) -> list[FolderInfo]: ...

    async def get_emails_in_time_range(
        self,
        start_utc: datetime,
        end_utc: datetime,
        folders: list[str],
        max_results: int,
    ) -> list[EmailSummary]: ...

    async def get_conversation_thread(
        self,
        conversation_id: str,
        conversation_topic: str | None,
        max_messages: int,
    ) -> list[EmailSummary]: ...

    async def get_email_full_metadata(self, entry_id: str) -> EmailFullMetadata: ...

    async def save_attachment(
        self, entry_id: str, attachment_index: int, save_path: str
    ) -> AttachmentInfo: ...

    async def search_emails(
        self,
        query: str,
        folders: list[str],
        field: str,
        max_results: int,
    ) -> list[EmailSummary]: ...

    async def list_calendar_events(
        self,
        start_utc: datetime,
        end_utc: datetime,
        calendar_id: str | None,
        max_results: int,
    ) -> list[CalendarEvent]: ...

    async def get_calendar_event(self, event_id: str) -> CalendarEvent: ...

    async def list_contacts(self, limit: int, offset: int) -> list[Contact]: ...

    async def search_contacts(self, query: str, limit: int) -> list[Contact]: ...

    async def get_contact(self, contact_id: str) -> Contact: ...

    async def diagnostics(self, sample_message_id: str | None = None) -> DiagnosticsResult: ...

    # ------------------------------------------------------------------
    # Phase 3 parity: mail write
    # ------------------------------------------------------------------
    async def send_email(
        self,
        to: list[str],
        subject: str,
        body: str,
        body_type: str,
        cc: list[str] | None,
        bcc: list[str] | None,
        attachments: list[str] | None,
    ) -> OperationResult: ...

    async def create_draft(
        self,
        to: list[str],
        subject: str,
        body: str,
        body_type: str,
        cc: list[str] | None,
        bcc: list[str] | None,
        attachments: list[str] | None,
    ) -> OperationResult: ...

    async def reply_email(
        self, entry_id: str, body: str, body_type: str, reply_all: bool
    ) -> OperationResult: ...

    async def forward_email(
        self, entry_id: str, to: list[str], body: str, body_type: str
    ) -> OperationResult: ...

    # ------------------------------------------------------------------
    # Phase 3 parity: mail organize
    # ------------------------------------------------------------------
    async def mark_email_read(self, entry_id: str, is_read: bool) -> OperationResult: ...

    async def set_email_flag(
        self, entry_id: str, flag_status: str, due_date_utc: datetime | None
    ) -> OperationResult: ...

    async def set_email_categories(
        self, entry_id: str, categories: list[str]
    ) -> OperationResult: ...

    # ------------------------------------------------------------------
    # Phase 3 parity: mail destructive
    # ------------------------------------------------------------------
    async def move_email(
        self, entry_id: str, destination_folder: str
    ) -> OperationResult: ...

    async def archive_email(self, entry_id: str) -> OperationResult: ...

    async def delete_email(self, entry_id: str, permanent: bool) -> OperationResult: ...

    async def junk_email(self, entry_id: str) -> OperationResult: ...

    # ------------------------------------------------------------------
    # Phase 3 parity: folder management
    # ------------------------------------------------------------------
    async def create_folder(
        self, parent_path: str | None, name: str
    ) -> OperationResult: ...

    async def rename_folder(
        self, folder_path: str, new_name: str
    ) -> OperationResult: ...

    async def move_folder(
        self, folder_path: str, new_parent_path: str
    ) -> OperationResult: ...

    async def delete_folder(self, folder_path: str) -> OperationResult: ...

    async def empty_folder(self, folder_path: str) -> OperationResult: ...

    # ------------------------------------------------------------------
    # Phase 3 parity: calendar write
    # ------------------------------------------------------------------
    async def create_calendar_event(
        self,
        subject: str,
        start_utc: datetime,
        end_utc: datetime,
        attendees: list[str],
        location: str,
        body: str,
        is_all_day: bool,
    ) -> OperationResult: ...

    async def update_calendar_event(
        self,
        event_id: str,
        subject: str | None,
        start_utc: datetime | None,
        end_utc: datetime | None,
        location: str | None,
        body: str | None,
    ) -> OperationResult: ...

    async def delete_calendar_event(self, event_id: str) -> OperationResult: ...

    async def respond_to_event(
        self, event_id: str, response: str, send_response: bool
    ) -> OperationResult: ...

    # ------------------------------------------------------------------
    # Phase 3 parity: tasks
    # ------------------------------------------------------------------
    async def list_tasks(self, limit: int, offset: int) -> list[TaskInfo]: ...

    async def search_tasks(self, query: str, limit: int) -> list[TaskInfo]: ...

    async def get_task(self, task_id: str) -> TaskInfo: ...

    # ------------------------------------------------------------------
    # Phase 3 parity: notes
    # ------------------------------------------------------------------
    async def list_notes(self, limit: int, offset: int) -> list[NoteInfo]: ...

    async def search_notes(self, query: str, limit: int) -> list[NoteInfo]: ...

    async def get_note(self, note_id: str) -> NoteInfo: ...

    # ------------------------------------------------------------------
    # Phase 3 parity: accounts
    # ------------------------------------------------------------------
    async def list_accounts(self) -> list[AccountInfo]: ...

    async def get_unread_count(self, folder: str | None) -> int: ...

    # ------------------------------------------------------------------
    # Phase 4 (Exchange-specific)
    # ------------------------------------------------------------------
    async def get_out_of_office(self) -> OutOfOfficeStatus: ...

    async def set_out_of_office(self, status: OutOfOfficeStatus) -> OperationResult: ...

    async def get_signature(self, account_id: str | None) -> SignatureInfo: ...

    async def set_signature(
        self, account_id: str | None, body_html: str, body_plain: str
    ) -> OperationResult: ...

    async def list_rules(self) -> list[RuleInfo]: ...

    async def toggle_rule(self, rule_id: str, enabled: bool) -> OperationResult: ...

    async def calendar_freebusy(
        self,
        smtps: list[str],
        start_utc: datetime,
        end_utc: datetime,
        slot_minutes: int,
    ) -> list[FreeBusyResponse]: ...

    async def meeting_room_finder(
        self, start_utc: datetime, end_utc: datetime, capacity: int, location_hint: str | None
    ) -> list[Contact]: ...

    async def gal_search(self, query: str, limit: int) -> list[Contact]: ...

    async def list_delegated_mailboxes(self) -> list[MailboxInfo]: ...

    async def list_public_folders(self) -> list[FolderInfo]: ...

    async def get_mailbox_quota(self) -> MailboxQuota: ...

    async def close(self) -> None:
        """Release resources (shut down the STA thread, etc.)."""
        ...
