"""Pydantic v2 input models for every MCP tool.

One class per tool. These models are shared across both backends — platform
dispatch happens inside the tool handler after input validation.
"""

from __future__ import annotations

from datetime import datetime
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


class _Base(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")


# ---------------------------------------------------------------------------
# Email / extraction tools
# ---------------------------------------------------------------------------


class GetCurrentTimeInput(_Base):
    lookback_minutes: int = Field(
        default=65,
        ge=1,
        le=1440,
        description="Size of the default lookback window in minutes.",
    )


class ListFoldersInput(_Base):
    include_subfolders: bool = Field(
        default=True,
        description="Recurse into subfolders.",
    )


class GetEmailsInTimeRangeInput(_Base):
    start_utc: datetime = Field(description="Inclusive window start (UTC, ISO 8601).")
    end_utc: datetime = Field(description="Exclusive window end (UTC, ISO 8601).")
    folders: list[str] = Field(
        default_factory=lambda: ["Inbox", "Sent Items"],
        description="Folder names to scan. Use outlook_list_folders to discover.",
    )
    max_results: int = Field(default=500, ge=1, le=5000)


class GetConversationThreadInput(_Base):
    conversation_id: str = Field(min_length=1)
    conversation_topic: str | None = Field(
        default=None,
        description="Optional; when supplied enables the Restrict fast path.",
    )
    max_messages: int = Field(default=200, ge=1, le=1000)


class GetEmailFullMetadataInput(_Base):
    entry_id: str = Field(min_length=1, description="EntryID (Windows) or message id (Mac).")


class SaveAttachmentInput(_Base):
    entry_id: str = Field(min_length=1)
    attachment_index: int = Field(ge=1, description="1-based index within the email.")
    save_path: str = Field(min_length=1, description="Absolute destination path.")


class GetMailboxInfoInput(_Base):
    pass


class SearchEmailsInput(_Base):
    query: str = Field(min_length=1)
    folders: list[str] = Field(default_factory=lambda: ["Inbox"])
    field: Literal["subject", "sender", "body", "any"] = "any"
    max_results: int = Field(default=100, ge=1, le=1000)


class WriteFileInput(_Base):
    path: str = Field(min_length=1, description="Absolute file path to write.")
    content: str = Field(description="Content to write; UTF-8.")
    overwrite: bool = Field(default=False)


class GetThreadMetadataInput(_Base):
    conversation_id: str = Field(min_length=1)
    conversation_topic: str | None = None
    mailbox_domain: str | None = Field(
        default=None,
        description="Override mailbox domain for internal/external classification.",
    )


# ---------------------------------------------------------------------------
# Calendar (read-only) tools
# ---------------------------------------------------------------------------


class ListCalendarEventsInput(_Base):
    start_utc: datetime
    end_utc: datetime
    calendar_id: str | None = Field(
        default=None,
        description="Specific calendar id; defaults to the primary calendar.",
    )
    max_results: int = Field(default=200, ge=1, le=1000)


class GetCalendarEventInput(_Base):
    event_id: str = Field(min_length=1)


# ---------------------------------------------------------------------------
# Contact (read-only) tools
# ---------------------------------------------------------------------------


class ListContactsInput(_Base):
    limit: int = Field(default=100, ge=1, le=500)
    offset: int = Field(default=0, ge=0)


class SearchContactsInput(_Base):
    query: str = Field(min_length=1)
    limit: int = Field(default=50, ge=1, le=500)


class GetContactInput(_Base):
    contact_id: str = Field(min_length=1)


# ---------------------------------------------------------------------------
# Composite (agent-friendly) tools
# ---------------------------------------------------------------------------


class SummarizeInboxInput(_Base):
    start_utc: datetime
    end_utc: datetime
    folders: list[str] = Field(default_factory=lambda: ["Inbox"])
    max_results: int = Field(default=300, ge=1, le=2000)


class ExtractActionItemsInput(_Base):
    start_utc: datetime
    end_utc: datetime
    folders: list[str] = Field(default_factory=lambda: ["Inbox"])
    max_results: int = Field(default=300, ge=1, le=2000)


class FindUnansweredInput(_Base):
    start_utc: datetime = Field(description="Earliest sent message to consider.")
    end_utc: datetime
    waiting_hours: int = Field(
        default=24, ge=1, le=720, description="Min hours since send with no inbound reply."
    )
    max_results: int = Field(default=200, ge=1, le=1000)


class FindPromisedActionsInput(_Base):
    start_utc: datetime
    end_utc: datetime
    max_results: int = Field(default=200, ge=1, le=1000)
    extra_phrases: list[str] = Field(
        default_factory=list,
        description="Additional commitment phrases to search beyond the built-in list.",
    )


class MeetingPrepInput(_Base):
    event_id: str = Field(min_length=1)
    history_window_days: int = Field(default=30, ge=1, le=365)
    max_threads_per_attendee: int = Field(default=3, ge=1, le=20)


class RelationshipGraphInput(_Base):
    start_utc: datetime
    end_utc: datetime
    folders: list[str] = Field(default_factory=lambda: ["Inbox", "Sent Items"])
    max_results: int = Field(default=2000, ge=1, le=10000)
    top_n: int = Field(default=25, ge=1, le=500)


class ThreadifyInput(_Base):
    entry_ids: list[str] = Field(min_length=1, max_length=2000)


# ---------------------------------------------------------------------------
# Phase 3 (parity) input models
# ---------------------------------------------------------------------------


class SendEmailInput(_Base):
    to: list[str] = Field(min_length=1)
    subject: str
    body: str
    body_type: Literal["plain", "html"] = "plain"
    cc: list[str] | None = None
    bcc: list[str] | None = None
    attachments: list[str] | None = Field(
        default=None, description="Absolute paths to attach.")


class CreateDraftInput(SendEmailInput):
    pass


class ReplyEmailInput(_Base):
    entry_id: str = Field(min_length=1)
    body: str
    body_type: Literal["plain", "html"] = "plain"
    reply_all: bool = False


class ForwardEmailInput(_Base):
    entry_id: str = Field(min_length=1)
    to: list[str] = Field(min_length=1)
    body: str
    body_type: Literal["plain", "html"] = "plain"


class MarkEmailReadInput(_Base):
    entry_id: str = Field(min_length=1)
    is_read: bool = True


class SetEmailFlagInput(_Base):
    entry_id: str = Field(min_length=1)
    flag_status: Literal["none", "marked", "complete"] = "marked"
    due_date_utc: datetime | None = None


class SetEmailCategoriesInput(_Base):
    entry_id: str = Field(min_length=1)
    categories: list[str] = Field(default_factory=list)


class MoveEmailInput(_Base):
    entry_id: str = Field(min_length=1)
    destination_folder: str = Field(min_length=1)


class ArchiveEmailInput(_Base):
    entry_id: str = Field(min_length=1)


class DeleteEmailInput(_Base):
    entry_id: str = Field(min_length=1)
    permanent: bool = False


class JunkEmailInput(_Base):
    entry_id: str = Field(min_length=1)


class CreateFolderInput(_Base):
    parent_path: str | None = None
    name: str = Field(min_length=1)


class RenameFolderInput(_Base):
    folder_path: str = Field(min_length=1)
    new_name: str = Field(min_length=1)


class MoveFolderInput(_Base):
    folder_path: str = Field(min_length=1)
    new_parent_path: str = Field(min_length=1)


class DeleteFolderInput(_Base):
    folder_path: str = Field(min_length=1)


class EmptyFolderInput(_Base):
    folder_path: str = Field(min_length=1)


class CreateCalendarEventInput(_Base):
    subject: str = Field(min_length=1)
    start_utc: datetime
    end_utc: datetime
    attendees: list[str] = Field(default_factory=list)
    location: str = ""
    body: str = ""
    is_all_day: bool = False


class UpdateCalendarEventInput(_Base):
    event_id: str = Field(min_length=1)
    subject: str | None = None
    start_utc: datetime | None = None
    end_utc: datetime | None = None
    location: str | None = None
    body: str | None = None


class DeleteCalendarEventInput(_Base):
    event_id: str = Field(min_length=1)


class RespondToEventInput(_Base):
    event_id: str = Field(min_length=1)
    response: Literal["accept", "tentative", "decline"]
    send_response: bool = True


class ListTasksInput(_Base):
    limit: int = Field(default=100, ge=1, le=500)
    offset: int = Field(default=0, ge=0)


class SearchTasksInput(_Base):
    query: str = Field(min_length=1)
    limit: int = Field(default=50, ge=1, le=500)


class GetTaskInput(_Base):
    task_id: str = Field(min_length=1)


class ListNotesInput(_Base):
    limit: int = Field(default=100, ge=1, le=500)
    offset: int = Field(default=0, ge=0)


class SearchNotesInput(_Base):
    query: str = Field(min_length=1)
    limit: int = Field(default=50, ge=1, le=500)


class GetNoteInput(_Base):
    note_id: str = Field(min_length=1)


class GetUnreadCountInput(_Base):
    folder: str | None = Field(
        default=None, description="Specific folder; omit for total across all folders.")


# ---------------------------------------------------------------------------
# Phase 4 (Exchange-specific) input models
# ---------------------------------------------------------------------------


class GetOutOfOfficeInput(_Base):
    pass


class SetOutOfOfficeInput(_Base):
    enabled: bool
    internal_message: str = ""
    external_message: str = ""
    start_utc: datetime | None = None
    end_utc: datetime | None = None
    external_audience: Literal["all", "contacts_only", "none"] = "all"


class GetSignatureInput(_Base):
    account_id: str | None = None


class SetSignatureInput(_Base):
    body_html: str
    body_plain: str = ""
    account_id: str | None = None


class ListRulesInput(_Base):
    pass


class ToggleRuleInput(_Base):
    rule_id: str = Field(min_length=1)
    enabled: bool = True


class CalendarFreeBusyInput(_Base):
    smtps: list[str] = Field(min_length=1)
    start_utc: datetime
    end_utc: datetime
    slot_minutes: int = Field(default=30, ge=5, le=240)


class MeetingRoomFinderInput(_Base):
    start_utc: datetime
    end_utc: datetime
    capacity: int = Field(default=10, ge=1, le=500)
    location_hint: str | None = None


class GalSearchInput(_Base):
    query: str = Field(min_length=1)
    limit: int = Field(default=25, ge=1, le=200)
