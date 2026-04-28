"""Windows COM backend backed by pywin32.

All COM calls are marshalled through `OutlookComBridge` onto the STA
thread. Helpers `_sync_*` run on that thread and take `namespace` as their
first argument.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from ..baseline import load_baseline
from ..com_bridge import OutlookComBridge
from ..constants import (
    ATTACH_TYPES_REAL,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_CONTACTS,
    OL_FOLDER_DELETED,
    OL_FOLDER_INBOX,
    OL_FOLDER_JUNK,
)
from ..errors import (
    AttachmentError,
    ConversationNotFoundError,
    EmailNotFoundError,
    FolderNotFoundError,
)
from ..utils.email_extract import detect_importance, domain_of, preview_text
from ..utils.filter_utils import build_time_range_restrict, escape_dasl
from ..utils.folder_utils import canonicalize_folder_name, split_folder_path
from ..utils.mapi_props import (
    PR_IN_REPLY_TO_ID,
    PR_INTERNET_MESSAGE_ID,
    PR_INTERNET_REFERENCES,
    PR_SENT_REPRESENTING_SMTP,
    PR_TRANSPORT_HEADERS,
    safe_get,
)
from ..utils.smtp_resolver import resolve_recipient, resolve_sender
from ..utils.subject_utils import normalize_subject
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
    FreeBusySlot,
    MailboxInfo,
    MailboxQuota,
    NoteInfo,
    OperationResult,
    OutOfOfficeStatus,
    RuleInfo,
    SignatureInfo,
    TaskInfo,
    ThreadMetadata,
)


class WindowsComBackend(Backend):
    def __init__(self) -> None:
        self._bridge = OutlookComBridge()

    # ------------------------------------------------------------------
    # Mailbox / folders
    # ------------------------------------------------------------------
    async def get_mailbox_info(self) -> MailboxInfo:
        return await self._bridge.execute(_sync_get_mailbox_info)

    async def list_folders(self, include_subfolders: bool) -> list[FolderInfo]:
        return await self._bridge.execute(_sync_list_folders, include_subfolders)

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
        return await self._bridge.execute(
            _sync_get_emails_in_time_range, start_utc, end_utc, folders, max_results
        )

    async def get_conversation_thread(
        self,
        conversation_id: str,
        conversation_topic: str | None,
        max_messages: int,
    ) -> list[EmailSummary]:
        return await self._bridge.execute(
            _sync_get_conversation_thread,
            conversation_id,
            conversation_topic,
            max_messages,
        )

    async def get_email_full_metadata(self, entry_id: str) -> EmailFullMetadata:
        return await self._bridge.execute(_sync_get_email_full_metadata, entry_id)

    async def save_attachment(
        self, entry_id: str, attachment_index: int, save_path: str
    ) -> AttachmentInfo:
        return await self._bridge.execute(
            _sync_save_attachment, entry_id, attachment_index, save_path
        )

    async def search_emails(
        self, query: str, folders: list[str], field: str, max_results: int
    ) -> list[EmailSummary]:
        return await self._bridge.execute(
            _sync_search_emails, query, folders, field, max_results
        )

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
        return await self._bridge.execute(
            _sync_list_calendar_events, start_utc, end_utc, calendar_id, max_results
        )

    async def get_calendar_event(self, event_id: str) -> CalendarEvent:
        return await self._bridge.execute(_sync_get_calendar_event, event_id)

    # ------------------------------------------------------------------
    # Contacts
    # ------------------------------------------------------------------
    async def list_contacts(self, limit: int, offset: int) -> list[Contact]:
        return await self._bridge.execute(_sync_list_contacts, limit, offset)

    async def search_contacts(self, query: str, limit: int) -> list[Contact]:
        return await self._bridge.execute(_sync_search_contacts, query, limit)

    async def get_contact(self, contact_id: str) -> Contact:
        return await self._bridge.execute(_sync_get_contact, contact_id)

    async def diagnostics(self, sample_message_id: str | None = None) -> DiagnosticsResult:
        return await self._bridge.execute(_sync_diagnostics, sample_message_id)

    # ------------------------------------------------------------------
    # Mail write
    # ------------------------------------------------------------------
    async def send_email(self, to, subject, body, body_type, cc, bcc, attachments) -> OperationResult:
        return await self._bridge.execute(_sync_send_email, to, subject, body, body_type, cc, bcc, attachments, send=True)

    async def create_draft(self, to, subject, body, body_type, cc, bcc, attachments) -> OperationResult:
        return await self._bridge.execute(_sync_send_email, to, subject, body, body_type, cc, bcc, attachments, send=False)

    async def reply_email(self, entry_id, body, body_type, reply_all) -> OperationResult:
        return await self._bridge.execute(_sync_reply_email, entry_id, body, body_type, reply_all)

    async def forward_email(self, entry_id, to, body, body_type) -> OperationResult:
        return await self._bridge.execute(_sync_forward_email, entry_id, to, body, body_type)

    # ------------------------------------------------------------------
    # Mail organize
    # ------------------------------------------------------------------
    async def mark_email_read(self, entry_id, is_read) -> OperationResult:
        return await self._bridge.execute(_sync_mark_email_read, entry_id, is_read)

    async def set_email_flag(self, entry_id, flag_status, due_date_utc) -> OperationResult:
        return await self._bridge.execute(_sync_set_email_flag, entry_id, flag_status, due_date_utc)

    async def set_email_categories(self, entry_id, categories) -> OperationResult:
        return await self._bridge.execute(_sync_set_email_categories, entry_id, categories)

    # ------------------------------------------------------------------
    # Mail destructive
    # ------------------------------------------------------------------
    async def move_email(self, entry_id, destination_folder) -> OperationResult:
        return await self._bridge.execute(_sync_move_email, entry_id, destination_folder)

    async def archive_email(self, entry_id) -> OperationResult:
        return await self._bridge.execute(_sync_move_email, entry_id, "Archive")

    async def delete_email(self, entry_id, permanent) -> OperationResult:
        return await self._bridge.execute(_sync_delete_email, entry_id, permanent)

    async def junk_email(self, entry_id) -> OperationResult:
        return await self._bridge.execute(_sync_junk_email, entry_id)

    # ------------------------------------------------------------------
    # Folder management
    # ------------------------------------------------------------------
    async def create_folder(self, parent_path, name) -> OperationResult:
        return await self._bridge.execute(_sync_create_folder, parent_path, name)

    async def rename_folder(self, folder_path, new_name) -> OperationResult:
        return await self._bridge.execute(_sync_rename_folder, folder_path, new_name)

    async def move_folder(self, folder_path, new_parent_path) -> OperationResult:
        return await self._bridge.execute(_sync_move_folder, folder_path, new_parent_path)

    async def delete_folder(self, folder_path) -> OperationResult:
        return await self._bridge.execute(_sync_delete_folder, folder_path)

    async def empty_folder(self, folder_path) -> OperationResult:
        return await self._bridge.execute(_sync_empty_folder, folder_path)

    # ------------------------------------------------------------------
    # Calendar write
    # ------------------------------------------------------------------
    async def create_calendar_event(self, subject, start_utc, end_utc, attendees, location, body, is_all_day) -> OperationResult:
        return await self._bridge.execute(_sync_create_calendar_event, subject, start_utc, end_utc, attendees, location, body, is_all_day)

    async def update_calendar_event(self, event_id, subject, start_utc, end_utc, location, body) -> OperationResult:
        return await self._bridge.execute(_sync_update_calendar_event, event_id, subject, start_utc, end_utc, location, body)

    async def delete_calendar_event(self, event_id) -> OperationResult:
        return await self._bridge.execute(_sync_delete_calendar_event, event_id)

    async def respond_to_event(self, event_id, response, send_response) -> OperationResult:
        return await self._bridge.execute(_sync_respond_to_event, event_id, response, send_response)

    # ------------------------------------------------------------------
    # Tasks
    # ------------------------------------------------------------------
    async def list_tasks(self, limit, offset) -> list[TaskInfo]:
        return await self._bridge.execute(_sync_list_tasks, limit, offset)

    async def search_tasks(self, query, limit) -> list[TaskInfo]:
        return await self._bridge.execute(_sync_search_tasks, query, limit)

    async def get_task(self, task_id) -> TaskInfo:
        return await self._bridge.execute(_sync_get_task, task_id)

    # ------------------------------------------------------------------
    # Notes
    # ------------------------------------------------------------------
    async def list_notes(self, limit, offset) -> list[NoteInfo]:
        return await self._bridge.execute(_sync_list_notes, limit, offset)

    async def search_notes(self, query, limit) -> list[NoteInfo]:
        return await self._bridge.execute(_sync_search_notes, query, limit)

    async def get_note(self, note_id) -> NoteInfo:
        return await self._bridge.execute(_sync_get_note, note_id)

    # ------------------------------------------------------------------
    # Accounts
    # ------------------------------------------------------------------
    async def list_accounts(self) -> list[AccountInfo]:
        return await self._bridge.execute(_sync_list_accounts)

    async def get_unread_count(self, folder) -> int:
        return await self._bridge.execute(_sync_get_unread_count, folder)

    # ------------------------------------------------------------------
    # Exchange-specific
    # ------------------------------------------------------------------
    async def get_out_of_office(self) -> OutOfOfficeStatus:
        return await self._bridge.execute(_sync_get_out_of_office)

    async def set_out_of_office(self, status: OutOfOfficeStatus) -> OperationResult:
        return await self._bridge.execute(_sync_set_out_of_office, status)

    async def get_signature(self, account_id: str | None) -> SignatureInfo:
        return await self._bridge.execute(_sync_get_signature, account_id)

    async def set_signature(self, account_id: str | None, body_html: str, body_plain: str) -> OperationResult:
        return await self._bridge.execute(_sync_set_signature, account_id, body_html, body_plain)

    async def list_rules(self) -> list[RuleInfo]:
        return await self._bridge.execute(_sync_list_rules)

    async def toggle_rule(self, rule_id: str, enabled: bool) -> OperationResult:
        return await self._bridge.execute(_sync_toggle_rule, rule_id, enabled)

    async def calendar_freebusy(self, smtps, start_utc, end_utc, slot_minutes) -> list[FreeBusyResponse]:
        return await self._bridge.execute(_sync_calendar_freebusy, smtps, start_utc, end_utc, slot_minutes)

    async def meeting_room_finder(self, start_utc, end_utc, capacity, location_hint) -> list[Contact]:
        return await self._bridge.execute(_sync_meeting_room_finder, start_utc, end_utc, capacity, location_hint)

    async def gal_search(self, query: str, limit: int) -> list[Contact]:
        return await self._bridge.execute(_sync_gal_search, query, limit)

    async def list_delegated_mailboxes(self) -> list[MailboxInfo]:
        return await self._bridge.execute(_sync_list_delegated_mailboxes)

    async def list_public_folders(self) -> list[FolderInfo]:
        return await self._bridge.execute(_sync_list_public_folders)

    async def get_mailbox_quota(self) -> MailboxQuota:
        return await self._bridge.execute(_sync_get_mailbox_quota)

    async def close(self) -> None:
        self._bridge.close()


# ---------------------------------------------------------------------------
# Synchronous helpers (STA thread).
# ---------------------------------------------------------------------------


def _resolve_folder(namespace: Any, path: str) -> Any:
    canonical = canonicalize_folder_name(path)
    parts = split_folder_path(canonical) or [canonical]
    try:
        folder = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
    except Exception:
        folder = None
    if folder is not None:
        for part in parts:
            try:
                folder = folder.Folders[part]
            except Exception:
                folder = None
                break
        if folder is not None:
            return folder
    raise FolderNotFoundError(
        f"Folder {path!r} not found. Call outlook_list_folders to see available folders."
    )


def _pywintime_to_utc(value: Any) -> datetime | None:
    if value is None:
        return None
    try:
        year = value.year
    except Exception:
        return None
    try:
        return datetime(
            value.year,
            value.month,
            value.day,
            value.hour,
            value.minute,
            value.second,
            tzinfo=timezone.utc,
        )
    except Exception:
        return None


def _sync_get_mailbox_info(namespace: Any) -> MailboxInfo:
    user = namespace.CurrentUser
    smtp = ""
    try:
        smtp = safe_get(user.AddressEntry, PR_SENT_REPRESENTING_SMTP) or ""
    except Exception:
        smtp = ""
    if not smtp:
        try:
            exchange_user = user.AddressEntry.GetExchangeUser()
            smtp = exchange_user.PrimarySmtpAddress if exchange_user else ""
        except Exception:
            smtp = ""
    display = ""
    try:
        display = user.Name or ""
    except Exception:
        display = ""
    return MailboxInfo(
        display_name=display,
        smtp_address=smtp,
        domain=domain_of(smtp),
    )


def _sync_list_folders(namespace: Any, include_subfolders: bool) -> list[FolderInfo]:
    results: list[FolderInfo] = []

    def walk(folder: Any, parent_path: str | None) -> None:
        name = folder.Name
        full_path = name if parent_path is None else f"{parent_path}/{name}"
        try:
            item_count = int(folder.Items.Count)
        except Exception:
            item_count = 0
        try:
            unread_count = int(folder.UnReadItemCount)
        except Exception:
            unread_count = 0
        results.append(
            FolderInfo(
                name=name,
                full_path=full_path,
                item_count=item_count,
                unread_count=unread_count,
                parent_path=parent_path,
            )
        )
        if not include_subfolders:
            return
        for sub in folder.Folders:
            walk(sub, full_path)

    root = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
    for folder in root.Folders:
        walk(folder, None)
    return results


def _extract_summary(item: Any, folder_path: str) -> EmailSummary:
    sender_smtp = resolve_sender(item)
    to_list: list[str] = []
    cc_list: list[str] = []
    try:
        for recip in item.Recipients:
            smtp = resolve_recipient(recip)
            if not smtp:
                continue
            kind = getattr(recip, "Type", 1)
            if kind == 2:
                cc_list.append(smtp)
            else:
                to_list.append(smtp)
    except Exception:
        pass
    try:
        received = _pywintime_to_utc(item.ReceivedTime)
    except Exception:
        received = None
    try:
        sent = _pywintime_to_utc(item.SentOn)
    except Exception:
        sent = None
    try:
        conv_id = item.ConversationID or ""
    except Exception:
        conv_id = ""
    try:
        has_real = any(
            int(getattr(a, "Type", 0)) in ATTACH_TYPES_REAL for a in item.Attachments
        )
    except Exception:
        has_real = False
    try:
        body_plain = item.Body or ""
    except Exception:
        body_plain = ""
    return EmailSummary(
        entry_id=item.EntryID,
        conversation_id=conv_id,
        subject=getattr(item, "Subject", "") or "",
        sender_name=getattr(item, "SenderName", "") or "",
        sender_smtp=sender_smtp,
        to_smtp=to_list,
        cc_smtp=cc_list,
        received_utc=received,
        sent_utc=sent,
        is_read=not bool(getattr(item, "UnRead", False)),
        importance=detect_importance(getattr(item, "Importance", 1)),
        has_real_attachments=has_real,
        folder_path=folder_path,
        preview=preview_text(body_plain),
    )


def _sync_get_emails_in_time_range(
    namespace: Any,
    start_utc: datetime,
    end_utc: datetime,
    folders: list[str],
    max_results: int,
) -> list[EmailSummary]:
    results: list[EmailSummary] = []
    for folder_name in folders:
        folder = _resolve_folder(namespace, folder_name)
        field = "SentOn" if "sent" in folder_name.lower() else "ReceivedTime"
        restrict = build_time_range_restrict(field, start_utc, end_utc)
        items = folder.Items
        try:
            items.Sort(f"[{field}]", True)
        except Exception:
            pass
        try:
            filtered = items.Restrict(restrict)
        except Exception as exc:
            raise EmailNotFoundError(
                f"Restrict failed on {folder_name!r}: {exc}"
            ) from exc
        for item in filtered:
            if len(results) >= max_results:
                break
            if getattr(item, "Class", 43) != 43:  # olMail
                continue
            results.append(_extract_summary(item, folder_name))
        if len(results) >= max_results:
            break
    return results


def _sync_get_conversation_thread(
    namespace: Any,
    conversation_id: str,
    conversation_topic: str | None,
    max_messages: int,
) -> list[EmailSummary]:
    collected: list[EmailSummary] = []
    root = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
    for folder in root.Folders:
        if len(collected) >= max_messages:
            break
        try:
            items = folder.Items
            if conversation_topic:
                try:
                    items = items.Restrict(
                        f"[ConversationTopic] = '{escape_dasl(conversation_topic)}'"
                    )
                except Exception:
                    pass
            for item in items:
                if len(collected) >= max_messages:
                    break
                if getattr(item, "Class", 43) != 43:
                    continue
                try:
                    if item.ConversationID != conversation_id:
                        continue
                except Exception:
                    continue
                collected.append(_extract_summary(item, folder.Name))
        except Exception:
            continue
    if not collected:
        raise ConversationNotFoundError(
            f"No messages found for conversation_id {conversation_id!r}."
        )
    collected.sort(key=lambda s: s.sent_utc or s.received_utc or datetime.min.replace(tzinfo=timezone.utc))
    return collected


def _sync_get_email_full_metadata(namespace: Any, entry_id: str) -> EmailFullMetadata:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    folder_path = ""
    try:
        folder_path = item.Parent.Name if item.Parent else ""
    except Exception:
        folder_path = ""
    summary = _extract_summary(item, folder_path)
    try:
        plain = item.Body or ""
    except Exception:
        plain = ""
    try:
        html = item.HTMLBody or None
    except Exception:
        html = None
    attachments: list[AttachmentInfo] = []
    try:
        for idx, att in enumerate(item.Attachments, start=1):
            kind = int(getattr(att, "Type", 0))
            attachments.append(
                AttachmentInfo(
                    index=idx,
                    filename=getattr(att, "FileName", "") or "",
                    size_bytes=int(getattr(att, "Size", 0) or 0),
                    is_inline=kind not in ATTACH_TYPES_REAL,
                )
            )
    except Exception:
        pass

    msg_id = safe_get(item, PR_INTERNET_MESSAGE_ID)
    in_reply_to = safe_get(item, PR_IN_REPLY_TO_ID)
    references = safe_get(item, PR_INTERNET_REFERENCES)
    headers_raw = safe_get(item, PR_TRANSPORT_HEADERS)
    references_list: list[str] | None = None
    if references:
        references_list = [r for r in references.split() if r]

    return EmailFullMetadata(
        summary=summary,
        body=EmailBody(plain_text=plain, html=html),
        attachments=attachments,
        internet_message_id=msg_id or None,
        in_reply_to=in_reply_to or None,
        references=references_list,
        mapi_headers_raw=headers_raw or None,
        delegation_sender_smtp=summary.sender_smtp or None,
        delegation_representing_smtp=(
            safe_get(item, PR_SENT_REPRESENTING_SMTP) or None
        ),
    )


def _sync_save_attachment(
    namespace: Any, entry_id: str, attachment_index: int, save_path: str
) -> AttachmentInfo:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        att = item.Attachments.Item(attachment_index)
    except Exception as exc:
        raise AttachmentError(
            f"Attachment index {attachment_index} not found on email {entry_id!r}."
        ) from exc
    kind = int(getattr(att, "Type", 0))
    if kind not in ATTACH_TYPES_REAL:
        raise AttachmentError(
            f"Attachment {attachment_index} on {entry_id!r} is inline (type={kind}); refusing to save."
        )
    try:
        att.SaveAsFile(save_path)
    except Exception as exc:
        raise AttachmentError(f"SaveAsFile failed for {save_path!r}: {exc}") from exc
    return AttachmentInfo(
        index=attachment_index,
        filename=getattr(att, "FileName", "") or "",
        size_bytes=int(getattr(att, "Size", 0) or 0),
        is_inline=False,
    )


def _sync_search_emails(
    namespace: Any,
    query: str,
    folders: list[str],
    field: str,
    max_results: int,
) -> list[EmailSummary]:
    q = escape_dasl(query)
    if field == "subject":
        restrict = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{q}%'"
    elif field == "sender":
        restrict = f"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%{q}%'"
    elif field == "body":
        restrict = f"@SQL=\"urn:schemas:httpmail:textdescription\" LIKE '%{q}%'"
    else:
        restrict = (
            f"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{q}%' "
            f"OR \"urn:schemas:httpmail:fromemail\" LIKE '%{q}%' "
            f"OR \"urn:schemas:httpmail:textdescription\" LIKE '%{q}%')"
        )
    results: list[EmailSummary] = []
    for folder_name in folders:
        folder = _resolve_folder(namespace, folder_name)
        try:
            filtered = folder.Items.Restrict(restrict)
        except Exception:
            continue
        for item in filtered:
            if len(results) >= max_results:
                break
            if getattr(item, "Class", 43) != 43:
                continue
            results.append(_extract_summary(item, folder_name))
        if len(results) >= max_results:
            break
    return results


# ---------------------------------------------------------------------------
# Calendar
# ---------------------------------------------------------------------------


def _extract_calendar_event(item: Any) -> CalendarEvent:
    attendees: list[str] = []
    try:
        for recip in item.Recipients:
            smtp = resolve_recipient(recip)
            if smtp:
                attendees.append(smtp)
    except Exception:
        pass
    organizer = ""
    try:
        organizer = resolve_sender(item)
    except Exception:
        organizer = ""
    try:
        start = _pywintime_to_utc(item.StartUTC)
    except Exception:
        try:
            start = _pywintime_to_utc(item.Start)
        except Exception:
            start = None
    try:
        end = _pywintime_to_utc(item.EndUTC)
    except Exception:
        try:
            end = _pywintime_to_utc(item.End)
        except Exception:
            end = None
    return CalendarEvent(
        event_id=item.EntryID,
        subject=getattr(item, "Subject", "") or "",
        organizer_smtp=organizer,
        start_utc=start,
        end_utc=end,
        location=getattr(item, "Location", "") or "",
        is_all_day=bool(getattr(item, "AllDayEvent", False)),
        attendees_smtp=attendees,
        body_plain=getattr(item, "Body", "") or "",
        is_recurring=bool(getattr(item, "IsRecurring", False)),
    )


def _sync_list_calendar_events(
    namespace: Any,
    start_utc: datetime,
    end_utc: datetime,
    calendar_id: str | None,
    max_results: int,
) -> list[CalendarEvent]:
    calendar = (
        namespace.GetFolderFromID(calendar_id)
        if calendar_id
        else namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
    )
    items = calendar.Items
    try:
        items.IncludeRecurrences = True
        items.Sort("[Start]", False)
    except Exception:
        pass
    restrict = build_time_range_restrict("Start", start_utc, end_utc)
    try:
        filtered = items.Restrict(restrict)
    except Exception as exc:
        raise EmailNotFoundError(f"Calendar Restrict failed: {exc}") from exc
    results: list[CalendarEvent] = []
    for item in filtered:
        if len(results) >= max_results:
            break
        results.append(_extract_calendar_event(item))
    return results


def _sync_get_calendar_event(namespace: Any, event_id: str) -> CalendarEvent:
    try:
        item = namespace.GetItemFromID(event_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No calendar event with id {event_id!r}.") from exc
    return _extract_calendar_event(item)


# ---------------------------------------------------------------------------
# Contacts
# ---------------------------------------------------------------------------


def _extract_contact(item: Any) -> Contact:
    phones: dict[str, str] = {}
    for label, attr in (
        ("business", "BusinessTelephoneNumber"),
        ("home", "HomeTelephoneNumber"),
        ("mobile", "MobileTelephoneNumber"),
    ):
        value = getattr(item, attr, "") or ""
        if value:
            phones[label] = value
    other: list[str] = []
    for attr in ("Email2Address", "Email3Address"):
        value = getattr(item, attr, "") or ""
        if value:
            other.append(value)
    return Contact(
        contact_id=item.EntryID,
        display_name=getattr(item, "FullName", "") or getattr(item, "FileAs", "") or "",
        primary_smtp=getattr(item, "Email1Address", "") or "",
        other_smtps=other,
        company=getattr(item, "CompanyName", "") or "",
        job_title=getattr(item, "JobTitle", "") or "",
        phone_numbers=phones,
    )


def _sync_list_contacts(namespace: Any, limit: int, offset: int) -> list[Contact]:
    folder = namespace.GetDefaultFolder(OL_FOLDER_CONTACTS)
    results: list[Contact] = []
    idx = 0
    for item in folder.Items:
        if getattr(item, "Class", 40) != 40:  # olContact
            continue
        if idx < offset:
            idx += 1
            continue
        if len(results) >= limit:
            break
        results.append(_extract_contact(item))
        idx += 1
    return results


def _sync_search_contacts(namespace: Any, query: str, limit: int) -> list[Contact]:
    q = query.lower()
    folder = namespace.GetDefaultFolder(OL_FOLDER_CONTACTS)
    results: list[Contact] = []
    for item in folder.Items:
        if getattr(item, "Class", 40) != 40:
            continue
        haystack = " ".join(
            (
                getattr(item, "FullName", "") or "",
                getattr(item, "Email1Address", "") or "",
                getattr(item, "CompanyName", "") or "",
                getattr(item, "JobTitle", "") or "",
            )
        ).lower()
        if q in haystack:
            results.append(_extract_contact(item))
            if len(results) >= limit:
                break
    return results


def _sync_get_contact(namespace: Any, contact_id: str) -> Contact:
    try:
        item = namespace.GetItemFromID(contact_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No contact with id {contact_id!r}.") from exc
    return _extract_contact(item)


# ---------------------------------------------------------------------------
# Diagnostics
# ---------------------------------------------------------------------------

# Map from "field that broke" → list of tools impacted, used to populate
# DiagnosticsResult.affected_tools when a probe fails.
_FIELD_TO_TOOLS: dict[str, list[str]] = {
    "Inbox": ["outlook_get_emails_in_time_range", "outlook_search_emails"],
    "Calendar": ["outlook_list_calendar_events", "outlook_get_calendar_event"],
    "Contacts": ["outlook_list_contacts", "outlook_search_contacts", "outlook_get_contact"],
    "PR_INTERNET_MESSAGE_ID": ["outlook_get_email_full_metadata"],
    "PR_IN_REPLY_TO_ID": ["outlook_get_email_full_metadata"],
    "PR_INTERNET_REFERENCES": ["outlook_get_email_full_metadata"],
    "PR_TRANSPORT_HEADERS": ["outlook_get_email_full_metadata"],
    "ConversationID": [
        "outlook_get_conversation_thread",
        "outlook_get_thread_metadata",
        "outlook_get_emails_in_time_range",
    ],
    "PropertyAccessor": ["outlook_get_email_full_metadata"],
}


def _sync_diagnostics(namespace: Any, sample_message_id: str | None) -> DiagnosticsResult:
    baseline = load_baseline()
    notes: list[str] = []
    probed: dict[str, str] = {}
    affected: set[str] = set()

    # Outlook version
    try:
        outlook_app = namespace.Application
        version = str(outlook_app.Version)
    except Exception as exc:
        version = f"unknown (error: {exc})"

    # Default folder probes
    defaults = baseline.get("windows_com", {}).get("default_folders", {})
    for name, idx in defaults.items():
        try:
            folder = namespace.GetDefaultFolder(idx)
            _ = folder.Name
            probed[f"OlDefaultFolders.{name}"] = "ok"
        except Exception as exc:
            probed[f"OlDefaultFolders.{name}"] = f"error:{exc}"
            for t in _FIELD_TO_TOOLS.get(name, []):
                affected.add(t)

    # Pick a sample message if not supplied: first item in Inbox.
    sample_item = None
    if sample_message_id:
        try:
            sample_item = namespace.GetItemFromID(sample_message_id)
        except Exception as exc:
            notes.append(f"Could not load supplied sample id: {exc}")
    if sample_item is None:
        try:
            inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            if items.Count > 0:
                sample_item = items.GetFirst()
        except Exception as exc:
            notes.append(f"Could not auto-pick sample message: {exc}")

    if sample_item is None:
        notes.append("No sample message available; per-message field probes skipped.")
    else:
        # Item-level property probes
        for prop in baseline.get("windows_com", {}).get("item_properties", []):
            try:
                _ = getattr(sample_item, prop)
                probed[f"MailItem.{prop}"] = "ok"
            except Exception as exc:
                probed[f"MailItem.{prop}"] = f"error:{exc}"
                for t in _FIELD_TO_TOOLS.get(prop, []):
                    affected.add(t)
        # MAPI tag probes
        for tag_name, tag_value in (
            ("PR_INTERNET_MESSAGE_ID", PR_INTERNET_MESSAGE_ID),
            ("PR_IN_REPLY_TO_ID", PR_IN_REPLY_TO_ID),
            ("PR_INTERNET_REFERENCES", PR_INTERNET_REFERENCES),
            ("PR_TRANSPORT_HEADERS", PR_TRANSPORT_HEADERS),
            ("PR_SENT_REPRESENTING_SMTP", PR_SENT_REPRESENTING_SMTP),
        ):
            value = safe_get(sample_item, tag_value)
            probed[f"MAPI.{tag_name}"] = "ok" if value is not None else "missing"
            if value is None:
                for t in _FIELD_TO_TOOLS.get(tag_name, []):
                    affected.add(t)

    missing = [k for k, v in probed.items() if v != "ok"]
    if not missing:
        status = "HEALTHY"
    elif len(missing) >= len(probed) // 2:
        status = "BROKEN"
    else:
        status = f"DEGRADED:{len(missing)}_fields"

    baseline_version = baseline.get("baseline_outlook_versions", {}).get("windows", "unknown")
    return DiagnosticsResult(
        platform="windows",
        outlook_version=version,
        baseline_version=baseline_version,
        status=status,
        probed_fields=probed,
        affected_tools=sorted(affected),
        notes=notes,
    )


# ---------------------------------------------------------------------------
# Phase 3 (parity) sync helpers
# ---------------------------------------------------------------------------

OL_TASK_FOLDER_INDEX = 13  # OlDefaultFolders.olFolderTasks
OL_NOTES_FOLDER_INDEX = 12  # OlDefaultFolders.olFolderNotes


def _make_mail_item(namespace: Any, to, subject, body, body_type, cc, bcc, attachments) -> Any:
    app = namespace.Application
    item = app.CreateItem(0)  # olMailItem
    item.Subject = subject or ""
    if (body_type or "plain").lower() == "html":
        item.HTMLBody = body or ""
    else:
        item.Body = body or ""
    if to:
        item.To = "; ".join(to)
    if cc:
        item.CC = "; ".join(cc)
    if bcc:
        item.BCC = "; ".join(bcc)
    for path in attachments or []:
        try:
            item.Attachments.Add(path)
        except Exception:
            pass
    return item


def _sync_send_email(namespace: Any, to, subject, body, body_type, cc, bcc, attachments, send: bool) -> OperationResult:
    item = _make_mail_item(namespace, to, subject, body, body_type, cc, bcc, attachments)
    try:
        if send:
            item.Send()
            return OperationResult(ok=True, message="sent", affected_id=getattr(item, "EntryID", None))
        item.Save()
        return OperationResult(ok=True, message="saved as draft", affected_id=getattr(item, "EntryID", None))
    except Exception as exc:
        return OperationResult(ok=False, message=f"send/save failed: {exc}")


def _sync_reply_email(namespace: Any, entry_id: str, body: str, body_type: str, reply_all: bool) -> OperationResult:
    try:
        original = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    reply = original.ReplyAll() if reply_all else original.Reply()
    if (body_type or "plain").lower() == "html":
        reply.HTMLBody = (body or "") + (reply.HTMLBody or "")
    else:
        reply.Body = (body or "") + "\n\n" + (reply.Body or "")
    try:
        reply.Send()
    except Exception as exc:
        return OperationResult(ok=False, message=f"reply send failed: {exc}")
    return OperationResult(ok=True, message="replied")


def _sync_forward_email(namespace: Any, entry_id: str, to: list[str], body: str, body_type: str) -> OperationResult:
    try:
        original = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    fwd = original.Forward()
    fwd.To = "; ".join(to)
    if (body_type or "plain").lower() == "html":
        fwd.HTMLBody = (body or "") + (fwd.HTMLBody or "")
    else:
        fwd.Body = (body or "") + "\n\n" + (fwd.Body or "")
    try:
        fwd.Send()
    except Exception as exc:
        return OperationResult(ok=False, message=f"forward send failed: {exc}")
    return OperationResult(ok=True, message="forwarded")


def _sync_mark_email_read(namespace: Any, entry_id: str, is_read: bool) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        item.UnRead = not is_read
        item.Save()
        return OperationResult(ok=True, message=f"marked {'read' if is_read else 'unread'}", affected_id=entry_id)
    except Exception as exc:
        return OperationResult(ok=False, message=f"mark failed: {exc}")


def _sync_set_email_flag(namespace: Any, entry_id: str, flag_status: str, due_date_utc) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        # FlagStatus: 0 = none, 1 = complete, 2 = marked
        if flag_status == "complete":
            item.FlagStatus = 1
        elif flag_status == "marked":
            item.FlagStatus = 2
        else:
            item.FlagStatus = 0
        if due_date_utc is not None:
            try:
                item.TaskDueDate = due_date_utc
            except Exception:
                pass
        item.Save()
        return OperationResult(ok=True, message=f"flag set to {flag_status}", affected_id=entry_id)
    except Exception as exc:
        return OperationResult(ok=False, message=f"flag set failed: {exc}")


def _sync_set_email_categories(namespace: Any, entry_id: str, categories: list[str]) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        item.Categories = ", ".join(categories) if categories else ""
        item.Save()
        return OperationResult(ok=True, message="categories set", affected_id=entry_id)
    except Exception as exc:
        return OperationResult(ok=False, message=f"category set failed: {exc}")


def _sync_move_email(namespace: Any, entry_id: str, destination_folder: str) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        target = _resolve_folder(namespace, destination_folder)
        moved = item.Move(target)
        return OperationResult(ok=True, message=f"moved to {destination_folder}",
                               affected_id=getattr(moved, "EntryID", None))
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"move failed: {exc}")


def _sync_delete_email(namespace: Any, entry_id: str, permanent: bool) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        if permanent:
            item.Delete()
            return OperationResult(ok=True, message="permanently deleted")
        deleted = namespace.GetDefaultFolder(OL_FOLDER_DELETED)
        item.Move(deleted)
        return OperationResult(ok=True, message="moved to Deleted Items")
    except Exception as exc:
        return OperationResult(ok=False, message=f"delete failed: {exc}")


def _sync_junk_email(namespace: Any, entry_id: str) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        junk = namespace.GetDefaultFolder(OL_FOLDER_JUNK)
        item.Move(junk)
        return OperationResult(ok=True, message="moved to Junk")
    except Exception as exc:
        return OperationResult(ok=False, message=f"junk failed: {exc}")


def _sync_create_folder(namespace: Any, parent_path: str | None, name: str) -> OperationResult:
    try:
        if parent_path:
            parent = _resolve_folder(namespace, parent_path)
        else:
            parent = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
        new_folder = parent.Folders.Add(name)
        return OperationResult(ok=True, message="created", affected_id=getattr(new_folder, "EntryID", None))
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"create folder failed: {exc}")


def _sync_rename_folder(namespace: Any, folder_path: str, new_name: str) -> OperationResult:
    try:
        f = _resolve_folder(namespace, folder_path)
        f.Name = new_name
        return OperationResult(ok=True, message="renamed")
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"rename failed: {exc}")


def _sync_move_folder(namespace: Any, folder_path: str, new_parent_path: str) -> OperationResult:
    try:
        f = _resolve_folder(namespace, folder_path)
        new_parent = _resolve_folder(namespace, new_parent_path)
        f.MoveTo(new_parent)
        return OperationResult(ok=True, message="moved folder")
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"move folder failed: {exc}")


def _sync_delete_folder(namespace: Any, folder_path: str) -> OperationResult:
    try:
        f = _resolve_folder(namespace, folder_path)
        f.Delete()
        return OperationResult(ok=True, message="deleted folder")
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"delete folder failed: {exc}")


def _sync_empty_folder(namespace: Any, folder_path: str) -> OperationResult:
    try:
        f = _resolve_folder(namespace, folder_path)
        items = list(f.Items)  # snapshot before mutation
        for item in items:
            try:
                item.Delete()
            except Exception:
                pass
        return OperationResult(ok=True, message=f"emptied ({len(items)} items)")
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"empty folder failed: {exc}")


def _sync_create_calendar_event(namespace: Any, subject, start_utc, end_utc, attendees, location, body, is_all_day) -> OperationResult:
    try:
        app = namespace.Application
        item = app.CreateItem(1)  # olAppointmentItem
        item.Subject = subject or ""
        item.Start = start_utc
        item.End = end_utc
        item.Location = location or ""
        item.Body = body or ""
        item.AllDayEvent = bool(is_all_day)
        if attendees:
            item.MeetingStatus = 1  # olMeeting
            for a in attendees:
                recip = item.Recipients.Add(a)
                recip.Type = 1  # olRequired
            item.Recipients.ResolveAll()
            item.Send()
            return OperationResult(ok=True, message="meeting sent",
                                   affected_id=getattr(item, "EntryID", None))
        item.Save()
        return OperationResult(ok=True, message="event saved",
                               affected_id=getattr(item, "EntryID", None))
    except Exception as exc:
        return OperationResult(ok=False, message=f"create event failed: {exc}")


def _sync_update_calendar_event(namespace: Any, event_id, subject, start_utc, end_utc, location, body) -> OperationResult:
    try:
        item = namespace.GetItemFromID(event_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No event with id {event_id!r}.") from exc
    try:
        if subject is not None:
            item.Subject = subject
        if start_utc is not None:
            item.Start = start_utc
        if end_utc is not None:
            item.End = end_utc
        if location is not None:
            item.Location = location
        if body is not None:
            item.Body = body
        item.Save()
        return OperationResult(ok=True, message="updated", affected_id=event_id)
    except Exception as exc:
        return OperationResult(ok=False, message=f"update failed: {exc}")


def _sync_delete_calendar_event(namespace: Any, event_id: str) -> OperationResult:
    try:
        item = namespace.GetItemFromID(event_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No event with id {event_id!r}.") from exc
    try:
        item.Delete()
        return OperationResult(ok=True, message="deleted")
    except Exception as exc:
        return OperationResult(ok=False, message=f"delete failed: {exc}")


def _sync_respond_to_event(namespace: Any, event_id: str, response: str, send_response: bool) -> OperationResult:
    try:
        item = namespace.GetItemFromID(event_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No event with id {event_id!r}.") from exc
    try:
        # OlMeetingResponse: olMeetingAccepted=3, olMeetingTentative=2, olMeetingDeclined=4
        code = {"accept": 3, "tentative": 2, "decline": 4}.get(response.lower())
        if code is None:
            return OperationResult(ok=False, message=f"unknown response {response!r}")
        resp = item.Respond(code, True, send_response)
        if send_response and resp is not None:
            try:
                resp.Send()
            except Exception:
                pass
        return OperationResult(ok=True, message=f"responded {response}")
    except Exception as exc:
        return OperationResult(ok=False, message=f"respond failed: {exc}")


def _extract_task(item: Any) -> TaskInfo:
    return TaskInfo(
        task_id=item.EntryID,
        subject=getattr(item, "Subject", "") or "",
        due_date_utc=_pywintime_to_utc(getattr(item, "DueDate", None)),
        is_complete=bool(getattr(item, "Complete", False)),
        importance=detect_importance(getattr(item, "Importance", 1)),
        body_plain=getattr(item, "Body", "") or "",
    )


def _sync_list_tasks(namespace: Any, limit: int, offset: int) -> list[TaskInfo]:
    folder = namespace.GetDefaultFolder(OL_TASK_FOLDER_INDEX)
    out: list[TaskInfo] = []
    idx = 0
    for item in folder.Items:
        if idx < offset:
            idx += 1
            continue
        if len(out) >= limit:
            break
        out.append(_extract_task(item))
        idx += 1
    return out


def _sync_search_tasks(namespace: Any, query: str, limit: int) -> list[TaskInfo]:
    q = query.lower()
    folder = namespace.GetDefaultFolder(OL_TASK_FOLDER_INDEX)
    out: list[TaskInfo] = []
    for item in folder.Items:
        if q in (getattr(item, "Subject", "") or "").lower() or q in (getattr(item, "Body", "") or "").lower():
            out.append(_extract_task(item))
            if len(out) >= limit:
                break
    return out


def _sync_get_task(namespace: Any, task_id: str) -> TaskInfo:
    try:
        item = namespace.GetItemFromID(task_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No task with id {task_id!r}.") from exc
    return _extract_task(item)


def _extract_note(item: Any) -> NoteInfo:
    return NoteInfo(
        note_id=item.EntryID,
        subject=getattr(item, "Subject", "") or "",
        body_plain=getattr(item, "Body", "") or "",
        last_modified_utc=_pywintime_to_utc(getattr(item, "LastModificationTime", None)),
    )


def _sync_list_notes(namespace: Any, limit: int, offset: int) -> list[NoteInfo]:
    folder = namespace.GetDefaultFolder(OL_NOTES_FOLDER_INDEX)
    out: list[NoteInfo] = []
    idx = 0
    for item in folder.Items:
        if idx < offset:
            idx += 1
            continue
        if len(out) >= limit:
            break
        out.append(_extract_note(item))
        idx += 1
    return out


def _sync_search_notes(namespace: Any, query: str, limit: int) -> list[NoteInfo]:
    q = query.lower()
    folder = namespace.GetDefaultFolder(OL_NOTES_FOLDER_INDEX)
    out: list[NoteInfo] = []
    for item in folder.Items:
        if q in (getattr(item, "Subject", "") or "").lower() or q in (getattr(item, "Body", "") or "").lower():
            out.append(_extract_note(item))
            if len(out) >= limit:
                break
    return out


def _sync_get_note(namespace: Any, note_id: str) -> NoteInfo:
    try:
        item = namespace.GetItemFromID(note_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No note with id {note_id!r}.") from exc
    return _extract_note(item)


def _sync_list_accounts(namespace: Any) -> list[AccountInfo]:
    out: list[AccountInfo] = []
    try:
        for acct in namespace.Application.Session.Accounts:
            account_type_map = {0: "exchange", 1: "imap", 2: "pop", 3: "exchange", 5: "other"}
            try:
                kind = int(getattr(acct, "AccountType", 5))
            except Exception:
                kind = 5
            out.append(
                AccountInfo(
                    account_id=str(getattr(acct, "DisplayName", "") or ""),
                    display_name=getattr(acct, "DisplayName", "") or "",
                    smtp_address=getattr(acct, "SmtpAddress", "") or "",
                    account_type=account_type_map.get(kind, "other"),
                )
            )
    except Exception:
        pass
    return out


def _sync_get_unread_count(namespace: Any, folder: str | None) -> int:
    if folder:
        f = _resolve_folder(namespace, folder)
        try:
            return int(f.UnReadItemCount)
        except Exception:
            return 0
    # Total across all folders.
    total = 0

    def walk(folder: Any) -> None:
        nonlocal total
        try:
            total += int(folder.UnReadItemCount)
        except Exception:
            pass
        for sub in folder.Folders:
            walk(sub)

    root = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
    for f in root.Folders:
        walk(f)
    return total


# ---------------------------------------------------------------------------
# Phase 4 (Exchange-specific) sync helpers
# ---------------------------------------------------------------------------


def _sync_get_out_of_office(namespace: Any) -> OutOfOfficeStatus:
    try:
        ms = namespace.Application.Session.Stores.DefaultStore
        rules = namespace.Application.Session.OOFTemplates  # not standard; fallback below
    except Exception:
        rules = None
    # Outlook COM exposes Account.AutomaticReplies (Outlook 2016+).
    try:
        account = namespace.Application.Session.Accounts.Item(1)
        ar = account.AutomaticReplies
        return OutOfOfficeStatus(
            enabled=bool(ar.AutomaticRepliesEnabled),
            internal_message=str(ar.InternalReply or ""),
            external_message=str(ar.ExternalReply or ""),
            start_utc=_pywintime_to_utc(getattr(ar, "StartTime", None)),
            end_utc=_pywintime_to_utc(getattr(ar, "EndTime", None)),
            external_audience=("all" if getattr(ar, "ExternalAudience", 0) == 2 else
                               "contacts_only" if getattr(ar, "ExternalAudience", 0) == 1 else "none"),
        )
    except Exception as exc:
        return OutOfOfficeStatus(
            enabled=False,
            internal_message=f"unavailable: {exc}",
            external_message="",
        )


def _sync_set_out_of_office(namespace: Any, status: OutOfOfficeStatus) -> OperationResult:
    try:
        account = namespace.Application.Session.Accounts.Item(1)
        ar = account.AutomaticReplies
        ar.AutomaticRepliesEnabled = bool(status.enabled)
        ar.InternalReply = status.internal_message or ""
        ar.ExternalReply = status.external_message or ""
        if status.start_utc is not None:
            ar.StartTime = status.start_utc
        if status.end_utc is not None:
            ar.EndTime = status.end_utc
        audience_map = {"none": 0, "contacts_only": 1, "all": 2}
        ar.ExternalAudience = audience_map.get(status.external_audience or "all", 2)
        return OperationResult(ok=True, message="OOO updated")
    except Exception as exc:
        return OperationResult(ok=False, message=f"OOO set failed: {exc}")


def _sync_get_signature(namespace: Any, account_id: str | None) -> SignatureInfo:
    # Outlook stores signatures as files under %APPDATA%\Microsoft\Signatures
    # Reading them via COM is not supported; we read the filesystem directly.
    import os
    sig_dir = os.path.expandvars(r"%APPDATA%\Microsoft\Signatures")
    name = ""
    plain = ""
    html = ""
    try:
        if os.path.isdir(sig_dir):
            entries = sorted(
                [e for e in os.listdir(sig_dir) if e.lower().endswith(".htm") or e.lower().endswith(".html")]
            )
            if entries:
                name = os.path.splitext(entries[0])[0]
                with open(os.path.join(sig_dir, entries[0]), "r", encoding="utf-8", errors="replace") as fh:
                    html = fh.read()
                txt_path = os.path.join(sig_dir, name + ".txt")
                if os.path.exists(txt_path):
                    with open(txt_path, "r", encoding="utf-8", errors="replace") as fh:
                        plain = fh.read()
    except Exception:
        pass
    return SignatureInfo(name=name, body_html=html, body_plain=plain)


def _sync_set_signature(namespace: Any, account_id: str | None, body_html: str, body_plain: str) -> OperationResult:
    import os
    sig_dir = os.path.expandvars(r"%APPDATA%\Microsoft\Signatures")
    try:
        os.makedirs(sig_dir, exist_ok=True)
        sig_name = "default"
        with open(os.path.join(sig_dir, sig_name + ".htm"), "w", encoding="utf-8") as fh:
            fh.write(body_html or "")
        with open(os.path.join(sig_dir, sig_name + ".txt"), "w", encoding="utf-8") as fh:
            fh.write(body_plain or "")
        return OperationResult(ok=True, message=f"signature written to {sig_dir}")
    except Exception as exc:
        return OperationResult(ok=False, message=f"signature write failed: {exc}")


def _sync_list_rules(namespace: Any) -> list[RuleInfo]:
    try:
        store = namespace.Application.Session.DefaultStore
        rules = store.GetRules()
    except Exception:
        return []
    out: list[RuleInfo] = []
    for r in rules:
        try:
            out.append(
                RuleInfo(
                    rule_id=str(getattr(r, "ID", "") or getattr(r, "Name", "")),
                    name=str(getattr(r, "Name", "") or ""),
                    enabled=bool(getattr(r, "Enabled", False)),
                    description=str(getattr(r, "Description", "") or ""),
                )
            )
        except Exception:
            continue
    return out


def _sync_toggle_rule(namespace: Any, rule_id: str, enabled: bool) -> OperationResult:
    try:
        store = namespace.Application.Session.DefaultStore
        rules = store.GetRules()
        for r in rules:
            if str(getattr(r, "ID", "") or getattr(r, "Name", "")) == rule_id or str(getattr(r, "Name", "")) == rule_id:
                r.Enabled = bool(enabled)
                rules.Save()
                return OperationResult(ok=True, message=f"rule {'enabled' if enabled else 'disabled'}")
        return OperationResult(ok=False, message=f"rule {rule_id!r} not found")
    except Exception as exc:
        return OperationResult(ok=False, message=f"rule toggle failed: {exc}")


def _sync_calendar_freebusy(namespace: Any, smtps, start_utc, end_utc, slot_minutes) -> list[FreeBusyResponse]:
    out: list[FreeBusyResponse] = []
    duration_minutes = int((end_utc - start_utc).total_seconds() / 60)
    for smtp in smtps:
        try:
            recip = namespace.CreateRecipient(smtp)
            recip.Resolve()
            # GetFreeBusy returns a string like "00021..." where each char is a status code.
            fb = recip.AddressEntry.GetFreeBusy(start_utc, slot_minutes, True)
            slots: list[FreeBusySlot] = []
            status_map = {"0": "free", "1": "tentative", "2": "busy", "3": "oof", "4": "unknown"}
            from datetime import timedelta
            for i, ch in enumerate(fb[: max(1, duration_minutes // slot_minutes)]):
                slot_start = start_utc + timedelta(minutes=i * slot_minutes)
                slots.append(
                    FreeBusySlot(
                        start_utc=slot_start,
                        end_utc=slot_start + timedelta(minutes=slot_minutes),
                        status=status_map.get(ch, "unknown"),
                    )
                )
            out.append(FreeBusyResponse(smtp=smtp, slots=slots))
        except Exception as exc:
            out.append(FreeBusyResponse(smtp=smtp, slots=[]))
    return out


def _sync_meeting_room_finder(namespace: Any, start_utc, end_utc, capacity, location_hint):
    # Exchange's "Find Room" requires the Resource Address List, which COM
    # exposes via AddressLists with kind == 5 (olAddressListResource).
    rooms: list[Contact] = []
    try:
        for al in namespace.AddressLists:
            try:
                kind = int(getattr(al, "AddressListType", 0))
            except Exception:
                kind = 0
            if kind != 5:
                continue
            for entry in al.AddressEntries:
                try:
                    smtp = entry.GetExchangeUser().PrimarySmtpAddress if entry.GetExchangeUser() else (entry.Address or "")
                except Exception:
                    smtp = entry.Address or ""
                rooms.append(
                    Contact(
                        contact_id=getattr(entry, "ID", "") or smtp,
                        display_name=getattr(entry, "Name", "") or "",
                        primary_smtp=smtp,
                    )
                )
    except Exception:
        pass
    if location_hint:
        hint = location_hint.lower()
        rooms = [r for r in rooms if hint in (r.display_name or "").lower()]
    # Filter by free/busy in the requested window.
    available: list[Contact] = []
    smtps = [r.primary_smtp for r in rooms if r.primary_smtp]
    if smtps:
        slot_minutes = max(15, int((end_utc - start_utc).total_seconds() / 60))
        fb = _sync_calendar_freebusy(namespace, smtps, start_utc, end_utc, slot_minutes)
        free_set = {f.smtp for f in fb if all(s.status == "free" for s in f.slots)}
        available = [r for r in rooms if r.primary_smtp in free_set]
    return available[: max(1, capacity)]  # capacity limits results, not enforces room size (room size = no COM API)


def _sync_gal_search(namespace: Any, query: str, limit: int) -> list[Contact]:
    q = query.lower()
    out: list[Contact] = []
    try:
        gal = namespace.GetGlobalAddressList()
    except Exception:
        return []
    for entry in gal.AddressEntries:
        if len(out) >= limit:
            break
        try:
            name = getattr(entry, "Name", "") or ""
            smtp = ""
            try:
                user = entry.GetExchangeUser()
                if user:
                    smtp = user.PrimarySmtpAddress or ""
            except Exception:
                pass
            if not smtp:
                smtp = getattr(entry, "Address", "") or ""
            haystack = (name + " " + smtp).lower()
            if q not in haystack:
                continue
            company = ""
            title = ""
            try:
                user = entry.GetExchangeUser()
                if user:
                    company = getattr(user, "CompanyName", "") or ""
                    title = getattr(user, "JobTitle", "") or ""
            except Exception:
                pass
            out.append(
                Contact(
                    contact_id=getattr(entry, "ID", "") or smtp,
                    display_name=name,
                    primary_smtp=smtp,
                    company=company,
                    job_title=title,
                )
            )
        except Exception:
            continue
    return out


def _sync_list_delegated_mailboxes(namespace: Any) -> list[MailboxInfo]:
    out: list[MailboxInfo] = []
    try:
        for store in namespace.Stores:
            try:
                if not store.IsDataFileStore:
                    name = getattr(store, "DisplayName", "") or ""
                    smtp = ""
                    try:
                        # Best-effort SMTP via the root folder owner.
                        root = store.GetRootFolder()
                        smtp = getattr(root, "FolderPath", "") or ""
                    except Exception:
                        pass
                    out.append(
                        MailboxInfo(
                            display_name=name,
                            smtp_address=smtp,
                            domain=domain_of(smtp),
                        )
                    )
            except Exception:
                continue
    except Exception:
        pass
    return out


def _sync_list_public_folders(namespace: Any) -> list[FolderInfo]:
    out: list[FolderInfo] = []
    try:
        for store in namespace.Stores:
            name = getattr(store, "DisplayName", "") or ""
            if "public folder" in name.lower():
                root = store.GetRootFolder()
                for f in root.Folders:
                    try:
                        out.append(
                            FolderInfo(
                                name=getattr(f, "Name", "") or "",
                                full_path=f"public/{getattr(f, 'Name', '')}",
                                item_count=int(getattr(f.Items, "Count", 0)),
                                unread_count=int(getattr(f, "UnReadItemCount", 0)),
                                parent_path="public",
                            )
                        )
                    except Exception:
                        continue
    except Exception:
        pass
    return out


def _sync_get_mailbox_quota(namespace: Any) -> MailboxQuota:
    try:
        store = namespace.Application.Session.DefaultStore
        # PR_MESSAGE_SIZE_EXTENDED gives total store size.
        PR_MESSAGE_SIZE_EXTENDED = "http://schemas.microsoft.com/mapi/proptag/0x0E08000B"
        used = int(safe_get(store, PR_MESSAGE_SIZE_EXTENDED) or 0)
        # PR_PROHIBIT_SEND_QUOTA / PR_STORAGE_QUOTA are not always exposed.
        PR_PROHIBIT_SEND_QUOTA = "http://schemas.microsoft.com/mapi/proptag/0x666E0003"
        PR_PROHIBIT_RECEIVE_QUOTA = "http://schemas.microsoft.com/mapi/proptag/0x666A0003"
        PR_STORAGE_QUOTA = "http://schemas.microsoft.com/mapi/proptag/0x66700003"
        prohibit_send = safe_get(store, PR_PROHIBIT_SEND_QUOTA)
        warning = safe_get(store, PR_STORAGE_QUOTA)
        prohibit_recv = safe_get(store, PR_PROHIBIT_RECEIVE_QUOTA)
        total = int(prohibit_recv or 0) * 1024 if prohibit_recv else 0
        item_count = 0
        try:
            root = store.GetRootFolder()
            for f in root.Folders:
                try:
                    item_count += int(f.Items.Count)
                except Exception:
                    continue
        except Exception:
            pass
        return MailboxQuota(
            total_bytes=total,
            used_bytes=used,
            warning_bytes=int(warning) * 1024 if warning else None,
            prohibit_send_bytes=int(prohibit_send) * 1024 if prohibit_send else None,
            item_count=item_count,
        )
    except Exception:
        return MailboxQuota(total_bytes=0, used_bytes=0, warning_bytes=None, prohibit_send_bytes=None, item_count=0)


__all__ = [
    "WindowsComBackend",
    "normalize_subject",  # re-exported so tools/thread_tools can use from backend
    "ThreadMetadata",
]
