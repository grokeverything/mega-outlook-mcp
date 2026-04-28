"""Phase 3 (parity) tool handlers.

A single module exposing all ~30 parity tools. Each handler is a thin
wrapper around the matching Backend method.
"""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime
from typing import Any

from ..backends.base import Backend, NoteInfo, TaskInfo
from ..models.inputs import (
    ArchiveEmailInput,
    CreateCalendarEventInput,
    CreateDraftInput,
    CreateFolderInput,
    DeleteCalendarEventInput,
    DeleteEmailInput,
    DeleteFolderInput,
    EmptyFolderInput,
    ForwardEmailInput,
    GetNoteInput,
    GetTaskInput,
    GetUnreadCountInput,
    JunkEmailInput,
    ListNotesInput,
    ListTasksInput,
    MarkEmailReadInput,
    MoveEmailInput,
    MoveFolderInput,
    RenameFolderInput,
    ReplyEmailInput,
    RespondToEventInput,
    SearchNotesInput,
    SearchTasksInput,
    SendEmailInput,
    SetEmailCategoriesInput,
    SetEmailFlagInput,
    UpdateCalendarEventInput,
)


def _task_to_json(t: TaskInfo) -> dict[str, Any]:
    raw = asdict(t)
    if isinstance(raw.get("due_date_utc"), datetime):
        raw["due_date_utc"] = raw["due_date_utc"].isoformat()
    return raw


def _note_to_json(n: NoteInfo) -> dict[str, Any]:
    raw = asdict(n)
    if isinstance(raw.get("last_modified_utc"), datetime):
        raw["last_modified_utc"] = raw["last_modified_utc"].isoformat()
    return raw


def register(mcp: Any, backend: Backend) -> None:
    # ---------------- Mail write ----------------
    @mcp.tool(name="outlook_send_email", description="Send an email immediately.")
    async def outlook_send_email(
        to: list[str], subject: str, body: str,
        body_type: str = "plain", cc: list[str] | None = None,
        bcc: list[str] | None = None, attachments: list[str] | None = None,
    ) -> dict[str, Any]:
        params = SendEmailInput(to=to, subject=subject, body=body, body_type=body_type,
                                cc=cc, bcc=bcc, attachments=attachments)
        return asdict(await backend.send_email(
            params.to, params.subject, params.body, params.body_type,
            params.cc, params.bcc, params.attachments))

    @mcp.tool(name="outlook_create_draft", description="Save an email as a draft without sending.")
    async def outlook_create_draft(
        to: list[str], subject: str, body: str,
        body_type: str = "plain", cc: list[str] | None = None,
        bcc: list[str] | None = None, attachments: list[str] | None = None,
    ) -> dict[str, Any]:
        params = CreateDraftInput(to=to, subject=subject, body=body, body_type=body_type,
                                  cc=cc, bcc=bcc, attachments=attachments)
        return asdict(await backend.create_draft(
            params.to, params.subject, params.body, params.body_type,
            params.cc, params.bcc, params.attachments))

    @mcp.tool(name="outlook_reply", description="Reply to an email; reply_all sends to the full recipient list.")
    async def outlook_reply(entry_id: str, body: str, body_type: str = "plain", reply_all: bool = False) -> dict[str, Any]:
        params = ReplyEmailInput(entry_id=entry_id, body=body, body_type=body_type, reply_all=reply_all)
        return asdict(await backend.reply_email(params.entry_id, params.body, params.body_type, params.reply_all))

    @mcp.tool(name="outlook_reply_all", description="Reply to all recipients of an email.")
    async def outlook_reply_all(entry_id: str, body: str, body_type: str = "plain") -> dict[str, Any]:
        params = ReplyEmailInput(entry_id=entry_id, body=body, body_type=body_type, reply_all=True)
        return asdict(await backend.reply_email(params.entry_id, params.body, params.body_type, True))

    @mcp.tool(name="outlook_forward", description="Forward an email to one or more recipients.")
    async def outlook_forward(entry_id: str, to: list[str], body: str, body_type: str = "plain") -> dict[str, Any]:
        params = ForwardEmailInput(entry_id=entry_id, to=to, body=body, body_type=body_type)
        return asdict(await backend.forward_email(params.entry_id, params.to, params.body, params.body_type))

    # ---------------- Mail organize ----------------
    @mcp.tool(name="outlook_mark_email_read", description="Mark an email as read or unread.")
    async def outlook_mark_email_read(entry_id: str, is_read: bool = True) -> dict[str, Any]:
        params = MarkEmailReadInput(entry_id=entry_id, is_read=is_read)
        return asdict(await backend.mark_email_read(params.entry_id, params.is_read))

    @mcp.tool(name="outlook_set_email_flag", description="Set or clear a follow-up flag on an email.")
    async def outlook_set_email_flag(entry_id: str, flag_status: str = "marked", due_date_utc: str | None = None) -> dict[str, Any]:
        due = datetime.fromisoformat(due_date_utc) if due_date_utc else None
        params = SetEmailFlagInput(entry_id=entry_id, flag_status=flag_status, due_date_utc=due)
        return asdict(await backend.set_email_flag(params.entry_id, params.flag_status, params.due_date_utc))

    @mcp.tool(name="outlook_set_email_categories", description="Replace the category list on an email. Pass [] to clear.")
    async def outlook_set_email_categories(entry_id: str, categories: list[str]) -> dict[str, Any]:
        params = SetEmailCategoriesInput(entry_id=entry_id, categories=categories)
        return asdict(await backend.set_email_categories(params.entry_id, params.categories))

    # ---------------- Mail destructive ----------------
    @mcp.tool(name="outlook_move_email", description="Move an email to another folder.")
    async def outlook_move_email(entry_id: str, destination_folder: str) -> dict[str, Any]:
        params = MoveEmailInput(entry_id=entry_id, destination_folder=destination_folder)
        return asdict(await backend.move_email(params.entry_id, params.destination_folder))

    @mcp.tool(name="outlook_archive_email", description="Move an email to the Archive folder.")
    async def outlook_archive_email(entry_id: str) -> dict[str, Any]:
        params = ArchiveEmailInput(entry_id=entry_id)
        return asdict(await backend.archive_email(params.entry_id))

    @mcp.tool(name="outlook_delete_email", description="Delete an email. permanent=true bypasses the Deleted Items folder.")
    async def outlook_delete_email(entry_id: str, permanent: bool = False) -> dict[str, Any]:
        params = DeleteEmailInput(entry_id=entry_id, permanent=permanent)
        return asdict(await backend.delete_email(params.entry_id, params.permanent))

    @mcp.tool(name="outlook_junk_email", description="Move an email to the Junk folder.")
    async def outlook_junk_email(entry_id: str) -> dict[str, Any]:
        params = JunkEmailInput(entry_id=entry_id)
        return asdict(await backend.junk_email(params.entry_id))

    # ---------------- Folder management ----------------
    @mcp.tool(name="outlook_create_folder", description="Create a new mail folder. Omit parent_path for a top-level folder.")
    async def outlook_create_folder(name: str, parent_path: str | None = None) -> dict[str, Any]:
        params = CreateFolderInput(parent_path=parent_path, name=name)
        return asdict(await backend.create_folder(params.parent_path, params.name))

    @mcp.tool(name="outlook_rename_folder", description="Rename a mail folder.")
    async def outlook_rename_folder(folder_path: str, new_name: str) -> dict[str, Any]:
        params = RenameFolderInput(folder_path=folder_path, new_name=new_name)
        return asdict(await backend.rename_folder(params.folder_path, params.new_name))

    @mcp.tool(name="outlook_move_folder", description="Move a folder to a new parent. Mac may return ERROR-MAC-Support-Unavailable.")
    async def outlook_move_folder(folder_path: str, new_parent_path: str) -> dict[str, Any]:
        params = MoveFolderInput(folder_path=folder_path, new_parent_path=new_parent_path)
        return asdict(await backend.move_folder(params.folder_path, params.new_parent_path))

    @mcp.tool(name="outlook_delete_folder", description="Delete a folder and all its contents.")
    async def outlook_delete_folder(folder_path: str) -> dict[str, Any]:
        params = DeleteFolderInput(folder_path=folder_path)
        return asdict(await backend.delete_folder(params.folder_path))

    @mcp.tool(name="outlook_empty_folder", description="Delete every item in a folder, leaving the folder itself.")
    async def outlook_empty_folder(folder_path: str) -> dict[str, Any]:
        params = EmptyFolderInput(folder_path=folder_path)
        return asdict(await backend.empty_folder(params.folder_path))

    # ---------------- Calendar write ----------------
    @mcp.tool(name="outlook_create_calendar_event", description="Create a calendar event. If attendees are supplied, sends a meeting invite.")
    async def outlook_create_calendar_event(
        subject: str, start_utc: str, end_utc: str,
        attendees: list[str] | None = None, location: str = "",
        body: str = "", is_all_day: bool = False,
    ) -> dict[str, Any]:
        params = CreateCalendarEventInput(
            subject=subject, start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc), attendees=attendees or [],
            location=location, body=body, is_all_day=is_all_day,
        )
        return asdict(await backend.create_calendar_event(
            params.subject, params.start_utc, params.end_utc, params.attendees,
            params.location, params.body, params.is_all_day))

    @mcp.tool(name="outlook_update_calendar_event", description="Update one or more fields on an existing event.")
    async def outlook_update_calendar_event(
        event_id: str, subject: str | None = None,
        start_utc: str | None = None, end_utc: str | None = None,
        location: str | None = None, body: str | None = None,
    ) -> dict[str, Any]:
        params = UpdateCalendarEventInput(
            event_id=event_id, subject=subject,
            start_utc=datetime.fromisoformat(start_utc) if start_utc else None,
            end_utc=datetime.fromisoformat(end_utc) if end_utc else None,
            location=location, body=body,
        )
        return asdict(await backend.update_calendar_event(
            params.event_id, params.subject, params.start_utc, params.end_utc,
            params.location, params.body))

    @mcp.tool(name="outlook_delete_calendar_event", description="Delete a calendar event.")
    async def outlook_delete_calendar_event(event_id: str) -> dict[str, Any]:
        params = DeleteCalendarEventInput(event_id=event_id)
        return asdict(await backend.delete_calendar_event(params.event_id))

    @mcp.tool(name="outlook_respond_to_event", description="Accept, tentatively accept, or decline a meeting. Mac returns ERROR-MAC-Support-Unavailable.")
    async def outlook_respond_to_event(event_id: str, response: str, send_response: bool = True) -> dict[str, Any]:
        params = RespondToEventInput(event_id=event_id, response=response, send_response=send_response)  # type: ignore[arg-type]
        return asdict(await backend.respond_to_event(params.event_id, params.response, params.send_response))

    # ---------------- Tasks ----------------
    @mcp.tool(name="outlook_list_tasks", description="Paginated list of tasks.")
    async def outlook_list_tasks(limit: int = 100, offset: int = 0) -> dict[str, Any]:
        params = ListTasksInput(limit=limit, offset=offset)
        tasks = await backend.list_tasks(params.limit, params.offset)
        return {"tasks": [_task_to_json(t) for t in tasks]}

    @mcp.tool(name="outlook_search_tasks", description="Keyword search across task subject and body.")
    async def outlook_search_tasks(query: str, limit: int = 50) -> dict[str, Any]:
        params = SearchTasksInput(query=query, limit=limit)
        tasks = await backend.search_tasks(params.query, params.limit)
        return {"tasks": [_task_to_json(t) for t in tasks]}

    @mcp.tool(name="outlook_get_task", description="Return a single task by id.")
    async def outlook_get_task(task_id: str) -> dict[str, Any]:
        params = GetTaskInput(task_id=task_id)
        return _task_to_json(await backend.get_task(params.task_id))

    # ---------------- Notes ----------------
    @mcp.tool(name="outlook_list_notes", description="Paginated list of sticky notes.")
    async def outlook_list_notes(limit: int = 100, offset: int = 0) -> dict[str, Any]:
        params = ListNotesInput(limit=limit, offset=offset)
        notes = await backend.list_notes(params.limit, params.offset)
        return {"notes": [_note_to_json(n) for n in notes]}

    @mcp.tool(name="outlook_search_notes", description="Keyword search across note subject and body.")
    async def outlook_search_notes(query: str, limit: int = 50) -> dict[str, Any]:
        params = SearchNotesInput(query=query, limit=limit)
        notes = await backend.search_notes(params.query, params.limit)
        return {"notes": [_note_to_json(n) for n in notes]}

    @mcp.tool(name="outlook_get_note", description="Return a single sticky note by id.")
    async def outlook_get_note(note_id: str) -> dict[str, Any]:
        params = GetNoteInput(note_id=note_id)
        return _note_to_json(await backend.get_note(params.note_id))

    # ---------------- Accounts ----------------
    @mcp.tool(name="outlook_list_accounts", description="List every account configured in this Outlook profile.")
    async def outlook_list_accounts() -> dict[str, Any]:
        accounts = await backend.list_accounts()
        return {"accounts": [asdict(a) for a in accounts]}

    @mcp.tool(name="outlook_get_unread_count", description="Get unread count for one folder, or the total across all folders.")
    async def outlook_get_unread_count(folder: str | None = None) -> dict[str, Any]:
        params = GetUnreadCountInput(folder=folder)
        count = await backend.get_unread_count(params.folder)
        return {"unread_count": count, "folder": params.folder or "all"}
