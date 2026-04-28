"""Parameterised AppleScript templates.

Output format: each record is delimited by a sentinel that does not appear
naturally in Outlook data (`§§REC§§`). Within a record, fields use
`§§FLD§§key=value`. This lets us parse without worrying about commas,
quotes, or newlines inside values.

Each template function returns a single AppleScript source string ready to
pass to `run_osascript`.
"""

from __future__ import annotations

from datetime import datetime

from .escape import applescript_epoch_date, as_str

# Sentinels — pick strings unlikely to appear in email text.
REC = "§§REC§§"
FLD = "§§FLD§§"


def _preamble() -> str:
    """Common preamble: sentinels + locale-independent epoch helpers.

    `nowDate`/`nowEpoch` let us convert any AppleScript `date` to Unix
    seconds via `((d - nowDate) + nowEpoch) as integer`, avoiding the
    locale-dependent `date "<English string>"` literal entirely.
    """
    return (
        f"set recSep to \"{REC}\"\n"
        f"set fldSep to \"{FLD}\"\n"
        "set nowDate to (current date)\n"
        "set nowEpoch to (do shell script \"date +%s\") as integer\n"
    )


def _epoch_expr(varname: str) -> str:
    """AppleScript expression: convert date variable to Unix epoch seconds."""
    return f"(({varname} - nowDate) + nowEpoch) as integer"


def mailbox_info() -> str:
    return (
        'tell application "Microsoft Outlook"\n'
        "  try\n"
        "    set acct to default exchange account\n"
        "  on error\n"
        "    try\n"
        "      set acct to first exchange account\n"
        "    on error\n"
        "      set acct to first pop account\n"
        "    end try\n"
        "  end try\n"
        "  set accName to full name of acct\n"
        "  set accEmail to email address of acct\n"
        f"  return \"{FLD}name=\" & accName & \"{FLD}smtp=\" & accEmail\n"
        "end tell\n"
    )


def list_folders(include_subfolders: bool) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set allOut to ""\n'
        "  set rootFolders to mail folders\n"
        "  repeat with f in rootFolders\n"
        "    set fName to name of f\n"
        "    set fCount to count of messages of f\n"
        "    set fUnread to unread count of f\n"
        "    set allOut to allOut & recSep & fldSep & \"name=\" & fName & fldSep & \"path=\" & fName & fldSep & \"count=\" & fCount & fldSep & \"unread=\" & fUnread\n"
        + ("    repeat with sub in mail folders of f\n"
           "      set sName to name of sub\n"
           "      set sCount to count of messages of sub\n"
           "      set sUnread to unread count of sub\n"
           "      set allOut to allOut & recSep & fldSep & \"name=\" & sName & fldSep & \"path=\" & fName & \"/\" & sName & fldSep & \"count=\" & sCount & fldSep & \"unread=\" & sUnread\n"
           "    end repeat\n"
           if include_subfolders else "")
        + "  end repeat\n"
        "  return allOut\n"
        "end tell\n"
    )


def _message_record_expr(prefix: str = "m") -> str:
    """Build the AppleScript expression that emits one message record.

    Operates on variable `m` (an incoming/outgoing message). Fields emitted:
    id, subject, sender, senderAddress, received, sent, isRead, hasAttach,
    convId, folderName, preview.
    """
    return (
        "set msgId to id of " + prefix + "\n"
        "set msgSubject to subject of " + prefix + "\n"
        "try\n"
        "  set senderRec to sender of " + prefix + "\n"
        "  set senderName to name of senderRec\n"
        "  set senderAddr to address of senderRec\n"
        "on error\n"
        "  set senderName to \"\"\n"
        "  set senderAddr to \"\"\n"
        "end try\n"
        "set rTime to time received of " + prefix + "\n"
        "set sTime to time sent of " + prefix + "\n"
        "set isRead to is read of " + prefix + "\n"
        "set hasAttach to has attachment of " + prefix + "\n"
        "try\n"
        "  set convObj to conversation of " + prefix + "\n"
        "  set convId to id of convObj\n"
        "on error\n"
        "  set convId to \"\"\n"
        "end try\n"
        "try\n"
        "  set folderObj to mail folder of " + prefix + "\n"
        "  set folderName to name of folderObj\n"
        "on error\n"
        "  set folderName to \"\"\n"
        "end try\n"
        "try\n"
        "  set previewText to plain text content of " + prefix + "\n"
        "on error\n"
        "  set previewText to \"\"\n"
        "end try\n"
        "if length of previewText > 200 then set previewText to text 1 thru 200 of previewText\n"
        "set out to out & recSep & fldSep & \"id=\" & msgId & fldSep & \"subject=\" & msgSubject"
        " & fldSep & \"senderName=\" & senderName & fldSep & \"senderAddr=\" & senderAddr"
        " & fldSep & \"received=\" & ((rTime - nowDate) + nowEpoch) as integer"
        " & fldSep & \"sent=\" & ((sTime - nowDate) + nowEpoch) as integer"
        " & fldSep & \"isRead=\" & (isRead as string)"
        " & fldSep & \"hasAttach=\" & (hasAttach as string)"
        " & fldSep & \"convId=\" & convId"
        " & fldSep & \"folder=\" & folderName"
        " & fldSep & \"preview=\" & previewText\n"
    )


def emails_in_time_range(
    start_utc: datetime, end_utc: datetime, folders: list[str], max_results: int
) -> str:
    start_expr = applescript_epoch_date(start_utc)
    end_expr = applescript_epoch_date(end_utc)
    folder_literals = ", ".join(as_str(f) for f in folders)
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        "  set startDate to " + start_expr + "\n"
        "  set endDate to " + end_expr + "\n"
        f"  set folderNames to {{{folder_literals}}}\n"
        f"  set maxRes to {max_results}\n"
        "  set total to 0\n"
        "  repeat with fname in folderNames\n"
        "    try\n"
        "      set f to mail folder (fname as string)\n"
        "    on error\n"
        "      try\n"
        "        set f to first mail folder whose name is (fname as string)\n"
        "      on error\n"
        "        set f to missing value\n"
        "      end try\n"
        "    end try\n"
        "    if f is not missing value then\n"
        "      set msgs to (messages of f whose time received is greater than or equal to startDate and time received is less than endDate)\n"
        "      repeat with m in msgs\n"
        "        if total >= maxRes then exit repeat\n"
        + _indent(_message_record_expr("m"), 8)
        + "        set total to total + 1\n"
        "      end repeat\n"
        "    end if\n"
        "    if total >= maxRes then exit repeat\n"
        "  end repeat\n"
        "  return out\n"
        "end tell\n"
    )


def conversation_thread(conversation_id: str, max_messages: int) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        f"  set convId to {as_str(conversation_id)}\n"
        f"  set maxRes to {max_messages}\n"
        "  set total to 0\n"
        "  try\n"
        "    set convObj to conversation id convId\n"
        "    set msgs to messages of convObj\n"
        "    repeat with m in msgs\n"
        "      if total >= maxRes then exit repeat\n"
        + _indent(_message_record_expr("m"), 6)
        + "      set total to total + 1\n"
        "    end repeat\n"
        "  on error errMsg\n"
        "    return \"ERR:\" & errMsg\n"
        "  end try\n"
        "  return out\n"
        "end tell\n"
    )


def email_metadata(message_id: str) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        f"  set m to incoming message id {as_str(message_id)}\n"
        "  set out to \"\"\n"
        + _indent(_message_record_expr("m"), 2)
        + "  try\n"
        "    set rawSource to source of m\n"
        "    set out to out & fldSep & \"source=\" & rawSource\n"
        "  on error\n"
        "    set out to out & fldSep & \"source=__UNAVAILABLE__\"\n"
        "  end try\n"
        "  try\n"
        "    set htmlBody to content of m\n"
        "  on error\n"
        "    set htmlBody to \"\"\n"
        "  end try\n"
        "  set out to out & fldSep & \"html=\" & htmlBody\n"
        "  try\n"
        "    set attList to attachments of m\n"
        "    repeat with a in attList\n"
        "      set aName to name of a\n"
        "      set aSize to file size of a\n"
        "      try\n"
        "        set aCid to content identifier of a\n"
        "      on error\n"
        "        set aCid to \"\"\n"
        "      end try\n"
        "      set out to out & recSep & fldSep & \"attachName=\" & aName & fldSep & \"attachSize=\" & aSize & fldSep & \"attachCid=\" & aCid\n"
        "    end repeat\n"
        "  end try\n"
        "  return out\n"
        "end tell\n"
    )


def save_attachment(message_id: str, attachment_index: int, save_path: str) -> str:
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set m to incoming message id {as_str(message_id)}\n"
        f"  set a to attachment {attachment_index} of m\n"
        f"  save a in (POSIX file {as_str(save_path)})\n"
        "  set aName to name of a\n"
        "  set aSize to file size of a\n"
        f"  return \"{FLD}name=\" & aName & \"{FLD}size=\" & aSize\n"
        "end tell\n"
    )


def search_emails(query: str, folders: list[str], max_results: int) -> str:
    folder_literals = ", ".join(as_str(f) for f in folders)
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        f"  set q to {as_str(query)}\n"
        f"  set folderNames to {{{folder_literals}}}\n"
        f"  set maxRes to {max_results}\n"
        "  set total to 0\n"
        "  repeat with fname in folderNames\n"
        "    try\n"
        "      set f to mail folder (fname as string)\n"
        "    on error\n"
        "      set f to missing value\n"
        "    end try\n"
        "    if f is not missing value then\n"
        "      set msgs to (messages of f whose (subject contains q or (plain text content contains q)))\n"
        "      repeat with m in msgs\n"
        "        if total >= maxRes then exit repeat\n"
        + _indent(_message_record_expr("m"), 8)
        + "        set total to total + 1\n"
        "      end repeat\n"
        "    end if\n"
        "    if total >= maxRes then exit repeat\n"
        "  end repeat\n"
        "  return out\n"
        "end tell\n"
    )


def list_calendar_events(
    start_utc: datetime, end_utc: datetime, max_results: int
) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        "  set startDate to " + applescript_epoch_date(start_utc) + "\n"
        "  set endDate to " + applescript_epoch_date(end_utc) + "\n"
        f"  set maxRes to {max_results}\n"
        "  set total to 0\n"
        "  set evts to (calendar events whose start time is greater than or equal to startDate and start time is less than endDate)\n"
        "  repeat with e in evts\n"
        "    if total >= maxRes then exit repeat\n"
        "    set eId to id of e\n"
        "    set eSubject to subject of e\n"
        "    try\n"
        "      set eOrganizer to address of organizer of e\n"
        "    on error\n"
        "      set eOrganizer to \"\"\n"
        "    end try\n"
        "    set eStart to start time of e\n"
        "    set eEnd to end time of e\n"
        "    set eLocation to location of e\n"
        "    set eAllDay to all day flag of e\n"
        "    set out to out & recSep & fldSep & \"id=\" & eId & fldSep & \"subject=\" & eSubject"
        " & fldSep & \"organizer=\" & eOrganizer"
        " & fldSep & \"start=\" & ((eStart - nowDate) + nowEpoch) as integer"
        " & fldSep & \"end=\" & ((eEnd - nowDate) + nowEpoch) as integer"
        " & fldSep & \"location=\" & eLocation"
        " & fldSep & \"allDay=\" & (eAllDay as string)\n"
        "    set total to total + 1\n"
        "  end repeat\n"
        "  return out\n"
        "end tell\n"
    )


def get_calendar_event(event_id: str) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        f"  set e to calendar event id {as_str(event_id)}\n"
        "  set eSubject to subject of e\n"
        "  try\n"
        "    set eOrganizer to address of organizer of e\n"
        "  on error\n"
        "    set eOrganizer to \"\"\n"
        "  end try\n"
        "  set eStart to start time of e\n"
        "  set eEnd to end time of e\n"
        "  set eLocation to location of e\n"
        "  set eAllDay to all day flag of e\n"
        "  set eBody to plain text content of e\n"
        "  set out to fldSep & \"subject=\" & eSubject"
        " & fldSep & \"organizer=\" & eOrganizer"
        " & fldSep & \"start=\" & ((eStart - nowDate) + nowEpoch) as integer"
        " & fldSep & \"end=\" & ((eEnd - nowDate) + nowEpoch) as integer"
        " & fldSep & \"location=\" & eLocation"
        " & fldSep & \"allDay=\" & (eAllDay as string)"
        " & fldSep & \"body=\" & eBody\n"
        "  try\n"
        "    set attendeeList to attendees of e\n"
        "    repeat with att in attendeeList\n"
        "      set aAddr to address of att\n"
        "      set out to out & recSep & fldSep & \"attendee=\" & aAddr\n"
        "    end repeat\n"
        "  end try\n"
        "  return out\n"
        "end tell\n"
    )


def list_contacts(limit: int, offset: int) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        "  set cts to contacts\n"
        f"  set lim to {limit}\n"
        f"  set off to {offset}\n"
        "  set total to 0\n"
        "  set idx to 0\n"
        "  repeat with c in cts\n"
        "    if idx >= off then\n"
        "      if total >= lim then exit repeat\n"
        "      set cId to id of c\n"
        "      set cName to display name of c\n"
        "      try\n"
        "        set cEmail to address of first email address of c\n"
        "      on error\n"
        "        set cEmail to \"\"\n"
        "      end try\n"
        "      try\n"
        "        set cCompany to company of c\n"
        "      on error\n"
        "        set cCompany to \"\"\n"
        "      end try\n"
        "      try\n"
        "        set cTitle to job title of c\n"
        "      on error\n"
        "        set cTitle to \"\"\n"
        "      end try\n"
        "      set out to out & recSep & fldSep & \"id=\" & cId & fldSep & \"name=\" & cName"
        " & fldSep & \"email=\" & cEmail & fldSep & \"company=\" & cCompany"
        " & fldSep & \"title=\" & cTitle\n"
        "      set total to total + 1\n"
        "    end if\n"
        "    set idx to idx + 1\n"
        "  end repeat\n"
        "  return out\n"
        "end tell\n"
    )


def search_contacts(query: str, limit: int) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        f"  set q to {as_str(query)}\n"
        f"  set lim to {limit}\n"
        "  set total to 0\n"
        "  set cts to (contacts whose (display name contains q) or (company contains q))\n"
        "  repeat with c in cts\n"
        "    if total >= lim then exit repeat\n"
        "    set cId to id of c\n"
        "    set cName to display name of c\n"
        "    try\n"
        "      set cEmail to address of first email address of c\n"
        "    on error\n"
        "      set cEmail to \"\"\n"
        "    end try\n"
        "    try\n"
        "      set cCompany to company of c\n"
        "    on error\n"
        "      set cCompany to \"\"\n"
        "    end try\n"
        "    try\n"
        "      set cTitle to job title of c\n"
        "    on error\n"
        "      set cTitle to \"\"\n"
        "    end try\n"
        "    set out to out & recSep & fldSep & \"id=\" & cId & fldSep & \"name=\" & cName"
        " & fldSep & \"email=\" & cEmail & fldSep & \"company=\" & cCompany"
        " & fldSep & \"title=\" & cTitle\n"
        "    set total to total + 1\n"
        "  end repeat\n"
        "  return out\n"
        "end tell\n"
    )


def get_contact(contact_id: str) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        f"  set c to contact id {as_str(contact_id)}\n"
        "  set cName to display name of c\n"
        "  try\n"
        "    set cCompany to company of c\n"
        "  on error\n"
        "    set cCompany to \"\"\n"
        "  end try\n"
        "  try\n"
        "    set cTitle to job title of c\n"
        "  on error\n"
        "    set cTitle to \"\"\n"
        "  end try\n"
        "  set out to fldSep & \"name=\" & cName & fldSep & \"company=\" & cCompany & fldSep & \"title=\" & cTitle\n"
        "  try\n"
        "    set emails to email addresses of c\n"
        "    repeat with e in emails\n"
        "      set out to out & recSep & fldSep & \"email=\" & (address of e)\n"
        "    end repeat\n"
        "  end try\n"
        "  try\n"
        "    set phones to phone numbers of c\n"
        "    repeat with p in phones\n"
        "      set out to out & recSep & fldSep & \"phone=\" & (number of p) & fldSep & \"phoneLabel=\" & (label of p)\n"
        "    end repeat\n"
        "  end try\n"
        "  return out\n"
        "end tell\n"
    )


def diagnostics(message_props: list[str], sample_message_id: str | None) -> str:
    """Probe Outlook version + each message property on a sample message."""
    sample_clause = (
        f"  set m to incoming message id {as_str(sample_message_id)}\n"
        if sample_message_id
        else (
            "  try\n"
            "    set f to first mail folder whose name is \"Inbox\"\n"
            "    set msgs to messages of f\n"
            "    if (count of msgs) > 0 then\n"
            "      set m to first item of msgs\n"
            "    else\n"
            "      set m to missing value\n"
            "    end if\n"
            "  on error\n"
            "    set m to missing value\n"
            "  end try\n"
        )
    )
    probes_lines = []
    for prop in message_props:
        # Each probe: try to read the property; record ok | error.
        ascii_key = prop.replace(" ", "_")
        probes_lines.append(
            f"  try\n"
            f"    if m is not missing value then set _v to {prop} of m\n"
            f"    set out to out & fldSep & \"{ascii_key}=ok\"\n"
            f"  on error errMsg\n"
            f"    set out to out & fldSep & \"{ascii_key}=error:\" & errMsg\n"
            f"  end try\n"
        )
    probes = "".join(probes_lines)
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        "  set vers to version\n"
        f"  set out to fldSep & \"version=\" & vers\n"
        + sample_clause
        + "  if m is missing value then\n"
        + "    set out to out & fldSep & \"sample=missing\"\n"
        + "  else\n"
        + "    set out to out & fldSep & \"sample=ok\"\n"
        + "  end if\n"
        + probes
        + "  return out\n"
        + "end tell\n"
    )


# ---------------------------------------------------------------------------
# Phase 3 (parity) templates
# ---------------------------------------------------------------------------


def send_email(to: list[str], subject: str, body: str, body_type: str,
               cc: list[str] | None, bcc: list[str] | None,
               attachments: list[str] | None, send: bool) -> str:
    is_html = (body_type or "plain").lower() == "html"
    body_prop = "content" if is_html else "plain text content"
    to_lit = ", ".join(as_str(a) for a in to)
    cc_lit = ", ".join(as_str(a) for a in (cc or []))
    bcc_lit = ", ".join(as_str(a) for a in (bcc or []))
    att_lit = ", ".join(as_str(a) for a in (attachments or []))
    action = "send newMsg" if send else "save newMsg"
    cc_block = (
        f"  set ccRecipients to {{{cc_lit}}}\n"
        "  repeat with addr in ccRecipients\n"
        '    make new cc recipient at newMsg with properties {email address:{address:addr}}\n'
        "  end repeat\n"
    ) if cc else ""
    bcc_block = (
        f"  set bccRecipients to {{{bcc_lit}}}\n"
        "  repeat with addr in bccRecipients\n"
        '    make new bcc recipient at newMsg with properties {email address:{address:addr}}\n'
        "  end repeat\n"
    ) if bcc else ""
    att_block = (
        f"  set attPaths to {{{att_lit}}}\n"
        "  repeat with p in attPaths\n"
        "    make new attachment at newMsg with properties {file:(POSIX file (p as string))}\n"
        "  end repeat\n"
    ) if attachments else ""
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set newMsg to make new outgoing message with properties {{subject:{as_str(subject)}, {body_prop}:{as_str(body)}}}\n"
        f"  set toRecipients to {{{to_lit}}}\n"
        "  repeat with addr in toRecipients\n"
        '    make new recipient at newMsg with properties {email address:{address:addr}}\n'
        "  end repeat\n"
        + cc_block
        + bcc_block
        + att_block
        + f"  {action}\n"
        + f"  return \"{FLD}id=\" & (id of newMsg as string)\n"
        + "end tell\n"
    )


def reply_email(message_id: str, body: str, body_type: str, reply_all: bool) -> str:
    is_html = (body_type or "plain").lower() == "html"
    body_prop = "content" if is_html else "plain text content"
    verb = "reply to all" if reply_all else "reply"
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set m to incoming message id {as_str(message_id)}\n"
        f"  set replyMsg to {verb} m\n"
        f"  set {body_prop} of replyMsg to {as_str(body)} & ({body_prop} of replyMsg)\n"
        "  send replyMsg\n"
        f"  return \"{FLD}id=\" & (id of replyMsg as string)\n"
        "end tell\n"
    )


def forward_email(message_id: str, to: list[str], body: str, body_type: str) -> str:
    is_html = (body_type or "plain").lower() == "html"
    body_prop = "content" if is_html else "plain text content"
    to_lit = ", ".join(as_str(a) for a in to)
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set m to incoming message id {as_str(message_id)}\n"
        "  set fwdMsg to redirect m\n"
        f"  set toRecipients to {{{to_lit}}}\n"
        "  repeat with addr in toRecipients\n"
        '    make new recipient at fwdMsg with properties {email address:{address:addr}}\n'
        "  end repeat\n"
        f"  set {body_prop} of fwdMsg to {as_str(body)} & ({body_prop} of fwdMsg)\n"
        "  send fwdMsg\n"
        f"  return \"{FLD}id=\" & (id of fwdMsg as string)\n"
        "end tell\n"
    )


def mark_email_read(message_id: str, is_read: bool) -> str:
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set m to incoming message id {as_str(message_id)}\n"
        f"  set is read of m to {str(is_read).lower()}\n"
        f"  return \"{FLD}ok=true\"\n"
        "end tell\n"
    )


def set_email_categories(message_id: str, categories: list[str]) -> str:
    cat_lit = ", ".join(as_str(c) for c in categories) if categories else ""
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set m to incoming message id {as_str(message_id)}\n"
        + (
            f"  set newCats to {{{cat_lit}}}\n"
            "  set categoryList to {}\n"
            "  repeat with cName in newCats\n"
            "    try\n"
            "      set c to category named cName\n"
            "    on error\n"
            "      set c to make new category with properties {name:cName}\n"
            "    end try\n"
            "    set end of categoryList to c\n"
            "  end repeat\n"
            "  set categories of m to categoryList\n"
            if categories else
            "  set categories of m to {}\n"
        )
        + f"  return \"{FLD}ok=true\"\n"
        + "end tell\n"
    )


def move_email(message_id: str, destination_folder: str) -> str:
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set m to incoming message id {as_str(message_id)}\n"
        f"  set dst to first mail folder whose name is {as_str(destination_folder)}\n"
        "  move m to dst\n"
        f"  return \"{FLD}ok=true\"\n"
        "end tell\n"
    )


def delete_email(message_id: str, permanent: bool) -> str:
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set m to incoming message id {as_str(message_id)}\n"
        + ("  delete m\n"
           if permanent else
           '  set dst to first mail folder whose name is "Deleted Items"\n'
           "  move m to dst\n")
        + f"  return \"{FLD}ok=true\"\n"
        + "end tell\n"
    )


def junk_email(message_id: str) -> str:
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set m to incoming message id {as_str(message_id)}\n"
        '  set dst to first mail folder whose name is "Junk Email"\n'
        "  move m to dst\n"
        f"  return \"{FLD}ok=true\"\n"
        "end tell\n"
    )


def create_folder(parent_path: str | None, name: str) -> str:
    if parent_path:
        target = f"  set parentF to mail folder {as_str(parent_path)}\n"
        suffix = '  set newF to make new mail folder at parentF with properties {name:' + as_str(name) + '}\n'
    else:
        suffix = '  set newF to make new mail folder with properties {name:' + as_str(name) + '}\n'
        target = ""
    return (
        'tell application "Microsoft Outlook"\n'
        + target
        + suffix
        + f"  return \"{FLD}id=\" & (id of newF as string)\n"
        + "end tell\n"
    )


def rename_folder(folder_path: str, new_name: str) -> str:
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set f to mail folder {as_str(folder_path)}\n"
        f"  set name of f to {as_str(new_name)}\n"
        f"  return \"{FLD}ok=true\"\n"
        "end tell\n"
    )


def delete_folder(folder_path: str) -> str:
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set f to mail folder {as_str(folder_path)}\n"
        "  delete f\n"
        f"  return \"{FLD}ok=true\"\n"
        "end tell\n"
    )


def empty_folder(folder_path: str) -> str:
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set f to mail folder {as_str(folder_path)}\n"
        "  set msgs to messages of f\n"
        "  repeat with m in msgs\n"
        "    delete m\n"
        "  end repeat\n"
        f"  return \"{FLD}ok=true\"\n"
        "end tell\n"
    )


def create_calendar_event(subject: str, start_utc: datetime, end_utc: datetime,
                          attendees: list[str], location: str, body: str,
                          is_all_day: bool) -> str:
    start_expr = applescript_epoch_date(start_utc)
    end_expr = applescript_epoch_date(end_utc)
    att_lit = ", ".join(as_str(a) for a in attendees)
    att_block = (
        f"  set attendeeList to {{{att_lit}}}\n"
        "  repeat with addr in attendeeList\n"
        '    make new required attendee at evt with properties {email address:{address:addr}}\n'
        "  end repeat\n"
    ) if attendees else ""
    return (
        'tell application "Microsoft Outlook"\n'
        "  set evt to make new calendar event with properties {"
        f"subject:{as_str(subject)}, "
        f"start time:{start_expr}, "
        f"end time:{end_expr}, "
        f"location:{as_str(location)}, "
        f"plain text content:{as_str(body)}, "
        f"all day flag:{str(is_all_day).lower()}"
        "}\n"
        + att_block
        + f"  return \"{FLD}id=\" & (id of evt as string)\n"
        + "end tell\n"
    )


def update_calendar_event(event_id: str, subject: str | None, start_utc: datetime | None,
                          end_utc: datetime | None, location: str | None, body: str | None) -> str:
    set_lines = []
    if subject is not None:
        set_lines.append(f"  set subject of evt to {as_str(subject)}\n")
    if start_utc is not None:
        set_lines.append(f"  set start time of evt to {applescript_epoch_date(start_utc)}\n")
    if end_utc is not None:
        set_lines.append(f"  set end time of evt to {applescript_epoch_date(end_utc)}\n")
    if location is not None:
        set_lines.append(f"  set location of evt to {as_str(location)}\n")
    if body is not None:
        set_lines.append(f"  set plain text content of evt to {as_str(body)}\n")
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set evt to calendar event id {as_str(event_id)}\n"
        + "".join(set_lines)
        + f"  return \"{FLD}ok=true\"\n"
        + "end tell\n"
    )


def delete_calendar_event(event_id: str) -> str:
    return (
        'tell application "Microsoft Outlook"\n'
        f"  set evt to calendar event id {as_str(event_id)}\n"
        "  delete evt\n"
        f"  return \"{FLD}ok=true\"\n"
        "end tell\n"
    )


def list_tasks(limit: int, offset: int) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        f"  set lim to {limit}\n"
        f"  set off to {offset}\n"
        "  set total to 0\n"
        "  set idx to 0\n"
        "  set tList to tasks\n"
        "  repeat with t in tList\n"
        "    if idx >= off then\n"
        "      if total >= lim then exit repeat\n"
        "      set tId to id of t\n"
        "      set tSubj to subject of t\n"
        "      try\n"
        "        set tDue to due date of t\n"
        "        set tDueEpoch to (((tDue - nowDate) + nowEpoch) as integer) as string\n"
        "      on error\n"
        "        set tDueEpoch to \"\"\n"
        "      end try\n"
        "      try\n"
        "        set tDone to completed of t\n"
        "      on error\n"
        "        set tDone to false\n"
        "      end try\n"
        "      try\n"
        "        set tBody to plain text content of t\n"
        "      on error\n"
        "        set tBody to \"\"\n"
        "      end try\n"
        "      set out to out & recSep & fldSep & \"id=\" & tId"
        " & fldSep & \"subject=\" & tSubj"
        " & fldSep & \"due=\" & tDueEpoch"
        " & fldSep & \"done=\" & (tDone as string)"
        " & fldSep & \"body=\" & tBody\n"
        "      set total to total + 1\n"
        "    end if\n"
        "    set idx to idx + 1\n"
        "  end repeat\n"
        "  return out\n"
        "end tell\n"
    )


def search_tasks(query: str, limit: int) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        f"  set q to {as_str(query)}\n"
        f"  set lim to {limit}\n"
        "  set total to 0\n"
        "  set tList to (tasks whose (subject contains q))\n"
        "  repeat with t in tList\n"
        "    if total >= lim then exit repeat\n"
        "    set tId to id of t\n"
        "    set tSubj to subject of t\n"
        "    try\n"
        "      set tDue to due date of t\n"
        "      set tDueEpoch to (((tDue - nowDate) + nowEpoch) as integer) as string\n"
        "    on error\n"
        "      set tDueEpoch to \"\"\n"
        "    end try\n"
        "    try\n"
        "      set tDone to completed of t\n"
        "    on error\n"
        "      set tDone to false\n"
        "    end try\n"
        "    set out to out & recSep & fldSep & \"id=\" & tId"
        " & fldSep & \"subject=\" & tSubj"
        " & fldSep & \"due=\" & tDueEpoch"
        " & fldSep & \"done=\" & (tDone as string)\n"
        "    set total to total + 1\n"
        "  end repeat\n"
        "  return out\n"
        "end tell\n"
    )


def get_task(task_id: str) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        f"  set t to task id {as_str(task_id)}\n"
        "  set tSubj to subject of t\n"
        "  try\n"
        "    set tDue to due date of t\n"
        "    set tDueEpoch to (((tDue - nowDate) + nowEpoch) as integer) as string\n"
        "  on error\n"
        "    set tDueEpoch to \"\"\n"
        "  end try\n"
        "  try\n"
        "    set tDone to completed of t\n"
        "  on error\n"
        "    set tDone to false\n"
        "  end try\n"
        "  try\n"
        "    set tBody to plain text content of t\n"
        "  on error\n"
        "    set tBody to \"\"\n"
        "  end try\n"
        "  set out to fldSep & \"subject=\" & tSubj"
        " & fldSep & \"due=\" & tDueEpoch"
        " & fldSep & \"done=\" & (tDone as string)"
        " & fldSep & \"body=\" & tBody\n"
        "  return out\n"
        "end tell\n"
    )


def list_notes(limit: int, offset: int) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        f"  set lim to {limit}\n"
        f"  set off to {offset}\n"
        "  set total to 0\n"
        "  set idx to 0\n"
        "  set nList to notes\n"
        "  repeat with n in nList\n"
        "    if idx >= off then\n"
        "      if total >= lim then exit repeat\n"
        "      set nId to id of n\n"
        "      try\n"
        "        set nSubj to subject of n\n"
        "      on error\n"
        "        set nSubj to \"\"\n"
        "      end try\n"
        "      try\n"
        "        set nBody to plain text content of n\n"
        "      on error\n"
        "        set nBody to \"\"\n"
        "      end try\n"
        "      try\n"
        "        set nMod to modification date of n\n"
        "        set nModEpoch to (((nMod - nowDate) + nowEpoch) as integer) as string\n"
        "      on error\n"
        "        set nModEpoch to \"\"\n"
        "      end try\n"
        "      set out to out & recSep & fldSep & \"id=\" & nId"
        " & fldSep & \"subject=\" & nSubj"
        " & fldSep & \"body=\" & nBody"
        " & fldSep & \"modified=\" & nModEpoch\n"
        "      set total to total + 1\n"
        "    end if\n"
        "    set idx to idx + 1\n"
        "  end repeat\n"
        "  return out\n"
        "end tell\n"
    )


def search_notes(query: str, limit: int) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        f"  set q to {as_str(query)}\n"
        f"  set lim to {limit}\n"
        "  set total to 0\n"
        "  set nList to (notes whose (subject contains q or plain text content contains q))\n"
        "  repeat with n in nList\n"
        "    if total >= lim then exit repeat\n"
        "    set nId to id of n\n"
        "    set nSubj to subject of n\n"
        "    set out to out & recSep & fldSep & \"id=\" & nId & fldSep & \"subject=\" & nSubj\n"
        "    set total to total + 1\n"
        "  end repeat\n"
        "  return out\n"
        "end tell\n"
    )


def get_note(note_id: str) -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        f"  set n to note id {as_str(note_id)}\n"
        "  set nSubj to subject of n\n"
        "  try\n"
        "    set nBody to plain text content of n\n"
        "  on error\n"
        "    set nBody to \"\"\n"
        "  end try\n"
        "  try\n"
        "    set nMod to modification date of n\n"
        "    set nModEpoch to (((nMod - nowDate) + nowEpoch) as integer) as string\n"
        "  on error\n"
        "    set nModEpoch to \"\"\n"
        "  end try\n"
        "  set out to fldSep & \"subject=\" & nSubj & fldSep & \"body=\" & nBody & fldSep & \"modified=\" & nModEpoch\n"
        "  return out\n"
        "end tell\n"
    )


def list_accounts() -> str:
    return (
        _preamble()
        + 'tell application "Microsoft Outlook"\n'
        '  set out to ""\n'
        "  try\n"
        "    set accs to exchange accounts\n"
        "    repeat with a in accs\n"
        "      set out to out & recSep & fldSep & \"id=\" & (id of a as string)"
        " & fldSep & \"name=\" & (full name of a)"
        " & fldSep & \"smtp=\" & (email address of a)"
        " & fldSep & \"type=exchange\"\n"
        "    end repeat\n"
        "  end try\n"
        "  try\n"
        "    set accs to imap accounts\n"
        "    repeat with a in accs\n"
        "      set out to out & recSep & fldSep & \"id=\" & (id of a as string)"
        " & fldSep & \"name=\" & (full name of a)"
        " & fldSep & \"smtp=\" & (email address of a)"
        " & fldSep & \"type=imap\"\n"
        "    end repeat\n"
        "  end try\n"
        "  try\n"
        "    set accs to pop accounts\n"
        "    repeat with a in accs\n"
        "      set out to out & recSep & fldSep & \"id=\" & (id of a as string)"
        " & fldSep & \"name=\" & (full name of a)"
        " & fldSep & \"smtp=\" & (email address of a)"
        " & fldSep & \"type=pop\"\n"
        "    end repeat\n"
        "  end try\n"
        "  return out\n"
        "end tell\n"
    )


def get_unread_count(folder: str | None) -> str:
    if folder:
        return (
            'tell application "Microsoft Outlook"\n'
            f"  set f to mail folder {as_str(folder)}\n"
            f"  return \"{FLD}count=\" & (unread count of f as string)\n"
            "end tell\n"
        )
    return (
        'tell application "Microsoft Outlook"\n'
        "  set total to 0\n"
        "  repeat with f in mail folders\n"
        "    try\n"
        "      set total to total + (unread count of f)\n"
        "    end try\n"
        "    repeat with sub in mail folders of f\n"
        "      try\n"
        "        set total to total + (unread count of sub)\n"
        "      end try\n"
        "    end repeat\n"
        "  end repeat\n"
        f"  return \"{FLD}count=\" & (total as string)\n"
        "end tell\n"
    )


def _indent(text: str, spaces: int) -> str:
    pad = " " * spaces
    return "\n".join(pad + line if line.strip() else line for line in text.splitlines()) + ("\n" if text.endswith("\n") else "")
