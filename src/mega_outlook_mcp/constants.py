"""Platform-neutral constants shared by both backends."""

from __future__ import annotations

MAC_UNAVAILABLE = "ERROR-MAC-Support-Unavailable"

DEFAULT_LOOKBACK_MINUTES = 65

IMPORTANCE_MAP = {0: "low", 1: "normal", 2: "high"}

# Windows OlDefaultFolders enum values (Outlook COM).
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT = 5
OL_FOLDER_DRAFTS = 16
OL_FOLDER_DELETED = 3
OL_FOLDER_JUNK = 23
OL_FOLDER_OUTBOX = 4
OL_FOLDER_CALENDAR = 9
OL_FOLDER_CONTACTS = 10

# OlAttachmentType enum values used to filter inline signature images.
OL_ATTACH_BY_VALUE = 1
OL_ATTACH_EMBEDDED_ITEM = 5
ATTACH_TYPES_REAL = {OL_ATTACH_BY_VALUE, OL_ATTACH_EMBEDDED_ITEM}

# Bracketed subject tags that should be stripped by normalize_subject.
# [Action Required] is preserved via a placeholder swap in subject_utils.
BRACKETED_TAG_PATTERNS = (
    r"\[external\]",
    r"\[internal\]",
    r"\[suspicious\]",
    r"\[phishing\]",
    r"\[encrypted\]",
    r"\[confidential\]",
)

# Reply/forward prefixes stripped from normalized subjects.
REPLY_PREFIXES = ("re:", "fw:", "fwd:")
