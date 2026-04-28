"""Subject normalization used by the thread-metadata composite tool.

Rules:
- Strip leading Re:/Fw:/Fwd:
- Strip bracketed tags like [External], [Internal], [Suspicious], etc.
- Preserve [Action Required] — the agent relies on it for thread priority.
- Collapse whitespace, lowercase.
"""

from __future__ import annotations

import re

from ..constants import BRACKETED_TAG_PATTERNS, REPLY_PREFIXES

_PLACEHOLDER = "\x00actionrequired\x00"
_ACTION_REQUIRED_RE = re.compile(r"\[action required\]", re.IGNORECASE)
_BRACKETED_RE = re.compile("|".join(BRACKETED_TAG_PATTERNS), re.IGNORECASE)
_REPLY_PREFIX_RE = re.compile(
    r"^(?:" + "|".join(re.escape(p) for p in REPLY_PREFIXES) + r")\s*",
    re.IGNORECASE,
)
_WHITESPACE_RE = re.compile(r"\s+")


def normalize_subject(subject: str) -> str:
    s = subject or ""
    # Preserve [Action Required] across the strip step.
    s = _ACTION_REQUIRED_RE.sub(_PLACEHOLDER, s)
    # Remove other bracketed tags.
    s = _BRACKETED_RE.sub("", s)
    # Strip repeating reply/forward prefixes (e.g. "Re: Re: Fwd: ...").
    while True:
        stripped = _REPLY_PREFIX_RE.sub("", s, count=1)
        if stripped == s:
            break
        s = stripped
    # Restore the preserved tag.
    s = s.replace(_PLACEHOLDER, "[action required]")
    # Collapse whitespace and lowercase.
    s = _WHITESPACE_RE.sub(" ", s).strip().lower()
    return s
