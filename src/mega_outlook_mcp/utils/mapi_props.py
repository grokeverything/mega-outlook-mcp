"""MAPI property tags and safe accessor for the Windows COM backend.

Import-guarded: on non-Windows the `safe_get` helper still works for tests
against fake objects, since it only uses duck typing.
"""

from __future__ import annotations

from typing import Any

# PR_SMTP_ADDRESS
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
# PR_INTERNET_MESSAGE_ID
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
# PR_IN_REPLY_TO_ID
PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"
# PR_INTERNET_REFERENCES
PR_INTERNET_REFERENCES = "http://schemas.microsoft.com/mapi/proptag/0x1039001F"
# PR_TRANSPORT_MESSAGE_HEADERS
PR_TRANSPORT_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
# PR_SENT_REPRESENTING_EMAIL_ADDRESS (string)
PR_SENT_REPRESENTING_ADDR = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"
# PR_SENT_REPRESENTING_SMTP_ADDRESS
PR_SENT_REPRESENTING_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F"
# PR_SENDER_SMTP_ADDRESS
PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"


def safe_get(obj: Any, prop: str) -> Any:
    """Safely read a MAPI property via `PropertyAccessor.GetProperty`.

    Returns None if the property is absent or the call raises.
    """
    try:
        accessor = obj.PropertyAccessor
    except Exception:
        return None
    try:
        return accessor.GetProperty(prop)
    except Exception:
        return None
