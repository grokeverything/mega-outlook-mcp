"""Typed errors surfaced by tool handlers.

Each error carries a concise, user-actionable message. Tool handlers catch
these and convert them to MCP error responses; they should not leak raw COM
or AppleScript tracebacks to the agent.
"""

from __future__ import annotations


class OutlookMcpError(Exception):
    """Base class for all errors raised by this server."""


class UnsupportedPlatformError(OutlookMcpError):
    """Raised when no backend is available for the current OS."""


class OutlookNotRunningError(OutlookMcpError):
    """Outlook is not running or the MAPI subsystem is unreachable."""


class FolderNotFoundError(OutlookMcpError):
    """The requested folder name or path could not be resolved."""


class EmailNotFoundError(OutlookMcpError):
    """The requested email id does not exist or is inaccessible."""


class ConversationNotFoundError(OutlookMcpError):
    """No items matched the requested conversation id."""


class AttachmentError(OutlookMcpError):
    """Attachment index out of range, save failed, or similar."""


class AppleScriptError(OutlookMcpError):
    """`osascript` exited non-zero or timed out."""


class ComBridgeError(OutlookMcpError):
    """The Windows COM STA bridge failed to dispatch or complete a job."""


class ValidationError(OutlookMcpError):
    """Input passed Pydantic but was rejected by backend-specific logic."""
