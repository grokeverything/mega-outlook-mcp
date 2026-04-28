"""Backend selection for the current platform.

Importing `select_backend` does not instantiate a backend; call it to get
the singleton for this process. The Windows backend starts its STA thread
on first instantiation; the Mac backend is stateless beyond a small cache.
"""

from __future__ import annotations

import sys

from ..errors import UnsupportedPlatformError
from .base import Backend


def select_backend() -> Backend:
    if sys.platform == "win32":
        from .windows_com import WindowsComBackend

        return WindowsComBackend()
    if sys.platform == "darwin":
        from .macos_applescript import MacOSAppleScriptBackend

        return MacOSAppleScriptBackend()
    raise UnsupportedPlatformError(
        f"mega-outlook-mcp supports Windows and macOS only; got sys.platform={sys.platform!r}"
    )


__all__ = ["Backend", "select_backend"]
