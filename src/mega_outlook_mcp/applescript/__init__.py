"""`osascript` runner used by the macOS backend."""

from __future__ import annotations

import asyncio
import shutil
import subprocess

from ..errors import AppleScriptError

DEFAULT_TIMEOUT_SECONDS = 30.0
LARGE_SCAN_TIMEOUT_SECONDS = 120.0


def _osascript_path() -> str:
    path = shutil.which("osascript")
    if path is None:
        raise AppleScriptError(
            "osascript executable not found. The macOS backend requires macOS; "
            "is this running on a Mac?"
        )
    return path


def run_osascript_sync(script: str, timeout: float = DEFAULT_TIMEOUT_SECONDS) -> str:
    """Run an AppleScript source string via `osascript -l AppleScript -e <script>`."""
    path = _osascript_path()
    try:
        completed = subprocess.run(
            [path, "-l", "AppleScript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise AppleScriptError(
            f"osascript timed out after {timeout}s. Is Outlook responsive?"
        ) from exc
    if completed.returncode != 0:
        raise AppleScriptError(
            f"osascript exited {completed.returncode}: {completed.stderr.strip() or completed.stdout.strip()}"
        )
    return completed.stdout


async def run_osascript(script: str, timeout: float = DEFAULT_TIMEOUT_SECONDS) -> str:
    return await asyncio.to_thread(run_osascript_sync, script, timeout)


__all__ = [
    "DEFAULT_TIMEOUT_SECONDS",
    "LARGE_SCAN_TIMEOUT_SECONDS",
    "run_osascript",
    "run_osascript_sync",
]
