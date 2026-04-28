"""Single-threaded-apartment bridge for Outlook COM.

Outlook's COM interface requires all calls happen on a thread that has
called `CoInitialize`. This module runs a dedicated daemon thread, marshals
jobs onto it via a queue, and completes asyncio futures on the loop thread
via `call_soon_threadsafe`. Tool handlers await `bridge.execute(fn, ...)`
to run arbitrary synchronous COM work.

The Doc 1 snippet completed futures directly from the worker thread, which
is not thread-safe; that is fixed here.
"""

from __future__ import annotations

import asyncio
import queue
import threading
from typing import Any, Callable

from .errors import ComBridgeError, OutlookNotRunningError

_SENTINEL_SHUTDOWN = object()


class OutlookComBridge:
    """STA worker thread + asyncio-friendly job queue."""

    def __init__(self) -> None:
        self._queue: queue.Queue[Any] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._started = threading.Event()
        self._start_error: BaseException | None = None
        self._outlook_app = None
        self._namespace = None

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------
    def start(self) -> None:
        if self._thread is not None:
            return
        self._thread = threading.Thread(
            target=self._run, name="outlook-com-sta", daemon=True
        )
        self._thread.start()
        self._started.wait()
        if self._start_error is not None:
            raise self._start_error

    def close(self) -> None:
        if self._thread is None:
            return
        self._queue.put(_SENTINEL_SHUTDOWN)
        self._thread.join(timeout=5.0)
        self._thread = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    async def execute(self, fn: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
        """Run `fn(namespace, *args, **kwargs)` on the STA thread."""
        if self._thread is None:
            self.start()
        loop = asyncio.get_running_loop()
        future: asyncio.Future[Any] = loop.create_future()
        self._queue.put((fn, args, kwargs, loop, future))
        return await future

    # ------------------------------------------------------------------
    # Worker
    # ------------------------------------------------------------------
    def _run(self) -> None:
        try:
            import pythoncom  # type: ignore[import-not-found]
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            self._start_error = ComBridgeError(
                "pywin32 is not installed; install with `pip install pywin32` on Windows."
            )
            self._started.set()
            return

        try:
            pythoncom.CoInitialize()
        except Exception as exc:  # pragma: no cover
            self._start_error = ComBridgeError(f"CoInitialize failed: {exc}")
            self._started.set()
            return

        self._started.set()
        try:
            while True:
                job = self._queue.get()
                if job is _SENTINEL_SHUTDOWN:
                    break
                fn, args, kwargs, loop, future = job
                try:
                    namespace = self._ensure_connected(win32com, pythoncom)
                    result = fn(namespace, *args, **kwargs)
                    loop.call_soon_threadsafe(
                        _safe_set_result, future, result
                    )
                except BaseException as exc:  # noqa: BLE001
                    loop.call_soon_threadsafe(
                        _safe_set_exception, future, exc
                    )
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:  # pragma: no cover
                pass

    def _ensure_connected(self, win32com: Any, pythoncom: Any) -> Any:
        if self._namespace is not None:
            try:
                _ = self._namespace.CurrentUser
                return self._namespace
            except Exception:
                self._outlook_app = None
                self._namespace = None
        try:
            self._outlook_app = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook_app.GetNamespace("MAPI")
        except Exception as exc:  # noqa: BLE001
            raise OutlookNotRunningError(
                "Could not connect to Outlook. Make sure the Outlook desktop app is running "
                "and signed in, then retry."
            ) from exc
        return self._namespace


def _safe_set_result(future: asyncio.Future[Any], result: Any) -> None:
    if not future.done():
        future.set_result(result)


def _safe_set_exception(future: asyncio.Future[Any], exc: BaseException) -> None:
    if not future.done():
        future.set_exception(exc)
