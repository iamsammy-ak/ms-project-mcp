"""A single-threaded COM apartment for Microsoft Project.

Office applications are apartment-threaded COM servers. Calling their objects
from several threads at once produces intermittent ``RPC_E_WRONG_THREAD`` and
"Call was rejected by callee" failures, so every COM call in this project is
marshalled onto one dedicated STA thread and executed serially.

The MCP server is free to call :meth:`ComBridge.call` from any worker thread;
the bridge blocks until the COM thread finishes the job.
"""

from __future__ import annotations

import logging
import queue
import threading
import time
from concurrent.futures import Future
from concurrent.futures import TimeoutError as FutureTimeout
from typing import Callable, TypeVar

import pythoncom  # noqa: F401 - imported for CoInitialize in the worker thread

from . import config
from .errors import ComTimeout

log = logging.getLogger(__name__)

T = TypeVar("T")

# RPC_E_CALL_REJECTED and RPC_E_SERVERCALL_RETRYLATER: Project was busy and the
# call never ran, so retrying is safe.
_BUSY_HRESULTS = frozenset({-2147418111, -2147417846})


def is_busy_error(exc: BaseException) -> bool:
    args = getattr(exc, "args", ())
    return bool(args) and isinstance(args[0], int) and args[0] in _BUSY_HRESULTS


class ComBridge:
    """Serializes all COM work onto one STA thread."""

    def __init__(self) -> None:
        self._jobs: queue.Queue = queue.Queue()
        self._ready = threading.Event()
        self._thread = threading.Thread(
            target=self._loop, name="msproject-com", daemon=True
        )
        self._thread.start()
        if not self._ready.wait(10):
            raise RuntimeError("Timed out starting the Microsoft Project COM worker.")

    def _loop(self) -> None:
        pythoncom.CoInitialize()
        try:
            self._ready.set()
            while True:
                job = self._jobs.get()
                if job is None:
                    return
                fn, future = job
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    future.set_result(fn())
                except BaseException as exc:  # noqa: BLE001 - re-raised in the caller
                    future.set_exception(exc)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:  # pragma: no cover - shutdown best effort
                pass

    def call(
        self,
        fn: Callable[[], T],
        *,
        timeout: float | None = None,
        retries: int = 6,
    ) -> T:
        """Run ``fn`` on the COM thread and return its result."""

        future: Future = Future()

        def runner() -> T:
            return self._with_retry(fn, retries)

        self._jobs.put((runner, future))
        limit = timeout if timeout is not None else config.CALL_TIMEOUT
        try:
            return future.result(timeout=limit)
        except FutureTimeout:
            raise ComTimeout(
                "Microsoft Project did not respond within "
                f"{limit:g}s. It is probably showing a dialog or busy with a large "
                "operation; close any Project dialogs and retry."
            ) from None

    @staticmethod
    def _with_retry(fn: Callable[[], T], retries: int) -> T:
        delay = 0.1
        attempts = max(1, retries)
        last: BaseException | None = None
        for attempt in range(attempts):
            try:
                return fn()
            except Exception as exc:  # noqa: BLE001 - inspected, then re-raised
                if not is_busy_error(exc) or attempt == attempts - 1:
                    raise
                last = exc
                log.debug("Project busy (attempt %s), retrying in %.2fs", attempt + 1, delay)
                time.sleep(delay)
                delay = min(delay * 2, 2.0)
        assert last is not None
        raise last

    def close(self) -> None:
        self._jobs.put(None)
