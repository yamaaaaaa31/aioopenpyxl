"""Coarse-grained Excel jobs in a reusable, explicitly owned process pool."""

from __future__ import annotations

import functools
import math
import multiprocessing
import threading
from collections.abc import Callable
from concurrent.futures import ProcessPoolExecutor
from types import TracebackType
from typing import ParamSpec, TypeVar

from anyio import CapacityLimiter
from typing_extensions import Self

from ._executor import Runner, _acquire_token

__all__ = ["ProcessPool"]

P = ParamSpec("P")
R = TypeVar("R")


class ProcessPool:
    """Run complete synchronous Excel jobs in reusable worker processes.

    Use ``async with ProcessPool(max_workers=2) as pool:`` followed by
    ``await pool.run(func, *args, **kwargs)``. Works with asyncio and trio.
    Keep a pool for the application's lifetime rather than per request.

    ``func`` must be an importable, module-level synchronous callable, and its
    arguments and result must be picklable. Create/load, edit and save the raw
    openpyxl workbook *inside* that callable, passing paths, bytes or plain
    data across the process boundary. Workbook/Cell objects are not shared;
    this does not change the execution of the existing async wrappers.

    At most ``max_workers`` jobs are admitted at once (one by default). Jobs
    waiting for a slot can be cancelled. Once admitted, cancellation waits
    for the coordinating thread and any submitted job to finish, including
    native asyncio cancellation; it never kills a process midway through a
    save. Closing rejects new jobs and waits for admitted jobs before joining
    the workers, even if the closing task is cancelled.

    Workers use ``spawn`` on every platform, avoiding a fork of an already
    multithreaded server. A script's entry point must be guarded with
    ``if __name__ == "__main__":``. Context variables are not copied to workers;
    pass request metadata explicitly as arguments when needed.
    """

    def __init__(self, max_workers: int = 1) -> None:
        if not isinstance(max_workers, int):
            raise TypeError("max_workers must be an integer")
        if max_workers < 1:
            raise ValueError("max_workers must be >= 1")
        self._max_workers = max_workers
        self._limiter = CapacityLimiter(max_workers)
        self._executor: ProcessPoolExecutor | None = None
        self._closed = False
        self._active = 0
        self._idle = threading.Event()
        self._idle.set()
        self._shutdown_done = threading.Event()
        # Shutdown only waits: it must not queue behind the application's
        # default thread limiter or consume a slot needed by an admitted job.
        self._closer = Runner(CapacityLimiter(math.inf))

    async def __aenter__(self) -> Self:
        if self._closed:
            raise RuntimeError("ProcessPool is closed")
        if self._executor is not None:
            raise RuntimeError("ProcessPool is already open")
        self._executor = ProcessPoolExecutor(
            max_workers=self._max_workers,
            mp_context=multiprocessing.get_context("spawn"),
        )
        return self

    async def __aexit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        await self.aclose()

    def _require_open(self) -> ProcessPoolExecutor:
        if self._closed:
            raise RuntimeError("ProcessPool is closed")
        if self._executor is None:
            raise RuntimeError("Use ProcessPool with 'async with' before submitting jobs")
        return self._executor

    async def run(self, func: Callable[P, R], /, *args: P.args, **kwargs: P.kwargs) -> R:
        """Execute ``func(*args, **kwargs)`` in a worker process and await its result.

        Exceptions from the worker propagate to the caller. Cancelling the
        await does not undo a job's file writes or retry the job.
        """
        self._require_open()
        borrower = object()
        await _acquire_token(self._limiter, borrower)
        try:
            # Recheck after waiting: aclose() may have stopped admission.
            executor = self._require_open()
            self._active += 1
            self._idle.clear()
            try:
                call: Callable[[], R] = functools.partial(func, *args, **kwargs)

                def execute() -> R:
                    # Spawn, submit/pickle and wait off the event loop. Only
                    # call crosses the process boundary, not this closure.
                    return executor.submit(call).result()

                return await Runner().wait_in_thread(execute)
            finally:
                self._active -= 1
                if self._active == 0:
                    self._idle.set()
        finally:
            self._limiter.release_on_behalf_of(borrower)

    def _shutdown(self, executor: ProcessPoolExecutor) -> None:
        # _closer serialises concurrent aclose() calls. Include admitted jobs
        # whose coordinating thread has not yet reached executor.submit().
        if not self._shutdown_done.is_set():
            self._idle.wait()
            executor.shutdown(wait=True)
            self._shutdown_done.set()

    async def aclose(self) -> None:
        """Stop accepting work, finish admitted jobs and join all worker processes.

        Idempotent. Clean-up completes before cancellation is propagated.
        """
        self._closed = True
        executor = self._executor
        if executor is not None:
            try:
                await self._closer.run_uncancellable(self._shutdown, executor)
            finally:
                if self._shutdown_done.is_set():
                    self._executor = None
