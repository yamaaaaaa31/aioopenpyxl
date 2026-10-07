"""Thread offloading primitives shared by the async wrappers.

Built on `anyio <https://anyio.readthedocs.io/>`_, so everything works
unchanged on asyncio and on trio.

openpyxl objects are not thread-safe, so every blocking call that belongs to
one workbook is serialised through a per-workbook :class:`Runner`.  Callers
``await`` the runner like ``anyio.to_thread.run_sync``; an optional
:class:`anyio.CapacityLimiter` bounds how many worker threads a workbook (or a
group of workbooks sharing the limiter) may occupy.

One runner per raw workbook: :func:`runner_for` keeps a weak registry mapping
each ``openpyxl.Workbook`` to its :class:`Runner`, so wrapping the same
workbook twice (``Workbook.wrap(raw)`` twice, or ``AsyncWorksheet(raw.active)``
next to a ``Workbook``) still yields *one* lock.

Cancellation: a thread cannot be interrupted, so when the awaiting task is
cancelled we keep waiting (and keep the workbook busy) until the thread has
actually finished, then let the cancellation propagate.  This is what makes
the busy guard trustworthy: ``Runner.busy`` is never False while a thread
still owns the workbook.  anyio's ``abandon_on_cancel=False`` covers anyio
style cancellation (cancel scopes, trio); a *native* asyncio cancellation
(``Task.cancel()``, ``asyncio.wait_for``, ``asyncio.timeout``) is not stopped
by anyio's shield and interrupts the ``await`` while the thread keeps running.
:class:`_Call` and :meth:`Runner.wait_for_thread` close that hole: every
offloaded call records whether it has started and when it has finished, and
the awaiting side waits for the ``done`` event -- holding back further
cancellations until then -- before the runner is released.  The
``CancelledError`` then propagates as usual.

The mutual exclusion primitive is deliberately *not* :class:`anyio.Lock`:

* ``anyio.Lock.release()`` checks that the releasing task is the owner.  The
  streaming iterators hold the runner across ``yield`` (``prefetch > 0``) and
  an abandoned ``async for`` is finalised by the event loop's async-generator
  hook from a *different* task, which would raise and leak the lock forever.
* ``anyio.Lock.locked()`` raises ``NoEventLoopError`` outside a running loop,
  which would make something as innocent as ``Workbook().sheetnames`` blow up
  in synchronous code.

A binary :class:`anyio.Semaphore` has neither problem: its ``release()`` has
no owner check on either backend and ``.value`` is a plain attribute read.
"""

from __future__ import annotations

import functools
import math
import threading
import weakref
from collections.abc import Callable
from types import TracebackType
from typing import Any, ParamSpec, TypeVar

import anyio
import anyio.to_thread
from anyio import CapacityLimiter

__all__ = ["Runner", "run_sync", "runner_for"]

P = ParamSpec("P")
R = TypeVar("R")


async def run_sync(
    limiter: CapacityLimiter | None,
    func: Callable[P, R],
    /,
    *args: P.args,
    **kwargs: P.kwargs,
) -> R:
    """Run ``func(*args, **kwargs)`` in a worker thread and await the result.

    Context variables are propagated.  ``limiter`` bounds thread concurrency
    (``None`` = anyio's default limiter).  The call waits for the thread to
    finish when cancelled through anyio (cancel scopes, trio); a native
    asyncio cancellation interrupts the wait -- :class:`Runner` adds the
    bookkeeping needed to survive that.
    """
    call = functools.partial(func, *args, **kwargs)
    return await anyio.to_thread.run_sync(call, abandon_on_cancel=False, limiter=limiter)


def _wait_limiter() -> CapacityLimiter:
    """A fresh, unbounded :class:`CapacityLimiter` for one wait-only thread hop.

    Waiting must not consume tokens of the limiter that bounds *work*: the
    thread that blocks on ``done.wait()`` or on the prefetch queue would
    otherwise queue up behind other workbooks' parsing and, with a shared
    ``CapacityLimiter(1)``, deadlock against the producer that holds the only
    token.  A limiter is bound to the async backend it is first used on, so
    rather than caching one per backend (which would need ``sniffio``, not a
    dependency of anyio any more) a throw-away unbounded one is created per
    hop; it is a tiny object next to the thread hop it accompanies.
    """
    return CapacityLimiter(math.inf)


def _cancellation_in(
    exc: BaseException, cancelled_exc: type[BaseException]
) -> BaseException | None:
    """*exc* if it is a cancellation, else the cancellation it was raised while handling, if any."""
    cause: BaseException | None = exc
    while cause is not None and not isinstance(cause, cancelled_exc):
        cause = cause.__context__
    return cause


async def _acquire_token(
    limiter: CapacityLimiter, borrower: object, *, uncancellable: bool = False
) -> BaseException | None:
    """``limiter.acquire_on_behalf_of(borrower)`` that never leaks a token.

    anyio cannot always undo an acquire interrupted by a native asyncio
    cancellation: a cancellation at the checkpoint right after taking the
    token surfaces as a ``RuntimeError`` while ``borrower`` *is* holding it.
    Cancellable mode (default) then gives the token back and raises the
    cancellation itself; the caller holds the token exactly when this
    returns.  ``uncancellable=True`` shields anyio-style cancellation, retries
    through native cancellation and returns the held-back cancellation once
    the token is held (keeping a token anyio could not undo).
    """
    cancelled_exc = anyio.get_cancelled_exc_class()
    if not uncancellable:
        try:
            await limiter.acquire_on_behalf_of(borrower)
        except BaseException as exc:
            if borrower in limiter.statistics().borrowers:
                limiter.release_on_behalf_of(borrower)
                cancel = _cancellation_in(exc, cancelled_exc)
                if cancel is not None and cancel is not exc:
                    raise cancel from None
            raise
        return None
    pending: BaseException | None = None
    with anyio.CancelScope(shield=True):
        while True:
            try:
                await limiter.acquire_on_behalf_of(borrower)
                return pending
            except BaseException as exc:
                cancel = _cancellation_in(exc, cancelled_exc)
                if borrower in limiter.statistics().borrowers:
                    if cancel is not None:
                        return cancel  # acquired after all; report the cancellation
                    limiter.release_on_behalf_of(borrower)
                    raise
                if cancel is not None and cancel is exc:
                    pending = exc
                    continue  # nothing taken: ask again
                raise
    raise AssertionError("unreachable")  # the shielded scope never swallows anything


class _Call:
    """A callable handed to a worker thread, with start/finish bookkeeping.

    ``done`` is set once the call has returned (or raised).  The awaiting
    side calls :meth:`abandon` when its ``await`` was interrupted: if the
    thread has not started the call yet it never will (anyio skips jobs whose
    future is cancelled, and the flag closes the tiny window between that
    check and the actual call), otherwise the caller must wait for ``done``.
    The check and the start are both taken under one lock, so exactly one of
    the two outcomes happens.
    """

    __slots__ = ("_abandoned", "_func", "_lock", "_started", "done", "error", "result")

    def __init__(self, func: Callable[[], Any]) -> None:
        self._func = func
        self._lock = threading.Lock()
        self._started = False
        self._abandoned = False
        self.done = threading.Event()
        # The outcome, for callers that were interrupted while the call ran.
        self.result: Any = None
        self.error: BaseException | None = None

    def __call__(self) -> Any:
        with self._lock:
            if self._abandoned:
                return None  # nobody is waiting for the result
            self._started = True
        try:
            self.result = self._func()
            return self.result
        except BaseException as exc:
            self.error = exc
            raise
        finally:
            self.done.set()

    def outcome(self) -> Any:
        """The recorded result (or re-raise the recorded exception); only valid once ``done``."""
        if self.error is not None:
            raise self.error
        return self.result

    def abandon(self) -> bool:
        """Stop the call from starting; return True when it already runs (wait for ``done``)."""
        with self._lock:
            if not self._started:
                self._abandoned = True
                return False
            return True


class Runner:
    """Serialises blocking openpyxl calls for one workbook.

    Awaiting the runner runs the callable in a worker thread while holding the
    workbook's binary semaphore, so two coroutines working on the same
    workbook can never touch openpyxl state concurrently from two threads.
    The semaphore is held until the thread has finished, cancellation
    included -- native asyncio cancellation too (see the module docstring).

    The runner is also an async context manager: ``async with runner:`` holds
    the workbook (``busy`` is True) across arbitrary awaits -- the streaming
    iterators use this while a producer thread parses ahead.  Inside such a
    block use :meth:`offload` rather than awaiting the runner again, which
    would deadlock on the already-held semaphore.

    The semaphore is created lazily on first acquisition (always inside an
    event loop) so ``fast_acquire=True`` is honoured regardless of where the
    runner was constructed, and ``busy`` can be read from synchronous code
    outside any event loop.
    """

    __slots__ = ("_limiter", "_sem")

    def __init__(self, limiter: CapacityLimiter | None = None) -> None:
        self._limiter = limiter
        self._sem: anyio.Semaphore | None = None

    @property
    def limiter(self) -> CapacityLimiter | None:
        return self._limiter

    @property
    def work_limiter(self) -> CapacityLimiter:
        """The limiter that bounds this runner's *work* (anyio's default when ``limiter`` is None).

        Only valid inside an event loop.  The prefetch producer thread borrows
        one of its tokens for its whole lifetime.
        """
        if self._limiter is not None:
            return self._limiter
        return anyio.to_thread.current_default_thread_limiter()

    @property
    def busy(self) -> bool:
        """True while a call is executing (or queued) in the worker thread.

        Safe to read from any thread and outside an event loop.
        """
        sem = self._sem
        return sem is not None and sem.value == 0

    async def acquire(self) -> None:
        """Mark the workbook busy, waiting for any in-flight call to finish.

        Pair with :meth:`release` in a ``try``/``finally`` (or use
        ``async with runner:``).  Releasing from a different task than the
        one that acquired is allowed.
        """
        sem = self._sem
        if sem is None:
            # fast_acquire: skip the scheduling checkpoint on the uncontended
            # path; the thread hop that always follows is a checkpoint anyway.
            sem = self._sem = anyio.Semaphore(1, max_value=1, fast_acquire=True)
        await sem.acquire()

    def release(self) -> None:
        """Mark the workbook idle again (counterpart of :meth:`acquire`)."""
        if self._sem is None:
            raise RuntimeError("Runner.release() called without a matching acquire()")
        self._sem.release()

    async def __aenter__(self) -> Runner:
        await self.acquire()
        return self

    async def __aexit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ) -> None:
        self.release()

    async def wait_for_thread(self, done: threading.Event) -> None:
        """Wait until ``done`` (set by a thread when it has finished) -- whatever happens.

        Cancellation of any kind, repeated or not, is deferred while waiting:
        this is used from ``finally`` blocks that must not give the workbook
        back while a thread still owns it.  Once the thread is done the last
        cancellation that was held back is re-raised, so a cancel request is
        delayed, never lost.  Anything that is not a cancellation
        (``KeyboardInterrupt``, a closed event loop ...) propagates at once.
        Returns immediately when ``done`` is already set, so the normal path
        pays nothing.

        The first attempt blocks in a worker thread on ``done.wait()`` (exact,
        no polling) taken from the wait-only limiter, so it never queues
        behind -- or deadlocks against -- work holding the runner's limiter.
        anyio-style cancellation is shielded; only a native asyncio
        cancellation can interrupt it, and then the wait degrades to short
        sleeps with a capped back-off.
        """
        if done.is_set():
            return
        cancelled_exc = anyio.get_cancelled_exc_class()
        pending: BaseException | None = None
        try:
            with anyio.CancelScope(shield=True):
                await anyio.to_thread.run_sync(done.wait, limiter=_wait_limiter())
        except cancelled_exc as exc:
            pending = exc
        delay = 0.001
        while not done.is_set():
            try:
                with anyio.CancelScope(shield=True):
                    await anyio.sleep(delay)
            except cancelled_exc as exc:
                pending = exc
            delay = min(delay * 2, 0.01)
        if pending is not None:
            raise pending

    async def offload(
        self,
        func: Callable[P, R],
        /,
        *args: P.args,
        **kwargs: P.kwargs,
    ) -> R:
        """Run ``func`` in a worker thread *without* acquiring (caller already holds the runner).

        Like awaiting the runner, this never returns (or raises) while the
        thread is still running, native asyncio cancellation included, and
        the work-limiter token is held for exactly as long as the thread runs.
        """
        return await self._offload_with(True, func, *args, **kwargs)

    async def wait_in_thread(
        self,
        func: Callable[P, R],
        /,
        *args: P.args,
        **kwargs: P.kwargs,
    ) -> R:
        """Like :meth:`offload`, for a call that only *waits* (``queue.get`` ...).

        Holds no work token: the prefetch consumer blocks on its queue while
        the producer thread holds the runner's limiter token.
        """
        return await self._offload_with(False, func, *args, **kwargs)

    async def _offload_with(
        self,
        work: bool,
        func: Callable[P, R],
        /,
        *args: P.args,
        **kwargs: P.kwargs,
    ) -> R:
        """One thread hop whose work-limiter token we manage ourselves.

        anyio's own ``limiter=`` would hand the token back the moment a native
        cancellation interrupts the ``await`` -- while the thread keeps
        running, so another workbook could start a worker and exceed the
        limit.  Instead the token is borrowed on behalf of the call before
        the hop (cancellable while queued, as usual), the thread runs on an
        unbounded limiter, and the token is returned only once the thread has
        finished, cancellation or not.
        """
        call = _Call(functools.partial(func, *args, **kwargs))
        limiter = self.work_limiter if work else None
        if limiter is not None:
            await _acquire_token(limiter, call)
        try:
            return await run_sync(_wait_limiter(), call)
        finally:
            try:
                if not call.done.is_set() and call.abandon():
                    await self.wait_for_thread(call.done)
            finally:
                if limiter is not None:
                    limiter.release_on_behalf_of(call)

    async def _offload_to_completion(
        self, target: Callable[[], R]
    ) -> tuple[R, BaseException | None]:
        """Run ``target`` in a worker thread exactly once; return (result, held-back cancellation).

        anyio-style cancellation is shielded.  A native asyncio cancellation
        is held back: if it interrupts the hand-off before the thread started
        the call, the call is handed off again; if the call is already
        running, its completion is awaited and its recorded result used (or
        its exception raised -- the callable's own error always wins).  The
        work-limiter token is held from before the first hand-off until the
        thread has finished.
        """
        cancelled_exc = anyio.get_cancelled_exc_class()
        limiter = self.work_limiter
        pending = await _acquire_token(limiter, target, uncancellable=True)
        try:
            with anyio.CancelScope(shield=True):
                while True:
                    call = _Call(target)
                    try:
                        result = await run_sync(_wait_limiter(), call)
                    except cancelled_exc as exc:
                        pending = exc
                        if not call.abandon():
                            continue  # never started: hand it off again
                        try:
                            await self.wait_for_thread(call.done)
                        except cancelled_exc as later:
                            pending = later
                        result = call.outcome()
                    return result, pending
            raise AssertionError("unreachable")  # the shielded scope never swallows anything
        finally:
            limiter.release_on_behalf_of(target)

    async def offload_uncancellable(
        self,
        func: Callable[P, R],
        /,
        *args: P.args,
        **kwargs: P.kwargs,
    ) -> R:
        """:meth:`offload` that runs ``func`` to completion no matter how it is cancelled.

        The last cancellation held back while ``func`` ran is re-raised once
        it has finished (see :meth:`_offload_to_completion`).  For clean-up
        (``close()``) that must happen exactly once even when the task is
        being torn down.
        """
        target = functools.partial(func, *args, **kwargs)
        result, pending = await self._offload_to_completion(target)
        if pending is not None:
            raise pending
        return result

    async def run_uncancellable(
        self,
        func: Callable[P, R],
        /,
        *args: P.args,
        **kwargs: P.kwargs,
    ) -> R:
        """Acquire, run ``func`` in a worker thread and release -- completing whatever happens.

        Like awaiting the runner, but a cancellation of any kind (anyio or
        native asyncio, repeated or not) cannot skip or interrupt any of the
        three steps: waiting for the semaphore is retried, the call runs
        exactly once and the semaphore is released.  The last cancellation
        held back is re-raised afterwards.  Used for ``__aexit__`` clean-up,
        which must close the workbook even when the task is being torn down
        with ``Task.cancel()``.
        """
        cancelled_exc = anyio.get_cancelled_exc_class()
        pending: BaseException | None = None
        with anyio.CancelScope(shield=True):
            while True:
                try:
                    await self.acquire()
                    break
                except cancelled_exc as exc:
                    pending = exc  # anyio's semaphore undoes a half-acquire; just retry
            try:
                result, later = await self._offload_to_completion(
                    functools.partial(func, *args, **kwargs)
                )
            finally:
                self.release()
        if later is not None:
            pending = later
        if pending is not None:
            raise pending
        return result

    async def __call__(
        self,
        func: Callable[P, R],
        /,
        *args: P.args,
        **kwargs: P.kwargs,
    ) -> R:
        await self.acquire()
        try:
            return await self.offload(func, *args, **kwargs)
        finally:
            self.release()


# --------------------------------------------------------------------------- registry
# Raw openpyxl workbook -> its Runner.  Weak keys: the entry disappears with the
# workbook.  A lock guards the dictionary because ``busy`` and wrapping may be
# used from synchronous code in arbitrary threads.
_runners: weakref.WeakKeyDictionary[Any, Runner] = weakref.WeakKeyDictionary()
_runners_lock = threading.Lock()


def runner_for(
    workbook: Any,
    limiter: CapacityLimiter | None = None,
    *,
    runner: Runner | None = None,
) -> Runner:
    """Return *the* :class:`Runner` of a raw ``openpyxl.Workbook``, creating it on first use.

    Every wrapper of the same workbook -- ``Workbook.wrap(raw)`` called twice,
    ``AsyncWorksheet(raw.active)`` created by hand, the ``parent`` of a loose
    sheet -- shares this runner, so the busy guard and the serialisation hold
    across all of them.

    ``limiter`` is applied when the runner is created.  Passing a *different*
    limiter for a workbook that already has a runner raises :class:`ValueError`
    (silently ignoring it would make ``wb.limiter`` lie); ``None`` means "no
    preference" and accepts whatever the existing runner uses.  ``runner``
    registers an already constructed runner for a workbook that has none
    (and must be that workbook's runner if it has one).
    """
    with _runners_lock:
        existing = _runners.get(workbook)
        if existing is None:
            existing = runner if runner is not None else Runner(limiter)
            _runners[workbook] = existing
        elif runner is not None and runner is not existing:
            raise ValueError("this openpyxl workbook already has a different Runner")
        elif limiter is not None and existing.limiter is not limiter:
            raise ValueError(
                "this openpyxl workbook is already wrapped with a different limiter; "
                "pass limiter=None (or the same limiter) to share its Runner"
            )
        return existing
