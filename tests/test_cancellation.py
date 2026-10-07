"""Cancellation while a worker thread owns the workbook.

Two flavours are covered:

* anyio cancellation (``CancelScope``, ``move_on_after``, a task group being
  cancelled) on both backends: the runner is held until the thread has
  finished and the clean-up paths (``__aexit__``, abandoned ``prefetch``
  loops, interrupted loads) still run.
* native asyncio cancellation (``Task.cancel``, ``wait_for``,
  ``asyncio.timeout``), marked ``@asyncio_only``: anyio's
  ``CancelScope(shield=True)`` / ``abandon_on_cancel=False`` only cover
  anyio-style cancellation; a native asyncio cancellation interrupts the await
  while the worker thread keeps running.  These tests pin down that the runner
  is nevertheless held until the thread has finished.  On trio every
  cancellation goes through cancel scopes, which the first group covers.
"""

from __future__ import annotations

import asyncio
import contextlib
import sys
import threading
from collections.abc import AsyncGenerator, Awaitable, Callable, Coroutine
from pathlib import Path
from typing import Any, cast

import anyio
import anyio.lowlevel
import anyio.to_thread
import openpyxl
import pytest

import aioopenpyxl
from aioopenpyxl import WorkbookBusyError
from tests.conftest import (
    ROWS,
    abandons_generator,
    no_errors_logged,
    prefetch_threads,
    wait_until_idle,
)

pytestmark = pytest.mark.anyio

asyncio_only = pytest.mark.parametrize("anyio_backend", ["asyncio"])


# ======================================================== anyio cancellation
async def test_cancelled_call_keeps_the_lock_until_the_thread_finishes() -> None:
    wb = aioopenpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    started = threading.Event()
    release = threading.Event()
    finished = threading.Event()
    scope = anyio.CancelScope()

    def slow(raw: openpyxl.Workbook) -> None:
        started.set()
        release.wait(5)
        finished.set()

    async def run_slow() -> None:
        with scope:
            await wb.run(slow)

    async with anyio.create_task_group() as tg:
        tg.start_soon(run_slow)
        await anyio.to_thread.run_sync(started.wait, 5)
        scope.cancel()
        await anyio.sleep(0.05)
        # Cancelled, but the thread is still running: still guarded.
        assert not finished.is_set()
        assert wb._runner.busy is True
        with pytest.raises(WorkbookBusyError):
            ws.cell(1, 1, 1)
        release.set()

    assert finished.is_set()
    # anyio semantics: the thread's result is delivered and the cancellation
    # takes effect at the task's next checkpoint.
    assert scope.cancel_called
    assert wb._runner.busy is False
    ws.cell(1, 1, 1)


async def test_cancelled_load_workbook_does_not_hang_or_leak(sample_xlsx: Path) -> None:
    async def load() -> None:
        await aioopenpyxl.load_workbook(sample_xlsx)

    with anyio.fail_after(10):
        async with anyio.create_task_group() as tg:
            tg.start_soon(load)
            await anyio.lowlevel.checkpoint()
            tg.cancel_scope.cancel()


@abandons_generator
async def test_cancelled_prefetch_loop_leaves_no_thread_behind(
    sample_xlsx: Path, caplog: pytest.LogCaptureFixture
) -> None:
    async with aioopenpyxl.load_workbook(sample_xlsx, read_only=True) as wb:
        ws = wb["Data"]
        with anyio.move_on_after(0.05) as scope:
            async for _ in ws.iter_rows(prefetch=2, chunk_size=4):
                await anyio.sleep(10)
        assert scope.cancelled_caught
        await wait_until_idle(wb)
        assert prefetch_threads() == []
        assert wb._runner.busy is False
        assert await ws.read_rows(max_row=2) == ROWS[:2]
    assert no_errors_logged(caplog)


async def test_cancelled_prefetch_loop_with_aclosing_cleans_up_synchronously(
    sample_xlsx: Path,
) -> None:
    async with aioopenpyxl.load_workbook(sample_xlsx, read_only=True) as wb:
        ws = wb["Data"]
        with anyio.move_on_after(0.05) as scope:
            async with contextlib.aclosing(ws.iter_rows(prefetch=2, chunk_size=4)) as rows:
                async for _ in rows:
                    await anyio.sleep(10)
        assert scope.cancelled_caught
        # aclosing() ran aclose() inside the cancelled scope: the join is shielded,
        # so by now the producer is gone and the runner released.
        assert prefetch_threads() == []
        assert wb._runner.busy is False
        assert await ws.read_rows(max_row=2) == ROWS[:2]


async def test_cancelled_scope_still_closes_the_archive_on_exit(sample_xlsx: Path) -> None:
    holder: dict[str, aioopenpyxl.Workbook] = {}
    with anyio.move_on_after(0.05) as scope:
        async with aioopenpyxl.load_workbook(sample_xlsx, read_only=True) as wb:
            holder["wb"] = wb
            await anyio.sleep(10)
    assert scope.cancelled_caught
    # ``__aexit__`` is shielded, so the read-only zip archive was closed.
    assert cast(Any, holder["wb"].wrapped)._archive.fp is None


async def test_cancelled_scope_still_closes_the_archive_for_workbook_aexit(
    sample_xlsx: Path,
) -> None:
    wb = await aioopenpyxl.load_workbook(sample_xlsx, read_only=True)
    with anyio.move_on_after(0.05) as scope:
        async with wb:
            await anyio.sleep(10)
    assert scope.cancelled_caught
    assert cast(Any, wb.wrapped)._archive.fp is None


# ============================================== native asyncio cancellation
Canceller = Callable[..., Awaitable[Any]]

#: Default deadline before the native cancellation fires.
_DEADLINE = 0.05


async def _via_wait_for(coro: Coroutine[Any, Any, Any], deadline: float = _DEADLINE) -> None:
    with pytest.raises((asyncio.TimeoutError, TimeoutError)):
        await asyncio.wait_for(coro, deadline)


async def _via_task_cancel(coro: Coroutine[Any, Any, Any], deadline: float = _DEADLINE) -> None:
    task = asyncio.ensure_future(coro)
    await asyncio.sleep(deadline)
    task.cancel()
    with pytest.raises(asyncio.CancelledError):
        await task


async def _via_timeout(coro: Coroutine[Any, Any, Any], deadline: float = _DEADLINE) -> None:
    if sys.version_info < (3, 11):
        coro.close()
        pytest.skip("asyncio.timeout() needs Python 3.11+")
    with pytest.raises(TimeoutError):
        async with asyncio.timeout(deadline):
            await coro


CANCELLERS = pytest.mark.parametrize(
    "cancel",
    [_via_wait_for, _via_task_cancel, _via_timeout],
    ids=["wait_for", "Task.cancel", "asyncio.timeout"],
)


@asyncio_only
@CANCELLERS
async def test_run_holds_the_runner_until_the_thread_finishes(cancel: Canceller) -> None:
    wb = aioopenpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    started = threading.Event()
    finished = threading.Event()

    def slow(raw: openpyxl.Workbook) -> None:
        started.set()
        threading.Event().wait(0.3)
        finished.set()

    await cancel(wb.run(slow))
    # The cancellation only came back once the thread was done.
    assert started.is_set() and finished.is_set()
    assert wb._runner.busy is False
    ws.cell(1, 1, 1)
    assert ws["A1"].value == 1


@asyncio_only
async def test_busy_guard_holds_while_a_natively_cancelled_call_is_still_running() -> None:
    wb = aioopenpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    started = threading.Event()
    release = threading.Event()

    def slow(raw: openpyxl.Workbook) -> None:
        started.set()
        release.wait(5)

    task = asyncio.ensure_future(wb.run(slow))
    await asyncio.to_thread(started.wait, 5)
    task.cancel()
    await asyncio.sleep(0.05)
    # Cancelled, but the thread still owns the workbook: still guarded, task still alive.
    assert not task.done()
    assert wb._runner.busy is True
    with pytest.raises(WorkbookBusyError):
        ws.cell(1, 1, 1)
    # A second cancellation while waiting is swallowed as well.
    task.cancel()
    await asyncio.sleep(0.02)
    assert not task.done()
    release.set()
    with pytest.raises(asyncio.CancelledError):
        await task
    assert wb._runner.busy is False
    ws.cell(1, 1, 1)


@asyncio_only
async def test_cancelled_before_the_thread_started_does_not_hang() -> None:
    # The job may be cancelled while it is still queued (``_Call.abandon`` path).
    wb = aioopenpyxl.Workbook()
    ran = threading.Event()

    async def run() -> None:
        await wb.run(lambda raw: ran.set())

    for _ in range(20):
        task = asyncio.ensure_future(run())
        await asyncio.sleep(0)  # let it reach the thread hand-off
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await asyncio.wait_for(task, 5)
        assert wb._runner.busy is False
    assert await wb.run(lambda raw: raw.sheetnames) == ["Sheet"]


def _slow_rows(ws: aioopenpyxl.AsyncWorksheet, started: threading.Event) -> None:
    """Make every row of the raw sheet take a while to parse."""
    raw: Any = ws.wrapped
    real = raw.iter_rows

    def slow_iter_rows(**kwargs: Any) -> Any:
        for row in real(**kwargs):
            started.set()
            threading.Event().wait(0.02)
            yield row

    raw.iter_rows = slow_iter_rows


async def _consume(rows: AsyncGenerator[Any, None]) -> None:
    # ``aclosing``: the stream is closed from the cancelled task itself rather
    # than by the event loop's finaliser, whose timing depends on when the
    # abandoned generator is garbage collected (the traceback kept by
    # ``pytest.raises`` can keep it alive for the rest of the test).
    async with contextlib.aclosing(rows) as it:
        async for _ in it:
            await asyncio.sleep(0.005)


@asyncio_only
@CANCELLERS
@pytest.mark.parametrize("prefetch", [0, 2])
async def test_iter_rows_leaves_no_thread_behind(cancel: Canceller, prefetch: int) -> None:
    wb = aioopenpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    await ws.append_rows([[i] for i in range(200)])
    started = threading.Event()
    _slow_rows(ws, started)

    # 200 rows x 20 ms keeps the worker busy for ~4 s, so a 0.5 s deadline
    # reliably fires *after* parsing started and *before* it finished, even on
    # a loaded CI runner (a 50 ms deadline occasionally beat the thread start).
    await cancel(_consume(ws.iter_rows(chunk_size=2, prefetch=prefetch)), deadline=0.5)
    assert started.is_set(), "cancellation fired before the worker started parsing"
    # The consumer task was cancelled: the generator's ``finally`` (or the
    # loop's finaliser) closes the stream; either way no thread survives it
    # and the workbook is idle again shortly after.
    for _ in range(100):
        if not wb._runner.busy and not prefetch_threads():
            break
        await asyncio.sleep(0.01)
    assert prefetch_threads() == []
    assert wb._runner.busy is False
    assert ws.max_row == 200
    assert await ws.read_rows(max_row=1) == [(0,)]


@asyncio_only
@CANCELLERS
async def test_prefetch_cleanup_survives_native_cancellation_inside_aclose(
    cancel: Canceller,
) -> None:
    # ``aclose()`` itself (stop + wait for the producer + release) is what runs
    # when the generator is closed from the cancelled task.
    wb = aioopenpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    await ws.append_rows([[i] for i in range(200)])
    started = threading.Event()
    _slow_rows(ws, started)
    rows = ws.iter_rows(chunk_size=3, prefetch=2)
    await anext(rows)

    await cancel(rows.aclose())
    assert prefetch_threads() == []
    assert wb._runner.busy is False
    assert await ws.read_rows(max_row=1) == [(0,)]


# ------------------------------------------------------------- loading (I)
@asyncio_only
@CANCELLERS
async def test_interrupted_read_only_load_closes_the_orphaned_archive(
    sample_xlsx: Path, cancel: Canceller, monkeypatch: pytest.MonkeyPatch
) -> None:
    produced: list[openpyxl.Workbook] = []
    real = openpyxl.load_workbook

    def slow_load(*args: Any, **kwargs: Any) -> openpyxl.Workbook:
        threading.Event().wait(0.2)  # the cancellation lands while we "parse"
        wb = real(*args, **kwargs)
        produced.append(wb)
        return wb

    monkeypatch.setattr(openpyxl, "load_workbook", slow_load)

    async def load() -> None:
        await aioopenpyxl.load_workbook(sample_xlsx, read_only=True)

    await cancel(load())
    # The thread finished (the cancellation waited for it) and produced a
    # workbook nobody owns: its zip archive was closed before we got here.
    assert len(produced) == 1
    assert cast(Any, produced[0])._archive.fp is None


# ------------------------------------------------------------ __aexit__ (J)
def _hold_runner(
    wb: aioopenpyxl.Workbook,
) -> tuple[asyncio.Future[None], threading.Event, threading.Event]:
    """Occupy ``wb``'s runner from another task; returns (task, started, release)."""
    started = threading.Event()
    release = threading.Event()

    def slow(raw: Any) -> None:
        started.set()
        release.wait(5)

    task = asyncio.ensure_future(wb.run(slow))
    return task, started, release


@asyncio_only
@CANCELLERS
@pytest.mark.parametrize("via", ["load_workbook", "Workbook"])
async def test_aexit_closes_the_workbook_despite_native_cancellation_while_waiting(
    sample_xlsx: Path, cancel: Canceller, via: str
) -> None:
    # Set everything up *outside* the cancelled coroutine so that the deadline
    # is guaranteed to land while ``__aexit__`` is waiting for the semaphore,
    # however slow loading the file happens to be on this machine.
    if via == "load_workbook":
        cm = aioopenpyxl.load_workbook(sample_xlsx, read_only=True)
        wb = await cm.__aenter__()
        leave = cm.__aexit__(None, None, None)
    else:
        wb = await aioopenpyxl.load_workbook(sample_xlsx, read_only=True)
        leave = wb.__aexit__(None, None, None)
    task, started, release = _hold_runner(wb)
    await asyncio.to_thread(started.wait, 5)  # the holder is in its thread now
    assert wb._runner.busy

    # ``__aexit__`` waits for the semaphore; the cancellation lands (0.05 s)
    # and the holder lets go a little later.
    asyncio.get_running_loop().call_later(0.15, release.set)
    await cancel(leave)
    await task
    archive = cast(Any, wb.wrapped)._archive
    assert archive.fp is None
    assert wb._runner.busy is False


@asyncio_only
async def test_run_uncancellable_runs_exactly_once_and_reraises_the_cancellation() -> None:
    wb = aioopenpyxl.Workbook()
    calls: list[int] = []
    started = threading.Event()
    release = threading.Event()

    def work(raw: Any) -> int:
        calls.append(1)
        started.set()
        release.wait(5)
        return 42

    task = asyncio.ensure_future(wb._runner.run_uncancellable(work, wb.wrapped))
    await asyncio.to_thread(started.wait, 5)
    task.cancel()
    await asyncio.sleep(0.02)
    task.cancel()
    await asyncio.sleep(0.02)
    assert not task.done() and wb._runner.busy is True
    release.set()
    with pytest.raises(asyncio.CancelledError):
        await task
    assert calls == [1]
    assert wb._runner.busy is False
    # Without interference the result comes back as usual.
    assert await wb._runner.run_uncancellable(lambda: 7) == 7


# --------------------------------------------------- shared limiter waits (K)
@asyncio_only
async def test_cancelled_call_does_not_wait_for_other_workbooks_sharing_the_limiter() -> None:
    # anyio hands A's token back the moment the native cancellation lands, so
    # B (queued behind A) starts.  Waiting for A's thread must not need a
    # token of that limiter, or A's cancellation would complete only after B.
    limiter = anyio.CapacityLimiter(1)
    a = aioopenpyxl.Workbook(limiter=limiter)
    b = aioopenpyxl.Workbook(limiter=limiter)
    a_started = threading.Event()
    a_done = threading.Event()
    b_started = threading.Event()
    b_release = threading.Event()

    def slow_a(raw: Any) -> None:
        a_started.set()
        threading.Event().wait(0.2)
        a_done.set()

    def slow_b(raw: Any) -> None:
        b_started.set()
        b_release.wait(5)

    task_a = asyncio.ensure_future(a.run(slow_a))
    await asyncio.to_thread(a_started.wait, 5)  # A holds the only token
    task_b = asyncio.ensure_future(b.run(slow_b))
    await asyncio.sleep(0.02)  # B is queued on the limiter
    task_a.cancel()
    with pytest.raises(asyncio.CancelledError):
        await task_a
    # A came back as soon as its own thread was done ...
    assert a_done.is_set()
    assert a._runner.busy is False
    # ... while B (which took the token anyio released on the cancellation)
    # is still running.
    await asyncio.to_thread(b_started.wait, 5)
    assert not task_b.done()
    b_release.set()
    await task_b
    assert b._runner.busy is False


# ------------------------------------------- limiter tokens under cancellation (N, O)
@asyncio_only
async def test_cancelled_prefetch_start_does_not_leak_a_limiter_token() -> None:
    # anyio cannot undo an ``acquire_on_behalf_of`` interrupted right after the
    # token was taken; the token must still come back, or a shared
    # CapacityLimiter(1) would hang every later read.
    limiter = anyio.CapacityLimiter(1)
    wb = aioopenpyxl.Workbook(limiter=limiter)
    ws = wb.active
    assert ws is not None
    await ws.append_rows([[i] for i in range(30)])
    for _ in range(5):
        rows = ws.iter_rows(chunk_size=3, prefetch=2)
        task = asyncio.ensure_future(anext(rows))
        await asyncio.sleep(0)  # one step: inside the limiter acquire
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task
        await rows.aclose()
        assert limiter.borrowed_tokens == 0
        assert wb._runner.busy is False
    with anyio.fail_after(5):
        assert await ws.read_rows(max_row=1) == [(0,)]
        assert [r async for r in ws.iter_rows(values_only=True, prefetch=2)] == [
            (i,) for i in range(30)
        ]


@asyncio_only
async def test_cancelled_worker_keeps_its_limiter_token_until_the_thread_finishes() -> None:
    # anyio's own ``limiter=`` hands the token back the moment the native
    # cancellation lands; the Runner manages the token itself so B cannot
    # start while A's thread is still running.
    limiter = anyio.CapacityLimiter(1)
    a = aioopenpyxl.Workbook(limiter=limiter)
    b = aioopenpyxl.Workbook(limiter=limiter)
    lock = threading.Lock()
    active = 0
    peak = 0
    a_started = threading.Event()
    a_release = threading.Event()
    b_done = threading.Event()

    def worker(release: threading.Event | None, started: threading.Event | None) -> None:
        nonlocal active, peak
        with lock:
            active += 1
            peak = max(peak, active)
        if started is not None:
            started.set()
        if release is not None:
            release.wait(5)
        with lock:
            active -= 1

    task_a = asyncio.ensure_future(a.run(lambda raw: worker(a_release, a_started)))
    await asyncio.to_thread(a_started.wait, 5)
    assert limiter.borrowed_tokens == 1

    def worker_b(raw: Any) -> None:
        worker(None, None)
        b_done.set()

    task_b = asyncio.ensure_future(b.run(worker_b))
    await asyncio.sleep(0.02)
    task_a.cancel()
    await asyncio.sleep(0.05)
    # A is cancelled but still running: it keeps the only token, B waits.
    assert not task_a.done() and not b_done.is_set()
    assert limiter.borrowed_tokens == 1
    a_release.set()
    with pytest.raises(asyncio.CancelledError):
        await task_a
    await task_b
    assert peak == 1
    assert limiter.borrowed_tokens == 0
