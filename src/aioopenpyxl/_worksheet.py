"""Asynchronous proxies around openpyxl worksheets and chartsheets.

Streaming (``iter_rows`` & co.) fetches rows from a worker thread in chunks of
``chunk_size`` rows.  Each chunk costs one thread round trip, so very small
chunks are dominated by overhead; ``chunk_size`` of 256 or more is
recommended (the default is :data:`DEFAULT_CHUNK_SIZE`).

With ``prefetch=0`` (default) the workbook is only busy *during* each hop and
can be used between chunks.  With ``prefetch=N`` a dedicated producer thread
parses up to ``N`` chunks ahead; the workbook runner is then held for the
whole iteration, synchronous access raises
:class:`~aioopenpyxl.WorkbookBusyError`, and awaiting the runner from inside
the loop body (``await ws.run(...)``, ``await wb.save(...)``) would deadlock.

The underlying openpyxl iterator is created, advanced *and closed* in worker
threads under the workbook runner, so concurrent ``async for`` loops and
``read_rows`` calls on one workbook simply queue up instead of failing with
``WorkbookBusyError``, and closing a read-only sheet's XML stream never races
with another thread using the shared zip archive.

Clean-up is robust: whether the loop finishes, is broken out of, is
cancelled (anyio style or native asyncio), or the async generator is simply
dropped and finalised by the event loop later, the producer thread is waited
for and the runner released before the underlying openpyxl iterator is
closed.  No thread ever outlives the iteration and the workbook is never left
busy.  The public iterators are *one* async generator layer around a plain
:class:`_RowStream` object, so the event loop's ``shutdown_asyncgens()`` never
closes a generator that another generator's ``finally`` is already closing.
"""

from __future__ import annotations

import contextlib
import contextvars
import itertools
import queue
import threading
from collections.abc import AsyncGenerator, Callable, Iterable, Iterator
from typing import TYPE_CHECKING, Any, Literal, NoReturn, TypeVar, overload

import anyio
from openpyxl.chartsheet.chartsheet import Chartsheet
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet._write_only import WriteOnlyWorksheet

from ._base import GuardedProxy
from ._errors import BlockingCallError, WorkbookBusyError
from ._executor import Runner, _acquire_token, runner_for
from ._generated import ChartsheetProxy, WorksheetProxy

if TYPE_CHECKING:
    from openpyxl.cell import _CellGetValue, _CellOrMergedCell, _CellSetValue
    from openpyxl.cell.cell import Cell
    from openpyxl.cell.read_only import EmptyCell, ReadOnlyCell
    from openpyxl.descriptors.serialisable import _SerialisableTreeElement
    from openpyxl.worksheet.worksheet import Worksheet

    from ._workbook import AsyncWorkbook

    RawWorksheet = Worksheet | ReadOnlyWorksheet | WriteOnlyWorksheet
    # Regular sheets yield Cell / MergedCell, read-only sheets ReadOnlyCell / EmptyCell.
    CellRow = tuple[_CellOrMergedCell | ReadOnlyCell | EmptyCell, ...]
    ValueRow = tuple[_CellGetValue, ...]

__all__ = ["DEFAULT_CHUNK_SIZE", "AsyncChartsheet", "AsyncWorksheet"]

R = TypeVar("R")

#: Number of rows fetched per thread hop when streaming a worksheet.
DEFAULT_CHUNK_SIZE = 1024


def _append_all(ws: Any, rows: Iterable[Any]) -> int:
    count = 0
    for row in rows:
        ws.append(row)
        count += 1
    return count


class _RowStream:
    """Async iterator over a synchronous openpyxl iterator that lives in worker threads.

    ``make_iterator`` is called in a worker thread under the runner (the
    first hop, or the producer thread), never on the event loop: creating
    ``ws.iter_rows(...)`` must not race with another thread using the workbook
    and must not trip the busy guard.

    * ``prefetch == 0``: every chunk is one ``await runner(...)`` hop; the
      runner is free between hops.
    * ``prefetch > 0``: the runner is acquired on the first chunk and held
      until :meth:`aclose`; a producer thread fills a bounded ``queue.Queue``
      (backend-agnostic hand-off: the blocking ``get`` runs in a worker
      thread) ``prefetch`` chunks ahead.  The producer holds one token of the
      runner's work limiter for its lifetime (acquired on its behalf by the
      consumer), while the consumer's blocking ``get`` is a wait-only hop
      that consumes no token.

    :meth:`aclose` is idempotent -- a second caller waits for the first
    clean-up to finish -- and runs shielded: it stops the producer, waits for
    it (surviving native asyncio cancellation, see
    :meth:`Runner.wait_for_thread`), closes the openpyxl iterator through the
    runner when it was not exhausted (a finished generator needs no close, so
    a complete iteration costs no extra hop) and only then releases the
    runner.  A short chunk ends the stream without a round trip for the empty
    tail.
    """

    __slots__ = (
        "_buffer",
        "_chunk_size",
        "_closing",
        "_depth",
        "_done",
        "_exhausted",
        "_holding",
        "_iterator",
        "_make",
        "_pos",
        "_producer",
        "_rows",
        "_runner",
        "_stop",
        "_token",
    )

    def __init__(
        self,
        runner: Runner,
        make_iterator: Callable[[], Iterator[Any]],
        chunk_size: int,
        prefetch: int,
    ) -> None:
        if chunk_size < 1:
            raise ValueError("chunk_size must be >= 1")
        if prefetch < 0:
            raise ValueError("prefetch must be >= 0")
        self._runner = runner
        self._make = make_iterator
        self._chunk_size = chunk_size
        self._depth = prefetch
        # Created in a worker thread; read on the loop only after that thread is done.
        self._iterator: Iterator[Any] | None = None
        self._rows: list[Any] = []
        self._pos = 0
        self._exhausted = False  # the iterator raised StopIteration (or an error)
        self._closing: anyio.Event | None = None  # set once aclose() has started
        # prefetch > 0 only
        self._holding = False  # we hold the runner
        self._producer: threading.Thread | None = None
        self._buffer: queue.Queue[list[Any] | BaseException] | None = None
        self._stop = threading.Event()
        self._done = threading.Event()
        # (limiter, borrower) while the producer thread holds a work token.
        self._token: tuple[anyio.CapacityLimiter, object] | None = None

    def __aiter__(self) -> _RowStream:
        return self

    async def __anext__(self) -> Any:
        if self._pos < len(self._rows):
            row = self._rows[self._pos]
            self._pos += 1
            return row
        if self._exhausted or self._closing is not None:
            raise StopAsyncIteration
        try:
            chunk = await self._next_chunk()
        except Exception:
            # The iterator raised, so it is finished; cancellation is left to
            # the owner's ``finally`` (the iterator is intact, just not advanced).
            self._exhausted = True
            await self._release()
            raise
        if len(chunk) < self._chunk_size:
            self._exhausted = True
            await self._release()  # free the runner/thread before handing out the tail
        if not chunk:
            raise StopAsyncIteration
        self._rows = chunk
        self._pos = 1
        return chunk[0]

    async def _next_chunk(self) -> list[Any]:
        if not self._depth:
            return await self._runner(self._fetch)
        if self._producer is None:
            await self._runner.acquire()
            self._holding = True
            # The producer is a plain thread, so the consumer borrows its work
            # token for it: with a shared ``CapacityLimiter(1)`` only one
            # workbook parses at a time, producers included.  Cancelled here,
            # nothing has started and ``aclose()`` merely releases the runner.
            limiter = self._runner.work_limiter
            borrower = object()
            # Always ends up holding the token (anyio cannot always undo an
            # acquire interrupted by a native cancellation); record it first,
            # then let the held-back cancellation propagate -- ``aclose()``
            # gives the token back.
            pending = await _acquire_token(limiter, borrower)
            self._token = (limiter, borrower)
            if pending is not None:
                raise pending
            self._buffer = queue.Queue(maxsize=self._depth)
            # Run the producer in a copy of the caller's context so ContextVars
            # (request ids, custom file objects consulting them ...) are visible
            # to the parsing thread exactly as they are to anyio's worker threads.
            context = contextvars.copy_context()
            thread = threading.Thread(
                target=context.run,
                args=(self._produce,),
                name="aioopenpyxl-prefetch",
                daemon=True,
            )
            try:
                thread.start()
            except BaseException:
                self._release_token()
                raise
            self._producer = thread  # only once it really runs (``_done`` will be set)
        assert self._buffer is not None
        # A wait, not work: must not compete for the token the producer holds.
        item = await self._runner.wait_in_thread(self._buffer.get)
        if isinstance(item, BaseException):
            raise item
        return item

    def _release_token(self) -> None:
        token, self._token = self._token, None
        if token is not None:
            limiter, borrower = token
            limiter.release_on_behalf_of(borrower)

    # ------------------------------------------------------------ thread side
    def _fetch(self) -> list[Any]:
        """One chunk (``prefetch == 0``); runs in a worker thread under the runner."""
        it = self._iterator
        if it is None:
            it = self._iterator = self._make()
        return list(itertools.islice(it, self._chunk_size))

    def _produce(self) -> None:
        """Producer thread (``prefetch > 0``): parse ahead into the bounded queue."""
        buffer = self._buffer
        assert buffer is not None
        try:
            it = self._iterator = self._make()
            while not self._stop.is_set():
                chunk = list(itertools.islice(it, self._chunk_size))
                buffer.put(chunk)  # blocks while ``depth`` chunks are waiting
                if len(chunk) < self._chunk_size:
                    return  # exhausted; the consumer sees the short chunk
        except BaseException as exc:  # surfaced to the consumer
            buffer.put(exc)
        finally:
            self._done.set()

    # --------------------------------------------------------------- clean-up
    async def aclose(self) -> None:
        """Stop the producer, close the iterator through the runner, release the runner.

        Safe to call more than once and from any task (the runner has no
        owner check).  Shielded from anyio cancellation; a native asyncio
        cancellation is held back until the producer is done and the iterator
        closed (the runner wait is retried), then re-raised.
        """
        self._rows = []  # a closed stream yields nothing more
        await self._release()

    async def _release(self) -> None:
        """Run :meth:`_cleanup` exactly once; later callers wait for it."""
        if self._closing is not None:
            await self._closing.wait()
            return
        self._closing = anyio.Event()
        try:
            with anyio.CancelScope(shield=True):
                await self._cleanup()
        finally:
            self._closing.set()

    async def _cleanup(self) -> None:
        pending: BaseException | None = None
        try:
            if self._producer is not None:
                self._stop.set()
                # Free a producer blocked on put(); it re-checks ``stop`` and exits.
                assert self._buffer is not None
                with contextlib.suppress(queue.Empty):
                    while True:
                        self._buffer.get_nowait()
                try:
                    await self._runner.wait_for_thread(self._done)
                except anyio.get_cancelled_exc_class() as exc:
                    pending = exc  # the thread *is* done; finish the clean-up first
                self._release_token()  # the close hop below needs a work token
            it = self._iterator
            if it is not None and not self._exhausted:
                close = getattr(it, "close", None)
                if close is not None:
                    # Read-only sheets close a zip member here: never on the
                    # loop, never beside another thread using the archive, and
                    # exactly once however the task is being cancelled.
                    try:
                        if self._holding:
                            await self._runner.offload_uncancellable(close)
                        else:
                            await self._runner.run_uncancellable(close)
                    except anyio.get_cancelled_exc_class() as exc:
                        pending = exc
        finally:
            self._release_token()  # no-op unless an earlier step was skipped
            if self._holding:
                self._holding = False
                self._runner.release()
        if pending is not None:
            raise pending


class _SheetMixin(GuardedProxy):
    """State shared by worksheet and chartsheet wrappers."""

    __slots__ = ()

    _parent: AsyncWorkbook | None

    def _init(self, obj: Any, runner: Runner | None, parent: AsyncWorkbook | None) -> None:
        # A sheet shares the runner of its workbook, so the busy guard holds
        # across every wrapper of that workbook.  An explicit ``runner`` is
        # registered for the workbook (or must be the one already registered:
        # a second semaphore for the same workbook would defeat the guard).
        raw_parent = getattr(obj, "parent", None)
        if raw_parent is not None:
            runner = runner_for(raw_parent, runner=runner)
        elif runner is None:
            runner = Runner()
        object.__setattr__(self, "_obj", obj)
        object.__setattr__(self, "_runner", runner)
        object.__setattr__(self, "_parent", parent)

    @property
    def parent(self) -> AsyncWorkbook | None:
        """The owning :class:`AsyncWorkbook` (shares this sheet's runner)."""
        if self._parent is not None:
            return self._parent
        raw_parent = self._raw.parent
        if raw_parent is None:
            return None
        from ._workbook import AsyncWorkbook

        parent = AsyncWorkbook._from_runner(raw_parent, self._runner)
        object.__setattr__(self, "_parent", parent)
        return parent

    def __repr__(self) -> str:
        try:
            return f"<{type(self).__name__} {self._raw.title!r}>"
        except WorkbookBusyError:
            return f"<{type(self).__name__} busy>"

    def __eq__(self, other: object) -> bool:
        if isinstance(other, _SheetMixin):
            return self._obj is other._obj
        return NotImplemented

    def __hash__(self) -> int:
        return id(self._obj)


class AsyncWorksheet(_SheetMixin, WorksheetProxy):
    """Async-aware proxy for ``Worksheet``, ``ReadOnlyWorksheet`` and ``WriteOnlyWorksheet``.

    Every public openpyxl member is available with its original signature.
    In-memory members (``ws["A1"]``, ``ws.cell(...)``, ``ws.title``,
    ``ws.merge_cells(...)``, ``ws.insert_rows(...)`` ...) are forwarded
    directly.  Members that can block are coroutines / async iterators and run
    in the workbook's worker thread:

    * :meth:`iter_rows` / :meth:`iter_cols` / :attr:`rows` / :attr:`columns` /
      :attr:`values` stream in chunks (one hop per ``chunk_size`` rows);
      ``prefetch=N`` keeps the thread parsing ``N`` chunks ahead of the
      consumer.
    * :meth:`read_rows` reads a whole range in one hop.
    * :meth:`append` / :meth:`append_rows` write rows (write-only sheets spool
      to a temporary file, so they are offloaded).
    * :meth:`fetch`, :meth:`calculate_dimension`, :meth:`close`.
    * :meth:`run` executes any callable against the raw worksheet off-loop.

    While one of those is executing, synchronous access raises
    :class:`~aioopenpyxl.WorkbookBusyError` instead of racing with the thread.
    On read-only worksheets ``ws[key]`` and ``ws.cell(...)`` raise
    :class:`~aioopenpyxl.BlockingCallError` because openpyxl would parse XML
    on the event loop; use :meth:`fetch` / :meth:`read_rows` / :meth:`run`.
    The raw openpyxl object is always reachable through :attr:`wrapped`.
    """

    __slots__ = ("__weakref__", "_obj", "_parent", "_runner")

    def __init__(
        self,
        worksheet: RawWorksheet,
        runner: Runner | None = None,
        parent: AsyncWorkbook | None = None,
    ) -> None:
        self._init(worksheet, runner, parent)

    @property
    def wrapped(self) -> RawWorksheet:
        """The underlying openpyxl worksheet (no busy check)."""
        return self._obj

    @property
    def is_write_only(self) -> bool:
        return isinstance(self._obj, WriteOnlyWorksheet)

    @property
    def is_read_only(self) -> bool:
        return isinstance(self._obj, ReadOnlyWorksheet)

    async def run(self, func: Callable[[RawWorksheet], R]) -> R:
        """Run ``func(raw_worksheet)`` in the workbook's worker thread and return its result.

        The right tool for bulk in-memory work and for anything a read-only
        worksheet would otherwise do synchronously (``ws.cell``, ``ws[key]``).
        """
        return await self._runner(func, self._obj)

    # --------------------------------------------------------------- dunders
    def __getitem__(self, key: Any) -> Any:
        """In-memory ``ws["A1"]`` / ``ws["A1:C3"]`` / ``ws[2]``.

        Raises :class:`~aioopenpyxl.BlockingCallError` on read-only worksheets,
        where this would parse XML on the event loop; use :meth:`fetch`,
        :meth:`read_rows` or ``ws.wrapped[key]`` there.
        """
        if self.is_read_only:
            raise BlockingCallError(
                "ws[key] on a read-only worksheet parses XML synchronously; "
                "use `await ws.fetch(key)`, `await ws.read_rows(...)` or `ws.wrapped[key]`"
            )
        return self._raw[key]

    def __setitem__(self, key: Any, value: Any) -> None:
        self._raw[key] = value

    def __delitem__(self, key: Any) -> None:
        del self._raw[key]

    def __iter__(self) -> NoReturn:
        # Without this, ``for row in ws`` would fall back to the legacy
        # sequence protocol (``ws[0]``, ``ws[1]`` ...) and silently yield nothing.
        raise TypeError(
            "AsyncWorksheet is not iterable synchronously; "
            "use 'async for row in ws' (or ws.wrapped)"
        )

    def __aiter__(self) -> AsyncGenerator[CellRow, None]:
        """``async for row in ws`` -- same as :attr:`rows`."""
        return self.iter_rows()

    @overload
    def cell(  # type: ignore[overload-overlap]
        self, row: int, column: int, value: None = None
    ) -> _CellOrMergedCell: ...

    @overload
    def cell(self, row: int, column: int, value: _CellSetValue) -> Cell: ...

    def cell(self, row: int, column: int, value: Any = None) -> Any:
        """In-memory ``ws.cell(row, column, value=None)``.

        Raises :class:`~aioopenpyxl.BlockingCallError` on read-only worksheets,
        where ``ReadOnlyWorksheet.cell`` parses XML synchronously; use
        :meth:`fetch`, :meth:`read_rows` or :meth:`run` there.
        """
        if self.is_read_only:
            raise BlockingCallError(
                "ws.cell(...) on a read-only worksheet parses XML synchronously; "
                "use `await ws.fetch(key)`, `await ws.read_rows(...)` "
                "or `await ws.run(lambda raw: raw.cell(...))`"
            )
        return self._raw.cell(row, column, value)

    # ------------------------------------------------------------- async I/O
    async def fetch(self, key: Any) -> Any:
        """Off-loop equivalent of ``ws[key]`` (safe for read-only sheets)."""
        return await self._runner(self._obj.__getitem__, key)

    @overload
    def iter_rows(
        self,
        min_row: int | None,
        max_row: int | None,
        min_col: int | None,
        max_col: int | None,
        values_only: Literal[True],
        *,
        chunk_size: int = ...,
        prefetch: int = ...,
    ) -> AsyncGenerator[ValueRow, None]: ...

    @overload
    def iter_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        *,
        values_only: Literal[True],
        chunk_size: int = ...,
        prefetch: int = ...,
    ) -> AsyncGenerator[ValueRow, None]: ...

    @overload
    def iter_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: Literal[False] = False,
        *,
        chunk_size: int = ...,
        prefetch: int = ...,
    ) -> AsyncGenerator[CellRow, None]: ...

    @overload
    def iter_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = False,
        *,
        chunk_size: int = ...,
        prefetch: int = ...,
    ) -> AsyncGenerator[CellRow, None] | AsyncGenerator[ValueRow, None]: ...

    async def iter_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = False,
        *,
        chunk_size: int = DEFAULT_CHUNK_SIZE,
        prefetch: int = 0,
    ) -> AsyncGenerator[Any, None]:
        """Asynchronously iterate rows, fetching ``chunk_size`` rows per thread hop.

        ``prefetch=0`` (default) fetches the next chunk only after the previous
        one was consumed; the runner is released between chunks, so the sheet
        may be touched inside the loop body.  ``prefetch=N`` keeps the worker
        thread parsing up to ``N`` chunks ahead while the body runs, which
        overlaps parsing with your own I/O (and with CPU work on free-threaded
        Python).  The workbook then counts as busy for the whole iteration:
        synchronous access raises :class:`~aioopenpyxl.WorkbookBusyError` and
        awaiting the runner (``ws.run``, ``wb.save`` ...) from the loop body
        deadlocks.

        Each hop is a thread round trip, so prefer ``chunk_size`` >= 256.  Breaking
        out of the loop, cancellation and dropping the generator are all safe:
        the producer thread is waited for and the runner released before
        control returns.  Concurrent loops on one workbook queue up on the
        runner rather than failing.
        """
        stream = _RowStream(
            self._runner,
            lambda: self._obj.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
                values_only=values_only,
            ),
            chunk_size,
            prefetch,
        )
        try:
            async for row in stream:
                yield row
        finally:
            # Runs on normal exhaustion, ``break``, cancellation and when the
            # event loop finalises an abandoned generator: the stream's
            # clean-up is idempotent and cheap once it has already happened.
            await stream.aclose()

    @overload
    def iter_cols(
        self,
        min_col: int | None,
        max_col: int | None,
        min_row: int | None,
        max_row: int | None,
        values_only: Literal[True],
        *,
        chunk_size: int = ...,
        prefetch: int = ...,
    ) -> AsyncGenerator[ValueRow, None]: ...

    @overload
    def iter_cols(
        self,
        min_col: int | None = None,
        max_col: int | None = None,
        min_row: int | None = None,
        max_row: int | None = None,
        *,
        values_only: Literal[True],
        chunk_size: int = ...,
        prefetch: int = ...,
    ) -> AsyncGenerator[ValueRow, None]: ...

    @overload
    def iter_cols(
        self,
        min_col: int | None = None,
        max_col: int | None = None,
        min_row: int | None = None,
        max_row: int | None = None,
        values_only: Literal[False] = False,
        *,
        chunk_size: int = ...,
        prefetch: int = ...,
    ) -> AsyncGenerator[CellRow, None]: ...

    @overload
    def iter_cols(
        self,
        min_col: int | None = None,
        max_col: int | None = None,
        min_row: int | None = None,
        max_row: int | None = None,
        values_only: bool = False,
        *,
        chunk_size: int = ...,
        prefetch: int = ...,
    ) -> AsyncGenerator[CellRow, None] | AsyncGenerator[ValueRow, None]: ...

    async def iter_cols(
        self,
        min_col: int | None = None,
        max_col: int | None = None,
        min_row: int | None = None,
        max_row: int | None = None,
        values_only: bool = False,
        *,
        chunk_size: int = DEFAULT_CHUNK_SIZE,
        prefetch: int = 0,
    ) -> AsyncGenerator[Any, None]:
        """Asynchronously iterate columns (openpyxl does not support this on read-only sheets).

        ``chunk_size`` / ``prefetch`` behave as in :meth:`iter_rows`.
        """
        stream = _RowStream(
            self._runner,
            lambda: self._obj.iter_cols(
                min_col=min_col,
                max_col=max_col,
                min_row=min_row,
                max_row=max_row,
                values_only=values_only,
            ),
            chunk_size,
            prefetch,
        )
        try:
            async for col in stream:
                yield col
        finally:
            await stream.aclose()

    @property
    def rows(self) -> AsyncGenerator[CellRow, None]:
        """``async for row in ws.rows`` -- all rows as cell tuples."""
        return self.iter_rows()

    @property
    def columns(self) -> AsyncGenerator[CellRow, None]:
        """``async for col in ws.columns`` -- all columns as cell tuples."""
        return self.iter_cols()

    @property
    def values(self) -> AsyncGenerator[ValueRow, None]:
        """``async for row in ws.values`` -- all rows as value tuples."""
        return self.iter_rows(values_only=True)

    @overload
    async def read_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: Literal[True] = True,
    ) -> list[ValueRow]: ...

    @overload
    async def read_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        *,
        values_only: Literal[False],
    ) -> list[CellRow]: ...

    @overload
    async def read_rows(
        self,
        min_row: int | None,
        max_row: int | None,
        min_col: int | None,
        max_col: int | None,
        values_only: Literal[False],
    ) -> list[CellRow]: ...

    @overload
    async def read_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = True,
    ) -> list[ValueRow] | list[CellRow]: ...

    async def read_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = True,
    ) -> list[Any]:
        """Read a whole range in a single thread hop and return it as a list.

        The iterator is created inside the hop, so concurrent ``read_rows``
        calls (or one next to a ``run``) serialise on the runner.
        """
        return await self._runner(
            lambda: list(
                self._obj.iter_rows(
                    min_row=min_row,
                    max_row=max_row,
                    min_col=min_col,
                    max_col=max_col,
                    values_only=values_only,
                )
            )
        )

    async def append(self, row: Iterable[Any] | dict[Any, Any]) -> None:
        """Append one row.

        Regular worksheets append in memory, so this runs inline (no thread
        hop).  Write-only worksheets spool XML to disk and are offloaded, which
        costs a thread round trip per row; prefer :meth:`append_rows` for bulk
        inserts, which appends the whole batch in one hop.
        """
        if self.is_write_only:
            await self._runner(self._obj.append, row)
        else:
            self._raw.append(row)

    async def append_rows(self, rows: Iterable[Iterable[Any] | dict[Any, Any]]) -> int:
        """Append many rows in a single thread hop. Returns the number of rows appended."""
        return await self._runner(_append_all, self._obj, rows)

    async def calculate_dimension(self, force: bool = False) -> str:
        """Off-loop ``calculate_dimension`` (read-only sheets may scan the whole sheet).

        ``force`` is only meaningful for read-only worksheets.
        """
        if self.is_read_only:
            return await self._runner(self._obj.calculate_dimension, force)
        return await self._runner(self._obj.calculate_dimension)

    async def close(self) -> None:
        """Off-loop ``WriteOnlyWorksheet.close`` (flushes the spooled XML)."""
        await self._runner(self._obj.close)

    # ----------------------------------------------------------------- helpers
    def _stream(
        self, make_iterator: Callable[[], Iterator[Any]], chunk_size: int, prefetch: int
    ) -> _RowStream:
        """Stream the items of ``make_iterator()`` (called in a worker thread) in chunks."""
        return _RowStream(self._runner, make_iterator, chunk_size, prefetch)


class AsyncChartsheet(_SheetMixin, ChartsheetProxy):
    """Proxy for :class:`openpyxl.chartsheet.Chartsheet` (purely in-memory)."""

    __slots__ = ("__weakref__", "_obj", "_parent", "_runner")

    def __init__(
        self,
        chartsheet: Chartsheet,
        runner: Runner | None = None,
        parent: AsyncWorkbook | None = None,
    ) -> None:
        self._init(chartsheet, runner, parent)

    @classmethod
    def from_tree(cls, node: _SerialisableTreeElement) -> AsyncChartsheet | None:
        """``Chartsheet.from_tree(node)``, wrapped (a parentless sheet gets its own runner)."""
        raw = Chartsheet.from_tree(node)
        return None if raw is None else cls(raw)

    @property
    def wrapped(self) -> Chartsheet:
        """The underlying openpyxl chartsheet (no busy check)."""
        return self._obj

    async def run(self, func: Callable[[Chartsheet], R]) -> R:
        """Run ``func(raw_chartsheet)`` in the workbook's worker thread and return its result."""
        return await self._runner(func, self._obj)
