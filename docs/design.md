# Design: concurrency and safety

This page explains how aioopenpyxl keeps the event loop unblocked and openpyxl's thread-safety
assumptions intact. The [README](../README.md) has the user-facing rules; this is the reasoning
behind them. Nothing here is a performance claim: the design is argued from correctness.

## Why a wrapper, and why threads

openpyxl is synchronous and CPU-bound: loading a workbook, saving one or streaming the rows of
a read-only sheet parses or writes XML for as long as it takes, and calling it on the event loop
stalls every other request your server is handling. Moving those calls to a thread is not
enough on its own, because openpyxl is not thread-safe: a second coroutine touching the same
workbook while a thread is saving it corrupts the file silently. aioopenpyxl deals with both.
The operations that can block run in worker threads, all work on one workbook is serialised,
and synchronous access to a workbook that a thread currently owns raises instead of racing.

This is the right layer for async support. openpyxl's work is XML parsing and generation, pure
CPU time with nothing to wait on, so making openpyxl itself `async` would gain nothing; what an
async application needs is for that work to happen on a worker thread while the event loop
waits for it, which is exactly what a wrapper can provide.

The existing workbook/worksheet wrappers use threads and share the application's GIL on
standard builds. They do not parallelise Python CPU work. Free-threaded Python removes that
limit for independent workbooks. `lxml` is not needed; on free-threaded builds importing an
extension without free-threading support can re-enable the GIL.

For an independent interpreter/GIL, use `process=True` for a persistent workbook, or
`ProcessPool` to move an entire job across an explicit process boundary. Raw Workbook/Cell
references and open archives cannot be shared with a child as live objects; the process APIs
take callables and serializable data instead.

## What is offloaded

Only operations that can actually block are offloaded. In-memory operations are forwarded
directly: a thread round trip costs far more than setting a cell, so wrapping every `ws.cell()`
call would add latency without protecting anything. The table in the README lists which call
runs where.

Streaming parameters:

- Each chunk of `iter_rows` is one thread round trip, so keep `chunk_size` at 256 or more
  (the default is `DEFAULT_CHUNK_SIZE`).
- `prefetch=N` lets the worker thread keep parsing while the loop body awaits something else
  (database inserts, HTTP calls). The sheet counts as busy for the whole loop in that mode.
- For large exports use `Workbook(write_only=True)` with `append_rows`: rows are spooled to disk
  instead of kept as cell objects. On a write-only sheet every `await ws.append(row)` is a thread
  round trip, `append_rows` does the whole batch in one.
- A short final chunk ends the stream without an extra round trip for the empty tail.

## `process=True` workbooks

`Workbook(process=True)` and `load_workbook(..., process=True)` keep the same `Workbook` class
and async `run`, `save`, `to_bytes`, `close` and context-manager APIs. Each process workbook
owns a `ProcessPoolExecutor(max_workers=1, mp_context=spawn)`, so all commands reach the same
live workbook. It is never returned to the parent and reloaded between commands. In-memory,
read-only and write-only workbooks all work this way; a read-only workbook's open archive
remains in the child.

`Workbook(process=True)` is lazy: its constructor starts no process and creates no raw
workbook in the parent. The first awaited operation or context entry creates it in the child.
`load_workbook(..., process=True)` completes the load before returning. Paths are opened in
the child after resolving paths relative to the parent's working directory. File-like inputs
are snapshotted in a coordinating thread, preserving their seek position, then loaded from
bytes. Saving to a file-like object returns the Excel bytes from
the child and writes them in the coordinating thread.

Synchronous sheet/attribute/cell access is deliberately refused with `BlockingCallError`:
`wb.active`, `wb["Sheet"]`, `wb.sheetnames`, `wb.wrapped` and property assignments must move
inside `await wb.run(fn)`. Waiting for a remote property synchronously would block the event
loop. The callback receives a real synchronous openpyxl workbook, so use plain openpyxl code
inside it, without `await`. `cloudpickle` supports lambdas and local functions as well as
module-level functions. Captured data and return values must be serializable; event loops,
locks, open handles and the async workbook wrapper itself cannot be captured. Results are
copies: returning a Cell/Worksheet/Workbook does not create a live remote reference, so return
plain data and keep all edits inside `run`. Callback errors retain their type and message,
with the child traceback attached as the cause.

All coordination, callback/result serialization and process startup happen off the event
loop under the workbook's `Runner`. It serializes operations and holds cancellation until
an in-flight command finishes. `close()` and context exit close the workbook, its VBA copy
and any unsaved write-only spool files, then join the worker even during cancellation.
A failed/cancelled load closes its unreturned workbook/pool. Process-mode `close()` is
idempotent and ends the workbook's lifetime: subsequent commands fail instead of spawning
a new worker. Crashed workers are not restarted and commands are never retried automatically.

This is one worker process per live process workbook, not a shared application-wide pool;
bound concurrent workbook lifetimes to match the deployment's memory/CPU budget. `limiter=`
still bounds the coordinating thread work across workbooks, not the count of idle processes.
Use the shared `ProcessPool` job API below when whole jobs can run as one call and should share
a fixed number of application-wide workers. Serialization still has a cost and processes
still share the host's CPU/memory limits; neither API guarantees request latency.

## Process pool jobs

`async with ProcessPool(max_workers=2) as pool:` owns a reusable standard-library
`ProcessPoolExecutor`. `await pool.run(func, *args, **kwargs)` runs the synchronous callable
in a child process, on both asyncio and trio. This is an explicit job API, separate from
`wb.run(fn)` / `ws.run(fn)`, which by default execute against the existing live object in a
thread. With `process=True`, `wb.run(fn)` uses the persistent child described above. The flag
selects a process backend; it does not turn off the GIL for one call in a shared interpreter.

Keep the whole create/load, edit and save sequence inside the job, returning bytes, a path
or extracted plain data. Returning a workbook to the parent and then editing it there would
put the CPU work back in the server. Per-cell process messages would require remote proxies
for cells, styles and attributes; this API deliberately uses a single coarse-grained call.

The pool always uses `multiprocessing.get_context("spawn")`, including on Linux, so it never
forks an already multithreaded ASGI server. Callables must be importable module-level functions;
local functions, lambdas, live async wrappers and open file handles are not suitable payloads.
Arguments, results and worker exceptions follow standard `ProcessPoolExecutor` pickling rules.
Standalone scripts must guard their entry point with `if __name__ == "__main__":`. An
interactive interpreter/notebook's local definitions are not importable worker functions.
Context variables are not propagated: pass request metadata explicitly.

Process startup, submission and waiting for the result happen in a coordinating thread, so
they do not synchronously wait on the event loop. Excel work happens in the child. Pickling,
transport and memory copies still have a cost, and CPU/memory limits are shared with the
server. This is isolation from the same interpreter's GIL, not a guarantee of API latency or
faster Excel processing.

Admission is bounded to `max_workers` jobs (default **1**). A caller waiting for a slot can
cancel without submitting work. A cancelled admitted call waits for its coordinating thread
and any submitted job to finish; the existing `Runner` handles both anyio cancellation and
native asyncio `Task.cancel()` / timeouts. A running process is never killed midway through
a file save. Consequently a timeout is not a hard execution deadline, and cancellation does
not undo writes. Worker exceptions propagate; a crashed worker breaks the pool as in
`ProcessPoolExecutor`, and jobs are never automatically retried.

`aclose()` first rejects new work, waits for admitted calls (including ones not yet submitted)
and joins all workers off the event loop. It is idempotent, serialises concurrent close calls
and completes before propagating cancellation. Calls still waiting for admission when the
pool closes are rejected. A pool cannot be reopened or nested, and should be used within one
application event loop. Jobs must coordinate their own output paths; separate jobs writing
the same file are not protected by a shared workbook runner.

### FastAPI example

Put the synchronous job in an importable module, such as `excel_jobs.py`:

```python
import io
import openpyxl


def render_report(rows: list[list[str | int]]) -> bytes:
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    for row in rows:
        ws.append(row)
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
```

Create and close the pool in the application's lifespan, rather than per request (`app.py`):

```python
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, Response
from aioopenpyxl import ProcessPool
from excel_jobs import render_report


@asynccontextmanager
async def lifespan(app: FastAPI):
    async with ProcessPool(max_workers=2) as pool:
        app.state.excel_pool = pool
        yield


app = FastAPI(lifespan=lifespan)


@app.get("/report.xlsx")
async def report(request: Request):
    content = await request.app.state.excel_pool.run(
        render_report, [["name", "count"], ["item", 7]]
    )
    return Response(
        content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="report.xlsx"'},
    )
```

Start it as an importable application, for example `uvicorn app:app`. Each ASGI server worker
owns its own pool: four server workers with `max_workers=2` can start eight Excel processes.
Choose the worker count for the deployment's available CPU and memory.

## One binary semaphore per workbook

Every offloaded call on a workbook and its sheets goes through the same `anyio.Semaphore(1)`,
so two coroutines can never run openpyxl code on the same workbook in two threads at once.
Different workbooks are independent. Concurrent reads on one workbook
(`asyncio.gather(ws.read_rows(), ws.read_rows())`, two `async for` loops, a loop next to a
`run(...)`) queue up on the semaphore like any other call; the openpyxl iterator is created,
advanced and closed inside the worker thread.

The semaphore is created lazily, so `Runner.busy` (and therefore `Workbook().sheetnames`) can
be read outside an event loop. `Runner` exposes `acquire()` / `release()` / `async with runner`
/ `offload()` for holding a workbook across awaits.

### Why a semaphore and not a lock

`anyio.Lock` checks that the releasing task is the owner. `iter_rows(prefetch=N)` holds the
workbook across `yield`, and an `async for` that is abandoned (`break`, an exception, a dropped
generator) is finalised by the event loop from a *different* task; with a lock that release
would fail and the workbook would stay busy forever. A lock also raises when inspected outside a
running loop, whereas `Workbook().sheetnames` in plain synchronous code just works. The price is
that there is no owner, hence no self-deadlock detection: inside a `prefetch>0` loop,
`await ws.run(...)`, `await wb.save(...)`, `await ws.read_rows(...)` or a second `async for` on
the same workbook waits for a semaphore that the loop itself holds. That is a genuine hang, not
an error. Collect what you need and do it after the loop, or use another workbook.

## The busy guard: `WorkbookBusyError`

While the semaphore is held, synchronous access through the wrappers (`ws.cell(...)`,
`wb.sheetnames`, `ws.title = ...`, even `repr(wb)`, which renders `<Workbook busy>`) raises
`WorkbookBusyError`. Await the pending operation first or move the work into `run(...)`. With
`prefetch=0` the semaphore is released between chunk fetches of `iter_rows`, so touching the
sheet inside the `async for` body is fine; with `prefetch>0` it is held for the whole loop.

When you may leave a `prefetch>0` loop early, wrap the iterator in `contextlib.aclosing(...)`:
after a bare `break` the generator is only closed when the event loop's finaliser gets to it,
and until then the workbook still reads as busy (awaiting it is fine, it waits; synchronous
access raises).

Every synchronous forwarder goes through one guarded accessor (`GuardedProxy._raw`), which
refuses access while the workbook's runner is executing something in a worker thread.
`wb.wrapped` / `ws.wrapped` return the raw openpyxl objects without that check; if you use them,
keep blocking calls inside `run(...)`.

### What the guard does not cover

The guard lives on the wrapper objects (`Workbook`, `AsyncWorksheet`, `AsyncChartsheet`): every
forwarded attribute, method and item access checks the runner first. Objects *returned* by those
accesses are openpyxl's own: a `Cell` from `ws["A1"]` or `ws.cell()`, the cells yielded by
`iter_rows()`, a `Font`, a `ColumnDimension`, `ws.merged_cells`, `cell.parent`. Holding one of
them across an `await` on the same workbook and mutating it (`cell.value = ...`, `cell.font = ...`)
touches the workbook while a worker thread owns it, and nothing stops that.

This is a deliberate boundary rather than an oversight. Wrapping every returned object would
mean a proxy per cell on the row-streaming hot path, and proxies break openpyxl internals that
rely on `isinstance(x, Cell)` or identity (for example `ws.freeze_panes = ws["B2"]`, or passing
cells back into `ws.append`). The rule for users is the same one any thread-based wrapper has:
do not edit a workbook, through any object obtained from it, while an operation on that workbook
is in flight; `run(...)` exists precisely so that bulk edits happen inside the serialised thread.

## `BlockingCallError`

`ws["A1"]` and `ws.cell(...)` on a read-only sheet would parse XML synchronously, so they raise
`BlockingCallError`; use `await ws.fetch("A1")`, `read_rows`, `run` or the explicit
`ws.wrapped["A1"]`. Synchronous `for row in ws` raises `TypeError` instead of silently yielding
nothing through the legacy sequence protocol; `async for row in ws` is the supported spelling.

## Cancellation waits for the thread

A thread cannot be interrupted, so when the awaiting task is cancelled aioopenpyxl keeps the
semaphore and waits for the thread to finish (anyio's `abandon_on_cancel=False`); the
cancellation takes effect at the task's next checkpoint. The busy guard therefore never lies, at
the price of cancellation taking as long as the running call.

### Native asyncio cancellation

This holds for *native* asyncio cancellation too (`Task.cancel()`, `asyncio.wait_for`,
`asyncio.timeout`), which anyio's shield does not stop: every offloaded call carries a
completion event (a `threading.Event`), and the awaiting side waits for it, holding back
repeated cancellations until the thread is done, before the semaphore is released
(`Runner.wait_for_thread`, backend-agnostic). The `CancelledError` / `TimeoutError` then
propagates as usual.

### Uncancellable clean-up

Clean-up goes one step further (`Runner.run_uncancellable` / `Runner.offload_uncancellable`
carry acquire, call and release to completion whatever cancellation is in flight, re-raising it
afterwards):

- leaving `async with load_workbook(...)` or `async with wb` closes the workbook even if the
  task is being cancelled while it waits for the semaphore (another coroutine was using the
  workbook);
- a load interrupted by a cancellation closes the workbook it produced: `load_workbook` runs
  under the workbook's `Runner`, so the cancellation waits for the thread and the freshly parsed
  workbook (and its zip archive) is not left without an owner;
- a cancelled `prefetch` loop still waits for its producer thread and closes the underlying
  iterator, through the worker thread and under the runner, never on the event loop (for
  read-only sheets closing the iterator touches the shared zip archive). A fully consumed
  iterator needs no extra hop.

No thread outlives the iteration and no read-only zip archive is left open.

### Loop shutdown

The public iterators are a single async-generator layer around a plain stream object
(`_RowStream`) whose idempotent `aclose()` can be awaited by several callers, so asyncio's
`shutdown_asyncgens()` closes an abandoned `iter_rows` cleanly at loop shutdown. (Three nested
async generators would make `shutdown_asyncgens()` and the parent generator's `finally` race to
close the inner ones.)

## `limiter=`: bounding worker threads

`load_workbook`, `load_workbook_bytes`, `Workbook` and `Workbook.wrap` accept an
`anyio.CapacityLimiter` that bounds how many worker threads those workbooks may occupy
(default: anyio's global limiter).

- **The limiter counts work, including the `prefetch` producer.** That thread holds one token
  for the whole iteration (the consumer acquires it on the thread's behalf before starting it
  and releases it once the thread has finished), so a shared `CapacityLimiter(1)` really means
  one parsing thread at a time, producers included.
- **The Runner manages the token itself.** anyio's own `limiter=` hands the token back the
  moment a native cancellation lands, while the thread is still running. The runner borrows the
  work token before the hop and returns it only once the thread has finished; the thread hop
  itself runs on an unbounded limiter. A token taken by an `acquire_on_behalf_of` that is
  interrupted right after acquisition is still returned.
- **Waiting never consumes a token.** The consumer's blocking `queue.get` and the "wait for the
  thread to finish" hops used after a cancellation run on a separate unbounded limiter
  (`Runner.wait_in_thread`), so they cannot queue behind or deadlock against another workbook's
  work.

## Runner registry and wrapper cache

The semaphore belongs to the *raw* openpyxl workbook: a weak registry
(`aioopenpyxl._executor.runner_for`) maps each one to its `Runner`, so wrapping the same
workbook twice with `Workbook.wrap(raw)` or building `AsyncWorksheet(raw.active)` by hand
shares the lock instead of defeating it. Asking for a different `limiter` for an already wrapped
workbook raises `ValueError` (`None` accepts the existing one); an explicit `runner=` passed to
`AsyncWorksheet(...)` / `AsyncChartsheet(...)` is registered for the sheet's workbook, and one
that differs from the registered runner raises `ValueError` too. Parentless sheets keep whatever
runner they are given.

Members returning sheets return wrappers (`wb.active`, `wb["Data"]`, `wb.create_sheet()`,
`wb.create_chartsheet()`, `wb.worksheets`, `wb.chartsheets`, `wb.copy_worksheet()`), and
`ws.parent` is the `Workbook`. Sheet wrappers are cached per workbook, so
`wb["Data"] is wb["Data"]`; a sheet wrapped by hand still gets a working async parent sharing
its runner.

## Backends

Everything goes through anyio (`anyio.Semaphore`, `anyio.to_thread`, `anyio.CancelScope`), so
asyncio and trio are both supported and tested. Pick the backend with
`anyio.run(main, backend="trio")` or just run your framework as usual. The wrappers are written
to behave the same on both backends.

## Generated forwarders and `.pyi` shims

`Workbook`, `AsyncWorksheet` and `AsyncChartsheet` expose every public member of their openpyxl
counterparts. The forwarders live in `src/aioopenpyxl/_generated.py` and are generated from the
`types-openpyxl` stubs by `tools/gen_proxies.py`, which is why `types-openpyxl` is a runtime
dependency (the generated signatures reference stub-only names). The generator also checks the
generated surface for signature drift against the runtime, runtime members unknown to the stubs
and unresolvable defaults, and expands descriptor annotations (`Typed[...]`, `Set[...]`,
`Alias`) to the value types they produce.

The same tool writes the `.pyi` module shims (`aioopenpyxl/styles/__init__.pyi` is
`from openpyxl.styles import *`, and so on for every openpyxl module). At runtime openpyxl's
sub-packages are aliased through `sys.modules`, so `import aioopenpyxl as openpyxl` keeps working
for helpers; the shims make the same imports resolve for ty, mypy and pyright. Both the
forwarders and the shims must be regenerated after an openpyxl / stubs bump; see
[CONTRIBUTING.md](../CONTRIBUTING.md).
