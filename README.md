# aioopenpyxl

[![CI](https://github.com/yamaaaaaa31/aioopenpyxl/actions/workflows/ci.yml/badge.svg)](https://github.com/yamaaaaaa31/aioopenpyxl/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/aioopenpyxl.svg)](https://pypi.org/project/aioopenpyxl/)
[![Downloads](https://static.pepy.tech/badge/aioopenpyxl)](https://pepy.tech/project/aioopenpyxl)
[![Python versions](https://img.shields.io/pypi/pyversions/aioopenpyxl.svg)](https://pypi.org/project/aioopenpyxl/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/yamaaaaaa31/aioopenpyxl/blob/main/LICENSE)

Asynchronous wrapper around [openpyxl](https://openpyxl.readthedocs.io/) with the complete
openpyxl API surface in its default thread mode, built on [anyio](https://anyio.readthedocs.io/)
so it works on **asyncio and trio** alike. openpyxl is synchronous, CPU-bound and not thread-safe: calling it on the event
loop stalls every other request, and merely moving it to a thread lets a second coroutine corrupt
the file silently. By default, aioopenpyxl runs blocking operations in worker threads, serialises
all work on one workbook, and makes synchronous access to a workbook a thread currently owns
raise instead of racing. It is an independent project, not affiliated with or endorsed by
openpyxl; please report problems with this wrapper here.

For CPU-heavy work that must not share the server's GIL, opt into `process=True` or use the
explicit [`ProcessPool`](https://github.com/yamaaaaaa31/aioopenpyxl/blob/main/docs/design.md#process-pool-jobs) API below.

## Install

```sh
uv add aioopenpyxl   # or: pip install aioopenpyxl
```

Process execution is available starting in **0.2.0**. Upgrade an existing installation with
`uv add 'aioopenpyxl>=0.2.0'` or `pip install --upgrade aioopenpyxl`.

Supports Python 3.10–3.15, including the free-threaded 3.14t and 3.15t builds; depends on
`anyio`, `openpyxl`, `types-openpyxl`, `cloudpickle` and `typing_extensions`.

## Quick start

```python
import aioopenpyxl

async with aioopenpyxl.load_workbook("data.xlsx", read_only=True) as wb:  # read, off the loop
    ws = wb.active
    async for row in ws.iter_rows(values_only=True):
        ...

wb = aioopenpyxl.Workbook()  # accepts openpyxl's write_only=... and iso_dates=... arguments
ws = wb.active
ws["A1"] = "header"  # in-memory: forwarded directly, no thread hop
await ws.append_rows(rows)  # whole batch in one thread hop
await wb.save("output.xlsx")  # off the event loop


def fill(raw_ws):  # bulk in-memory work: plain openpyxl code, one hop, workbook locked meanwhile
    for i, item in enumerate(items, start=2):
        raw_ws.cell(i, 1, item.name)


await ws.run(fill)
```

Wrap an openpyxl workbook you already have with `aioopenpyxl.Workbook.wrap(existing_workbook)`.
`Workbook` is a real class: `isinstance`, annotations and subclassing work as with
`openpyxl.Workbook`, and a subclass can keep attributes of its own.

## Process execution

| Mode | Choose it when | How to edit |
| --- | --- | --- |
| Default `Workbook()` / `load_workbook(...)` | You want the existing synchronous cell API and async I/O | Direct cell access or `run`; blocking operations use threads |
| `process=True` | You want to keep one live workbook in a separate interpreter/GIL | All sheet/cell operations inside `await wb.run(fn)` |
| `ProcessPool(max_workers=N)` | You want complete Excel jobs to share a bounded pool | An importable synchronous job creates/loads, edits and saves the workbook |

Select a persistent child process with `Workbook(process=True)` or
`load_workbook(..., process=True)`; keep cell/sheet operations inside the existing `run` API:

```python
def edit(raw_wb):  # plain synchronous openpyxl code; local functions/lambdas also work
    ws = raw_wb["Data"]
    ws["A1"] = "updated"


async with aioopenpyxl.load_workbook("input.xlsx", process=True) as wb:
    await wb.run(edit)
    titles = await wb.run(lambda raw: raw.sheetnames)
    await wb.save("output.xlsx")

async with aioopenpyxl.Workbook(process=True) as wb:
    await wb.run(lambda raw: raw.create_sheet("Report").append(["item", 7]))
    xlsx = await wb.to_bytes()
```

The workbook stays in one worker for its lifetime. Synchronous access such as `wb.active`,
`wb["Data"]` or `wb.wrapped` raises `BlockingCallError`: move it inside `run` so the event
loop never waits synchronously for a process response. Captured values and results must be
serializable; results are copies. `await wb.close()` / context exit joins the worker and
ends the workbook's lifetime. See [process workbooks](https://github.com/yamaaaaaa31/aioopenpyxl/blob/main/docs/design.md#processtrue-workbooks).

### Shared pool for complete jobs

Use `async with ProcessPool(max_workers=2) as pool:` and `await pool.run(export, rows)` to
share a fixed number of workers across complete jobs. `export` must be an importable,
module-level synchronous function that creates/loads, edits and saves the workbook, with
picklable arguments/results. At most `max_workers` jobs are admitted at once (default **1**).
Running jobs finish before cancellation or shutdown completes. See
[process pool jobs](https://github.com/yamaaaaaa31/aioopenpyxl/blob/main/docs/design.md#process-pool-jobs) for a complete FastAPI example and the
`spawn` requirements (including a guarded `if __name__ == "__main__":` entry point for scripts
using either process API). The default workbook mode continues to use threads.

## What runs where

Rows without an explicit process option describe the default thread mode.

| Operation | Where | Why |
| --- | --- | --- |
| `Workbook(process=True)` / `load_workbook(..., process=True)`: `run`, `save`, `to_bytes`, `close` | same worker process for the workbook's lifetime | independent interpreter/GIL; cell operations belong inside `run` |
| `await pool.run(fn, *args, **kwargs)` inside `async with ProcessPool(...)` | worker process | complete jobs with an independent interpreter/GIL |
| `await load_workbook(...)`, `await load_workbook_bytes(data)`, `await wb.save(...)`, `await wb.to_bytes()`, `await wb.close()` | worker thread | real I/O, XML parsing / writing |
| `async for row in ws.iter_rows(..., chunk_size=1024, prefetch=0)`, `ws.rows`, `ws.values`, `ws.columns`, `async for row in ws` | worker thread, one hop per `chunk_size` rows; with `prefetch=N` the thread parses `N` chunks ahead | read-only sheets parse XML lazily |
| `await ws.read_rows(...)`, `await ws.fetch(key)`, `await ws.calculate_dimension()` | worker thread, one hop | coarse-grained reads |
| `await ws.append_rows(rows)` | worker thread, one hop | bulk insert |
| `await ws.append(row)` | inline for regular sheets, thread for write-only sheets | write-only sheets spool to a temp file |
| `await wb.run(fn)`, `await ws.run(fn)` | worker thread, one hop | *your* bulk in-memory work |
| `ws["A1"]`, `ws.cell()`, `ws.title`, `ws.merge_cells()`, `ws.insert_rows()`, `wb.create_sheet()`, styles, dimensions ... | inline on the event loop | in-memory operations |
| `ws["A1"]`, `ws.cell()` on a **read-only** sheet | raises `BlockingCallError` | openpyxl would parse XML on the event loop; use `fetch` / `read_rows` / `run` |

Keep `chunk_size` at 256 or more (each chunk is one thread round trip); for large exports use
`Workbook(write_only=True)` with `append_rows`. The internals are explained in [docs/design.md](https://github.com/yamaaaaaa31/aioopenpyxl/blob/main/docs/design.md).

## Rules of the road (default thread mode)

- **A thread owns the workbook while an `await` is pending.** Synchronous access meanwhile
  (`ws.cell(...)`, `wb.sheetnames`) raises `WorkbookBusyError`; await first or
  move the work into `run(...)`.
- **Inside a `prefetch>0` loop, never await the same workbook.** The loop holds the semaphore, so
  `await ws.run(...)` or a second `async for` from the body hangs rather than erroring. Use
  `contextlib.aclosing(...)` around the iterator when you may leave such a loop early.
- **Read-only sheets have no synchronous cell access.** `ws["A1"]` / `ws.cell()` raise
  `BlockingCallError`; use `await ws.fetch(key)`, `await ws.read_rows(...)` or `run`.
- **Cancellation waits for the thread**, including native `Task.cancel()` / `asyncio.timeout`;
  it takes effect at the next checkpoint, and `async with` always closes the workbook.
- **`import aioopenpyxl as openpyxl` keeps the helpers synchronous.** `aioopenpyxl.styles`,
  `aioopenpyxl.workbook.Workbook` ... are openpyxl's own modules; only the top-level `Workbook`,
  `load_workbook` and `open` are asynchronous.
- **The guard covers the wrappers, not the objects they hand out.** A `Cell` from `ws["A1"]`,
  a `Font`, a `ColumnDimension` or `cell.parent` is a plain openpyxl object; mutating it while an
  `await` on the same workbook is pending races with the thread. Finish or await the operation
  first, or do such edits inside `run(...)`.
- **`wb.wrapped` / `ws.wrapped` are unguarded**: the raw openpyxl objects, no busy check. Keep
  blocking calls on them inside `run(...)`.

## Differences from openpyxl

| openpyxl | aioopenpyxl |
| --- | --- |
| `load_workbook(...)` | `await load_workbook(...)` or `async with load_workbook(...) as wb` |
| `wb.save(f)`, `wb.close()` | `await wb.save(f)`, `await wb.to_bytes()`, `await wb.close()` |
| `for row in ws.iter_rows()` / `ws.rows` / `ws.values` / `for row in ws` | `async for` versions (chunked); `iter_rows` / `iter_cols` gain `chunk_size=` and `prefetch=` |
| `ws.append(row)` | `await ws.append(row)`; plus `await ws.append_rows(rows)` |
| `ws["A1"]` on a read-only sheet | `BlockingCallError`; use `await ws.fetch("A1")` |
| – | `await ws.read_rows(...)`, `await x.run(fn)`, `x.wrapped`, `limiter=` |

In the default thread mode, every other public member is forwarded with its typed openpyxl
signature. `Workbook` and `AsyncWorksheet` are generic in the kind of worksheet (`Workbook()` is a `Workbook[Worksheet]`,
`load_workbook(..., read_only=True)` yields a `Workbook[ReadOnlyWorksheet]`), so `ws.run(fn)`
accepts a callback written for the concrete openpyxl class. The full table and the typing,
sub-module and subclass rules are in [docs/compatibility.md](https://github.com/yamaaaaaa31/aioopenpyxl/blob/main/docs/compatibility.md).

## Development

```sh
uv sync && uv run pytest -q                                 # both backends, random order
uv run ruff check . && uv run ruff format --check .
uv run ty check src tests && uv run mypy
uv run python tools/gen_proxies.py && uv run ruff format src/aioopenpyxl/   # after an openpyxl / stubs bump
```

See [CONTRIBUTING.md](https://github.com/yamaaaaaa31/aioopenpyxl/blob/main/CONTRIBUTING.md) for the workflow and [CHANGELOG.md](https://github.com/yamaaaaaa31/aioopenpyxl/blob/main/CHANGELOG.md) for history.

## Acknowledgements

All the real work is done by [openpyxl](https://openpyxl.readthedocs.io/) (MIT), which this
package depends on rather than vendors. The typed forwarders and `.pyi` shims are generated from
the [types-openpyxl](https://pypi.org/project/types-openpyxl/) stubs maintained in
[typeshed](https://github.com/python/typeshed) (Apache-2.0); concurrency primitives come from
[anyio](https://anyio.readthedocs.io/) (MIT).

## License

[MIT](https://github.com/yamaaaaaa31/aioopenpyxl/blob/main/LICENSE). openpyxl, typeshed and anyio remain under their own licenses.
