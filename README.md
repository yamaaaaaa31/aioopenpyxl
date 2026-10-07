# aioopenpyxl

[![CI](https://github.com/yamaaaaaa31/aioopenpyxl/actions/workflows/ci.yml/badge.svg)](https://github.com/yamaaaaaa31/aioopenpyxl/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/aioopenpyxl.svg)](https://pypi.org/project/aioopenpyxl/)
[![Python versions](https://img.shields.io/pypi/pyversions/aioopenpyxl.svg)](https://pypi.org/project/aioopenpyxl/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Asynchronous wrapper around [openpyxl](https://openpyxl.readthedocs.io/) with the complete
openpyxl API surface, built on [anyio](https://anyio.readthedocs.io/) so it works on **asyncio
and trio** alike. openpyxl is synchronous, CPU-bound and not thread-safe: calling it on the event
loop stalls every other request, and merely moving it to a thread lets a second coroutine corrupt
the file silently. aioopenpyxl runs the operations that can block in worker threads, serialises
all work on one workbook, and makes synchronous access to a workbook a thread currently owns
raise instead of racing. It is an independent project, not affiliated with or endorsed by
openpyxl; please report problems with this wrapper here.

## Install

```sh
uv add aioopenpyxl   # or: pip install aioopenpyxl
```

Python 3.10 or newer (including the free-threaded 3.14t build); depends on `anyio`, `openpyxl`,
`types-openpyxl` and `typing_extensions`.

## Quick start

```python
import aioopenpyxl

async with aioopenpyxl.load_workbook("data.xlsx", read_only=True) as wb:  # read, off the loop
    ws = wb.active
    async for row in ws.iter_rows(values_only=True):
        ...

wb = aioopenpyxl.Workbook()  # same signature as openpyxl.Workbook(write_only=..., iso_dates=...)
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

## What runs where

| Operation | Where | Why |
| --- | --- | --- |
| `await load_workbook(...)`, `await load_workbook_bytes(data)`, `await wb.save(...)`, `await wb.to_bytes()`, `await wb.close()` | worker thread | real I/O, XML parsing / writing |
| `async for row in ws.iter_rows(..., chunk_size=1024, prefetch=0)`, `ws.rows`, `ws.values`, `ws.columns`, `async for row in ws` | worker thread, one hop per `chunk_size` rows; with `prefetch=N` the thread parses `N` chunks ahead | read-only sheets parse XML lazily |
| `await ws.read_rows(...)`, `await ws.fetch(key)`, `await ws.calculate_dimension()` | worker thread, one hop | coarse-grained reads |
| `await ws.append_rows(rows)` | worker thread, one hop | bulk insert |
| `await ws.append(row)` | inline for regular sheets, thread for write-only sheets | write-only sheets spool to a temp file |
| `await wb.run(fn)`, `await ws.run(fn)` | worker thread, one hop | *your* bulk in-memory work |
| `ws["A1"]`, `ws.cell()`, `ws.title`, `ws.merge_cells()`, `ws.insert_rows()`, `wb.create_sheet()`, styles, dimensions ... | inline on the event loop | in-memory operations |
| `ws["A1"]`, `ws.cell()` on a **read-only** sheet | raises `BlockingCallError` | openpyxl would parse XML on the event loop; use `fetch` / `read_rows` / `run` |

Keep `chunk_size` at 256 or more (each chunk is one thread round trip); for large exports use
`Workbook(write_only=True)` with `append_rows`. The internals are explained in [docs/design.md](docs/design.md).

## Rules of the road

- **A thread owns the workbook while an `await` is pending.** Synchronous access meanwhile
  (`ws.cell(...)`, `wb.sheetnames`, even `repr(wb)`) raises `WorkbookBusyError`; await first or
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

Every other public member is forwarded with its typed openpyxl signature; the full table and the
typing, sub-module and subclass rules are in [docs/compatibility.md](docs/compatibility.md).

## Development

```sh
uv sync && uv run pytest -q                                 # both backends, random order
uv run ruff check . && uv run ruff format --check .
uv run ty check src tests && uv run mypy
uv run python tools/gen_proxies.py && uv run ruff format src/aioopenpyxl/   # after an openpyxl / stubs bump
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for the workflow and [CHANGELOG.md](CHANGELOG.md) for history.

## Acknowledgements

All the real work is done by [openpyxl](https://openpyxl.readthedocs.io/) (MIT), which this
package depends on rather than vendors. The typed forwarders and `.pyi` shims are generated from
the [types-openpyxl](https://pypi.org/project/types-openpyxl/) stubs maintained in
[typeshed](https://github.com/python/typeshed) (Apache-2.0); concurrency primitives come from
[anyio](https://anyio.readthedocs.io/) (MIT).

## License

[MIT](LICENSE). openpyxl, typeshed and anyio remain under their own licenses.
