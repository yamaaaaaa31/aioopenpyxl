# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-10-08

Initial release.

- Asynchronous wrapper around openpyxl built on anyio; asyncio and trio are both supported and
  tested.
- `Workbook` (alias `AsyncWorkbook`), `AsyncWorksheet` and `AsyncChartsheet` forward **every
  public member** of their openpyxl counterparts with the original typed signatures, generated
  from the `types-openpyxl` stubs.
- `Workbook(...)` with openpyxl's constructor signature, `Workbook.wrap(raw)`,
  `await load_workbook(...)` / `async with load_workbook(...)`, `await load_workbook_bytes(data)`,
  `await wb.save(...)`, `await wb.to_bytes()`, `await wb.close()`.
- Chunked `async for` streaming of rows and columns (`iter_rows`, `iter_cols`, `rows`, `values`,
  `columns`, `async for row in ws`) with `chunk_size=` and `prefetch=`; coarse-grained
  `await ws.read_rows(...)`, `await ws.fetch(key)`; bulk `await ws.append_rows(rows)`;
  `await x.run(fn)` for arbitrary in-memory work in one thread hop.
- One lock per workbook: `WorkbookBusyError` on synchronous access while a worker thread owns
  the workbook, `BlockingCallError` for `ws["A1"]` / `ws.cell()` on read-only sheets.
- Cancellation-safe: a cancelled call (anyio or native asyncio) waits for its thread, clean-up
  always closes the workbook and the `prefetch` producer, and no zip archive is left open.
- `limiter=` (`anyio.CapacityLimiter`) on `load_workbook`, `load_workbook_bytes`, `Workbook` and
  `Workbook.wrap` to bound worker threads, producers included.
- openpyxl sub-packages aliased (`aioopenpyxl.styles`, `aioopenpyxl.utils`, ...) with generated
  `.pyi` shims, module constants re-exported, deprecated openpyxl methods forwarded with their
  warnings; `py.typed`.
- Subclasses of the wrappers can keep attributes of their own.
- Supports CPython 3.10 to 3.14, including the free-threaded 3.14t build.

[Unreleased]: https://github.com/yamaaaaaa31/aioopenpyxl/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/yamaaaaaa31/aioopenpyxl/releases/tag/v0.1.0
