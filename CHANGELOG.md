# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.1] - 2026-10-08

### Added

- `Workbook`, `AsyncWorksheet` and `LoadWorkbookContextManager` are generic in the raw worksheet
  type, fixed by where the workbook comes from: `Workbook()` is a `Workbook[Worksheet]`,
  `Workbook(write_only=True)` a `Workbook[WriteOnlyWorksheet]`, `load_workbook(...,
  read_only=True)` yields a `Workbook[ReadOnlyWorksheet]`. Sheets, `wrapped` and the `run(fn)`
  callback are typed accordingly, so a helper written for `Worksheet` no longer needs a cast.
  Type-only and backwards compatible: the parameter is covariant and a plain `Workbook` /
  `AsyncWorksheet` annotation still means any kind (`wb.worksheets` is now typed as a
  `Sequence`). `Workbook.wrap` and `async with wb` return `Self`.

### Fixed

- Leaving `async with wb:` / `async with load_workbook(...)` for a `keep_vba=True` workbook also
  closes the in-memory `vba_archive` copy that openpyxl leaves open; its finaliser could
  otherwise fail at interpreter shutdown with `ValueError: I/O operation on closed file`.
  `await wb.close()` is unchanged and still matches `openpyxl.Workbook.close`.

### Documentation

- `read_rows` defaults to `values_only=True` (openpyxl's `iter_rows` defaults to `False`); this is
  now stated in its docstring and in the compatibility table.

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

[Unreleased]: https://github.com/yamaaaaaa31/aioopenpyxl/compare/v0.1.1...HEAD
[0.1.1]: https://github.com/yamaaaaaa31/aioopenpyxl/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/yamaaaaaa31/aioopenpyxl/releases/tag/v0.1.0
