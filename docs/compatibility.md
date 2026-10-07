# openpyxl compatibility

aioopenpyxl aims to be a drop-in asynchronous replacement for openpyxl's `Workbook`, worksheet
and chartsheet objects. This page lists what is forwarded unchanged, how the pieces fit
together, and every deliberate difference. The design rationale is in [design.md](design.md).

## Every public member is forwarded

`Workbook` (alias `AsyncWorkbook`), `AsyncWorksheet` (covering `Worksheet`,
`ReadOnlyWorksheet`, `WriteOnlyWorksheet`) and `AsyncChartsheet` expose **every public member**
of their openpyxl counterparts with the original typed signatures. The forwarders are generated
from the `types-openpyxl` stubs (`tools/gen_proxies.py`), which is why `types-openpyxl` is a
runtime dependency; tests assert that nothing is missing, both on the classes and on live
instances (a regular, a read-only and a write-only sheet, a chartsheet and a workbook).

Members returning sheets return wrappers (`wb.active`, `wb["Data"]`, `wb.create_sheet()`,
`wb.create_chartsheet()`, `wb.worksheets`, `wb.chartsheets`, `wb.copy_worksheet()`), and
`ws.parent` is the `Workbook`. Sheet wrappers are cached per workbook, so
`wb["Data"] is wb["Data"]`. `AsyncChartsheet.from_tree(node)` returns a wrapped
`AsyncChartsheet | None`.

## The `Workbook` class and `wrap`

`Workbook` is a real class with openpyxl's constructor signature
(`Workbook(write_only=False, iso_dates=False, *, limiter=None)`), so annotations, `isinstance`
checks and subclassing work as with `openpyxl.Workbook`. Construction is purely in-memory, so
it is a plain call rather than a coroutine. Passing an existing openpyxl workbook to the
constructor raises `TypeError` instead of silently creating an empty workbook; use
`Workbook.wrap(raw, limiter=...)`. `repr(wb)` is `<Workbook sheets=[...]>` (or
`<Workbook busy>` while a thread owns it). `AsyncWorkbook` is kept as a backwards-compatible
alias (`AsyncWorkbook is Workbook`).

## Module constants

`aioopenpyxl.open` mirrors `openpyxl.open`; `aioopenpyxl.LXML`, `DEBUG`, `DEFUSEDXML`, `NUMPY`,
`__author__` etc. are re-exported explicitly (and listed in `__all__`), so they type-check and
can be read outside an event loop. A runtime-only `__getattr__` still forwards anything openpyxl
adds later.

## Sub-module aliases and `.pyi` shims

openpyxl's sub-packages are aliased, so `import aioopenpyxl as openpyxl` keeps working for
helpers: `from aioopenpyxl.styles import Font`, `aioopenpyxl.utils.get_column_letter`,
`import aioopenpyxl.worksheet.table` all resolve to the real openpyxl modules. The package ships
generated `.pyi` shims (`aioopenpyxl/styles/__init__.pyi` is `from openpyxl.styles import *`,
and so on for every openpyxl module), so those imports also resolve for ty, mypy and pyright.

**The aliases are openpyxl's modules unchanged:** `aioopenpyxl.workbook.Workbook` and
`aioopenpyxl.reader.excel.load_workbook` are the synchronous openpyxl originals; only the
top-level `aioopenpyxl.Workbook` / `aioopenpyxl.load_workbook` / `aioopenpyxl.open` are
asynchronous.

## Deprecated openpyxl API

openpyxl's deprecated methods (`wb.get_sheet_by_name`, `wb.get_sheet_names`, `wb.remove_sheet`,
`wb.get_index`, `wb.create_named_range`, ...) are forwarded and keep emitting their
`DeprecationWarning`. Generated forwarders of `@deprecated` stub members carry the same
decorator, so type checkers report them as well.

## Typing

Typing follows the stubs: `ws.max_row` / `ws.max_column` are `int | None`, and only the
generated, typed members exist statically (a typo on `ws.tilte` is a type error, not `Any`).
The `__getattr__` fallbacks are hidden from type checkers for that reason. `read_rows` and
`iter_rows` / `iter_cols` have overloads for literal and `bool` `values_only` (keyword or
positional); rows of read-only sheets are typed as containing `ReadOnlyCell` / `EmptyCell` next
to `Cell` / `MergedCell`. Note that `read_rows` defaults to `values_only=True` (a bulk read is
almost always after the values) where openpyxl's `iter_rows` defaults to `False`.
Descriptor-typed stub attributes (`Typed[PageMargins, ...]`, `Set[_VisibilityType]`, `Alias`) are
exposed as the value types they produce, exactly as on the raw openpyxl class. The package ships
a `py.typed` marker.

### Worksheet kinds

`Workbook`, `AsyncWorksheet` and `LoadWorkbookContextManager` are generic in the raw worksheet
type, which is fixed by where the workbook comes from:

| Factory | Type |
| --- | --- |
| `Workbook()`, `Workbook(write_only=False)` | `Workbook[Worksheet]` |
| `Workbook(write_only=True)` | `Workbook[WriteOnlyWorksheet]` |
| `load_workbook(...)`, `load_workbook_bytes(...)` | `Workbook[Worksheet]` |
| `load_workbook(..., read_only=True)` | `Workbook[ReadOnlyWorksheet]` |
| a plain `bool` for `write_only` / `read_only` | the union of the two possible kinds |
| `Workbook.wrap(raw)` | plain `Workbook` (a raw workbook does not say); `Workbook[Worksheet].wrap(raw)` when you know |

`wb.active`, `wb["Sheet"]`, `wb.worksheets`, `wb.create_sheet()` and iteration hand out
`AsyncWorksheet[...]` of that type, `ws.wrapped` is that type and `ws.run(fn)` types `fn` as
`Callable[[Worksheet], R]` (or `ReadOnlyWorksheet`, `WriteOnlyWorksheet`), so a helper written
against the concrete openpyxl class is accepted without a cast. `wb.run(fn)` takes
`Callable[[openpyxl.Workbook], R]` as before.

The parameter is covariant and defaults to the union of the three classes: a plain `Workbook` or
`AsyncWorksheet` annotation means "any kind" and accepts every precise one, so existing
annotations keep working (`wb.worksheets` is typed as a `Sequence` for the same reason). A
subclass declared as `class MyWorkbook(Workbook)` is a `Workbook` of any kind; declare it as
`class MyWorkbook(Workbook[Worksheet])` to keep the precise type.

## Private attributes

Private openpyxl attributes are not forwarded; reach them through the raw object
(`ws.wrapped._images`, `wb.wrapped._archive`). `wb.wrapped` / `ws.wrapped` return the raw
objects without the busy check.

## Attributes on subclasses

Subclassing a wrapper is supported, and a subclass may keep attributes of its own. An
assignment `obj.name = value` is resolved in this order:

1. Private names (`_x`) and anything defined on the wrapper's class, including slots, generated
   properties and class attributes of a subclass, are set on the wrapper itself. Generated
   properties forward to the raw object through their own setters, so `ws.title = "X"` still
   reaches openpyxl, while a subclass counter such as `rows_written = 0` is updated on the
   instance (`self.rows_written += 1` works as in any Python class).
2. Otherwise, an attribute that already exists on the raw openpyxl object is forwarded to it
   (and raises `WorkbookBusyError` while a thread owns the workbook, like any other synchronous
   access).
3. Otherwise the name is new on both sides: it is stored on the wrapper when the instance has a
   `__dict__` (a subclass without `__slots__`), and forwarded to the raw object when it does not
   (the base wrappers only define slots), so unknown names can still be put on the openpyxl
   object.

## Differences, all deliberate

| openpyxl | aioopenpyxl |
| --- | --- |
| `Workbook(write_only=..., iso_dates=...)` | `Workbook(write_only=..., iso_dates=..., limiter=...)` (same signature); wrap an existing workbook with `Workbook.wrap(wb)` |
| `load_workbook(...)` | `await load_workbook(...)` or `async with load_workbook(...) as wb` |
| `load_workbook(io.BytesIO(data))` | `await load_workbook_bytes(data)` |
| `wb.save(f)`, `wb.close()` | `await wb.save(f)`, `await wb.to_bytes()`, `await wb.close()` |
| `wb.close()` at the end of the workbook's life | `async with wb:` / `async with load_workbook(...) as wb:` closes the workbook and, for `keep_vba=True`, the in-memory `vba_archive` copy that openpyxl leaves open (its finaliser can otherwise fail at interpreter shutdown); `await wb.close()` alone is exactly `wb.close()` |
| `for row in ws.iter_rows()` / `ws.rows` / `ws.values` / `ws.columns` / `for row in ws` | `async for` versions (chunked); `iter_rows` / `iter_cols` gain `chunk_size=` and `prefetch=` |
| `for row in ws` (synchronous) | `TypeError`: there is no synchronous iteration; use `async for row in ws` or `ws.wrapped` |
| `ws.append(row)` | `await ws.append(row)`; plus `await ws.append_rows(rows)` |
| `ws.calculate_dimension()` | `await ws.calculate_dimension(force=False)` |
| `WriteOnlyWorksheet.close()` | `await ws.close()` |
| `ws["A1"]`, `ws.cell(1, 1)` on a read-only sheet | raise `BlockingCallError`; use `await ws.fetch("A1")`, `await ws.read_rows(...)`, `await ws.run(...)` |
| – | `await ws.read_rows(...)` (defaults to `values_only=True`, unlike `iter_rows`), `await ws.fetch(key)`, `await x.run(fn)`, `x.wrapped`, `ws.is_read_only`, `ws.is_write_only` |
