"""Static typing checks (run by ``ty``, nothing here executes at test time)."""

from __future__ import annotations

from collections.abc import AsyncGenerator, Sequence
from typing import TYPE_CHECKING

import aioopenpyxl

if TYPE_CHECKING:
    from typing import Any

    import openpyxl
    from anyio import CapacityLimiter
    from openpyxl import _VisibilityType
    from openpyxl.cell import _CellGetValue, _CellOrMergedCell
    from openpyxl.cell.cell import Cell
    from openpyxl.cell.read_only import EmptyCell, ReadOnlyCell
    from openpyxl.chartsheet.views import ChartsheetViewList
    from openpyxl.packaging.custom import CustomPropertyList
    from openpyxl.worksheet._read_only import ReadOnlyWorksheet
    from openpyxl.worksheet._write_only import WriteOnlyWorksheet
    from openpyxl.worksheet.page import PageMargins
    from openpyxl.worksheet.worksheet import Worksheet
    from typing_extensions import assert_type  # 3.10-compatible for the type checker

    RawWorksheet = Worksheet | ReadOnlyWorksheet | WriteOnlyWorksheet

    # Read-only sheets yield ReadOnlyCell / EmptyCell, regular ones Cell / MergedCell.
    CellRow = tuple[_CellOrMergedCell | ReadOnlyCell | EmptyCell, ...]
    ValueRow = tuple[_CellGetValue, ...]


async def _overloads(ws: aioopenpyxl.AsyncWorksheet, values_only: bool) -> None:
    assert_type(ws.iter_rows(values_only=True), AsyncGenerator[ValueRow, None])
    assert_type(ws.iter_rows(), AsyncGenerator[CellRow, None])
    assert_type(ws.iter_rows(1, 2, 1, 1, False), AsyncGenerator[CellRow, None])
    # All five arguments positional, like openpyxl's own overloads.
    assert_type(ws.iter_rows(1, 2, 1, 2, True), AsyncGenerator[ValueRow, None])
    assert_type(ws.iter_rows(1, 2, 1, 2, True, chunk_size=8), AsyncGenerator[ValueRow, None])
    assert_type(ws.iter_rows(None, None, None, None, True), AsyncGenerator[ValueRow, None])
    assert_type(
        ws.iter_rows(values_only=values_only),
        AsyncGenerator[CellRow, None] | AsyncGenerator[ValueRow, None],
    )
    assert_type(ws.iter_cols(values_only=True, chunk_size=8), AsyncGenerator[ValueRow, None])
    assert_type(ws.iter_cols(prefetch=2), AsyncGenerator[CellRow, None])
    assert_type(ws.iter_cols(1, 2, 1, 1, False), AsyncGenerator[CellRow, None])
    assert_type(ws.iter_cols(1, 2, 1, 2, True), AsyncGenerator[ValueRow, None])
    assert_type(ws.iter_cols(1, 2, 1, 2, True, prefetch=1), AsyncGenerator[ValueRow, None])
    assert_type(
        ws.iter_cols(values_only=values_only),
        AsyncGenerator[CellRow, None] | AsyncGenerator[ValueRow, None],
    )
    assert_type(ws.values, AsyncGenerator[ValueRow, None])
    assert_type(ws.rows, AsyncGenerator[CellRow, None])
    assert_type(ws.columns, AsyncGenerator[CellRow, None])
    assert_type(ws.__aiter__(), AsyncGenerator[CellRow, None])
    assert_type(await ws.read_rows(), list[ValueRow])
    assert_type(await ws.read_rows(values_only=False), list[CellRow])
    assert_type(await ws.read_rows(1, 2, 1, 1, False), list[CellRow])
    assert_type(await ws.read_rows(1, 2, 1, 1, True), list[ValueRow])
    assert_type(await ws.read_rows(values_only=values_only), list[ValueRow] | list[CellRow])
    assert_type(ws.max_row, int | None)
    assert_type(ws.max_column, int | None)
    assert_type(ws.title, str)
    assert_type(ws.is_read_only, bool)
    assert_type(await ws.run(lambda raw: raw.title), str)
    assert_type(await ws.append_rows([[1, 2]]), int)
    assert_type(await ws.calculate_dimension(), str)
    assert_type(ws.cell(1, 1), _CellOrMergedCell)
    assert_type(ws.cell(1, 1, "x"), Cell)
    parent = ws.parent
    assert_type(parent, aioopenpyxl.Workbook | None)


async def _workbook(wb: aioopenpyxl.Workbook, raw: openpyxl.Workbook) -> None:
    assert_type(wb.active, aioopenpyxl.AsyncWorksheet | None)
    assert_type(wb["x"], aioopenpyxl.AsyncWorksheet)
    assert_type(wb.create_sheet("x"), aioopenpyxl.AsyncWorksheet)
    assert_type(wb.create_chartsheet(), aioopenpyxl.AsyncChartsheet)
    assert_type(wb.worksheets, Sequence[aioopenpyxl.AsyncWorksheet])
    assert_type(wb.chartsheets, list[aioopenpyxl.AsyncChartsheet])
    assert_type(await wb.to_bytes(), bytes)
    assert_type(wb.read_only, bool)
    assert_type(wb.wrapped, openpyxl.Workbook)
    assert_type(wb.limiter, CapacityLimiter | None)
    assert_type(wb.custom_doc_props, CustomPropertyList[Any])
    assert_type(await wb.run(lambda raw: raw.sheetnames), list[str])
    # A raw workbook does not say which kind of sheets it holds.
    assert_type(aioopenpyxl.Workbook.wrap(raw), aioopenpyxl.Workbook)
    assert_type(aioopenpyxl.AsyncWorkbook.wrap(raw), aioopenpyxl.Workbook)
    assert_type(aioopenpyxl.Workbook.wrap(raw), aioopenpyxl.Workbook[RawWorksheet])
    # Deprecated forwarders are still typed.
    assert_type(wb.get_sheet_by_name("x"), aioopenpyxl.AsyncWorksheet)


async def _sheet_kinds(raw: openpyxl.Workbook, flag: bool) -> None:
    """The raw worksheet type follows from where the workbook comes from."""
    assert_type(aioopenpyxl.Workbook(), aioopenpyxl.Workbook[Worksheet])
    assert_type(aioopenpyxl.Workbook(write_only=False), aioopenpyxl.Workbook[Worksheet])
    assert_type(aioopenpyxl.Workbook(write_only=True), aioopenpyxl.Workbook[WriteOnlyWorksheet])
    assert_type(
        aioopenpyxl.Workbook(write_only=flag), aioopenpyxl.Workbook[Worksheet | WriteOnlyWorksheet]
    )
    assert_type(aioopenpyxl.Workbook[Worksheet].wrap(raw), aioopenpyxl.Workbook[Worksheet])
    assert_type(await aioopenpyxl.load_workbook("x.xlsx"), aioopenpyxl.Workbook[Worksheet])
    assert_type(
        await aioopenpyxl.load_workbook("x.xlsx", read_only=True),
        aioopenpyxl.Workbook[ReadOnlyWorksheet],
    )
    assert_type(
        await aioopenpyxl.load_workbook("x.xlsx", flag),
        aioopenpyxl.Workbook[Worksheet | ReadOnlyWorksheet],
    )
    async with aioopenpyxl.load_workbook_bytes(b"") as loaded:
        assert_type(loaded, aioopenpyxl.Workbook[Worksheet])
    async with aioopenpyxl.load_workbook_bytes(b"", read_only=True) as read_only:
        assert_type(read_only, aioopenpyxl.Workbook[ReadOnlyWorksheet])
        assert_type(read_only.active, aioopenpyxl.AsyncWorksheet[ReadOnlyWorksheet] | None)
        assert_type(read_only["x"].wrapped, ReadOnlyWorksheet)

    wb = aioopenpyxl.Workbook()
    ws = wb["x"]
    assert_type(ws, aioopenpyxl.AsyncWorksheet[Worksheet])
    assert_type(wb.active, aioopenpyxl.AsyncWorksheet[Worksheet] | None)
    assert_type(wb.worksheets, Sequence[aioopenpyxl.AsyncWorksheet[Worksheet]])
    assert_type(wb.create_sheet(), aioopenpyxl.AsyncWorksheet[Worksheet])
    assert_type(wb.copy_worksheet(ws), aioopenpyxl.AsyncWorksheet[Worksheet])
    assert_type(ws.wrapped, Worksheet)
    assert_type(ws.parent, aioopenpyxl.Workbook[Worksheet] | None)
    for each in wb:
        assert_type(each, aioopenpyxl.AsyncWorksheet[Worksheet])
    # ``run`` accepts a callback written for the precise worksheet class.
    assert_type(await ws.run(lambda sheet: sheet.cell(1, 1)), _CellOrMergedCell)

    def helper(sheet: Worksheet) -> int:
        return sheet.max_row

    assert_type(await ws.run(helper), int)

    write_only = aioopenpyxl.Workbook(write_only=True)
    assert_type(write_only.create_sheet().wrapped, WriteOnlyWorksheet)

    # Covariance: precise workbooks and sheets go wherever the plain ones are accepted.
    plain_wb: aioopenpyxl.Workbook = wb
    plain_ws: aioopenpyxl.AsyncWorksheet = ws
    assert_type(await plain_ws.run(lambda sheet: sheet.title), str)
    del plain_wb

    # A subclass of the plain ``Workbook`` is constructed as usual.
    class MyWorkbook(aioopenpyxl.Workbook):
        pass

    assert_type(MyWorkbook(), MyWorkbook)
    assert_type(MyWorkbook(write_only=True), MyWorkbook)
    assert_type(MyWorkbook.wrap(raw), MyWorkbook)
    assert_type(MyWorkbook().active, aioopenpyxl.AsyncWorksheet | None)

    class MyRegularWorkbook(aioopenpyxl.Workbook[Worksheet]):
        pass

    assert_type(MyRegularWorkbook().active, aioopenpyxl.AsyncWorksheet[Worksheet] | None)


def _chartsheet(cs: aioopenpyxl.AsyncChartsheet, node: Any) -> None:
    # Descriptor-typed stub attributes are exposed as the value types they
    # produce, exactly as on the raw openpyxl class.
    assert_type(cs.pageMargins, PageMargins | None)
    cs.pageMargins = PageMargins(left=1.0)
    cs.pageMargins = None
    margins = cs.pageMargins
    if margins is not None:
        assert_type(margins.left, float)
    assert_type(cs.sheet_state, _VisibilityType)
    cs.sheet_state = "hidden"
    assert_type(cs.sheetViews, ChartsheetViewList)  # Typed[..., Literal[False]]
    assert_type(cs.title, str)
    assert_type(aioopenpyxl.AsyncChartsheet.from_tree(node), aioopenpyxl.AsyncChartsheet | None)


def _module_constants() -> None:
    # Explicit re-exports: resolvable without the runtime-only ``__getattr__``.
    assert_type(aioopenpyxl.LXML, bool)
    assert_type(aioopenpyxl.DEBUG, bool)
    assert_type(aioopenpyxl.__author__, str)


def _shims() -> None:
    # The generated ``.pyi`` shims resolve the aliased sub-packages statically.
    from aioopenpyxl.styles import Font
    from aioopenpyxl.utils import get_column_letter

    assert_type(get_column_letter(1), str)
    assert_type(Font(), Font)


def test_module_is_importable() -> None:
    assert aioopenpyxl.AsyncWorksheet is not None
