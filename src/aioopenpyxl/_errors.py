"""Exceptions raised by aioopenpyxl."""

__all__ = ["BlockingCallError", "WorkbookBusyError"]


class WorkbookBusyError(RuntimeError):
    """Raised when a workbook is touched from the event loop while a worker thread owns it.

    openpyxl objects are not thread-safe.  While ``await wb.save()``,
    ``await ws.append_rows()``, ``await wb.run()`` or a chunk fetch of
    ``async for row in ws.iter_rows()`` is executing in a worker thread, every
    synchronous access through the async wrappers (``ws.cell(...)``,
    ``wb.sheetnames`` ...) raises this error instead of silently racing with the
    thread.  Await the pending operation first, or move the work into
    ``run(...)`` so it executes in the same thread hop.
    """


class BlockingCallError(RuntimeError):
    """Raised when a synchronous wrapper call would block the event loop.

    Currently raised by ``ws[key]`` on read-only worksheets, where openpyxl
    parses XML synchronously.  Use ``await ws.fetch(key)``,
    ``await ws.read_rows(...)`` or ``async for`` instead, or go through
    ``ws.wrapped[key]`` if you really want the blocking call.
    """
