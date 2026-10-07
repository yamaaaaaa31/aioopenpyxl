"""Shared proxy machinery for the async wrappers."""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any, TypeVar

from ._errors import WorkbookBusyError
from ._executor import Runner

__all__ = ["GuardedProxy"]

R = TypeVar("R")


class GuardedProxy:
    """Base for every wrapper: guarded access to the raw openpyxl object.

    Concrete subclasses store the wrapped object in ``_obj`` and their
    :class:`Runner` in ``_runner``.  All synchronous forwarders (generated or
    hand-written) go through :attr:`_raw`, which refuses access while the
    workbook's runner is executing something in a worker thread.
    Hand-written coroutines pass ``_obj`` into the runner directly.
    """

    __slots__ = ()

    _obj: Any
    _runner: Runner

    @property
    def _raw(self) -> Any:
        if self._runner.busy:
            raise WorkbookBusyError(
                f"{type(self).__name__} is being used by a worker thread; "
                "await the pending operation before touching it from the event loop, "
                "or do this work inside .run(...)"
            )
        return self._obj

    @property
    def wrapped(self) -> Any:
        """The raw openpyxl object, without the busy check (escape hatch)."""
        return self._obj

    async def run(self, func: Callable[[Any], R]) -> R:
        """Run ``func(raw_object)`` in the workbook's worker thread and return its result.

        This is the right tool for bulk in-memory work (thousands of ``cell()``
        calls, styling loops ...): one thread hop, no event-loop stalls, and
        the per-workbook runner guarantees nothing else touches the workbook
        meanwhile.
        """
        return await self._runner(func, self._obj)

    if not TYPE_CHECKING:
        # Runtime fallback for members the generated surface does not list.
        # Hidden from type checkers on purpose: statically only the generated,
        # typed members exist, so a typo is an error instead of ``Any``.
        def __getattr__(self, name: str) -> Any:
            # Reached only for attributes the proxy does not define itself.
            if name.startswith("_"):
                raise AttributeError(name)
            return getattr(self._raw, name)

    def __setattr__(self, name: str, value: Any) -> None:
        """Decide whether an assignment lands on the wrapper or on the raw object.

        The rules, in order:

        1. Private names (``_x``) and anything defined on the wrapper's class
           (slots, generated properties, descriptors, class attributes of a
           subclass) are set on the wrapper itself with ``object.__setattr__``.
           Generated properties forward to the raw object through their own
           setters, so ``ws.title = "X"`` still reaches openpyxl; a subclass
           counter such as ``rows_written = 0`` is updated on the instance.
        2. Otherwise, an attribute that already exists on the raw openpyxl
           object is forwarded to it (``wb.iso_dates = True`` while a thread owns
           the workbook raises :class:`WorkbookBusyError` like any other
           synchronous access).
        3. Otherwise the attribute is new on both sides: it is stored on the
           wrapper when the instance has a ``__dict__`` (a subclass without
           ``__slots__``), and forwarded to the raw object when it does not (the
           base wrappers only define slots), so unknown names can still be put
           on the openpyxl object as before.
        """
        if name.startswith("_") or hasattr(type(self), name):
            object.__setattr__(self, name, value)
            return
        try:
            raw = self._raw
        except AttributeError:
            # The wrapper is not initialised yet (a subclass assigning before
            # ``super().__init__()``): there is no raw object to forward to.
            raw = None
        if raw is not None and hasattr(raw, name):
            setattr(raw, name, value)
            return
        try:
            object.__setattr__(self, name, value)
        except AttributeError:
            if raw is None:
                raise
            setattr(raw, name, value)
