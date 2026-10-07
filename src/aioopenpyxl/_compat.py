"""Make ``aioopenpyxl`` a drop-in module path for openpyxl's sub-packages.

After :func:`install_module_aliases` runs, ``from aioopenpyxl.styles import Font``
or ``import aioopenpyxl.utils.cell`` resolve to the very same module objects as
their ``openpyxl.*`` counterparts, so ``import aioopenpyxl as openpyxl`` works
for everything that is not a Workbook/Worksheet entry point.

The aliases are *openpyxl's modules, unchanged*.  In particular
``aioopenpyxl.workbook.Workbook`` and ``aioopenpyxl.reader.excel.load_workbook``
are the synchronous openpyxl originals; only the top-level
``aioopenpyxl.Workbook`` / ``aioopenpyxl.load_workbook`` / ``aioopenpyxl.open``
are the asynchronous wrappers.  Code that imports the entry points from a
sub-module path will therefore keep running openpyxl synchronously.

How the aliases interact with the import system: ``import aioopenpyxl.styles``
first imports the parent package ``aioopenpyxl`` (whose ``__init__`` calls
:func:`install_module_aliases`), and only then looks the submodule up -- in
``sys.modules`` before any path finder.  The alias is therefore found even if
``src/aioopenpyxl/styles/`` exists on disk holding only a ``.pyi`` type shim
(no ``.py``), which would otherwise be importable as an empty namespace
package.  The flip side: if an alias is *missing* (the openpyxl module failed
to import, e.g. an optional dependency is absent), such a shim directory
resolves to an empty namespace package rather than raising ``ImportError``.
"""

from __future__ import annotations

import importlib
import pkgutil
import sys
from types import ModuleType

__all__ = ["install_module_aliases"]


def install_module_aliases(package: ModuleType) -> list[str]:
    """Register every ``openpyxl.*`` sub-module under ``<package>.*`` as an alias.

    For each importable ``openpyxl.<dotted>`` module, ``sys.modules`` gains the
    key ``<package.__name__>.<dotted>`` pointing at the same module object, and
    top-level sub-packages (``styles``, ``utils`` ...) are also set as
    attributes on ``package`` so ``aioopenpyxl.styles`` works after a bare
    ``import aioopenpyxl``.

    Names already present in ``sys.modules`` are left untouched, so modules the
    package defines itself (``aioopenpyxl._workbook`` ...) are never overridden.
    Because of that, the function is idempotent: a second call (or a call after
    another import already populated ``sys.modules``) registers nothing new and
    returns an empty list.  openpyxl modules that fail to import (optional
    dependencies such as numpy, pandas or PIL missing) are skipped.

    Returns the list of alias names that were registered by this call.
    """
    import openpyxl

    prefix = package.__name__
    registered: list[str] = []
    for info in pkgutil.walk_packages(openpyxl.__path__, "openpyxl."):
        try:
            module = importlib.import_module(info.name)
        except Exception:  # optional deps (numpy, pandas, PIL ...) may be missing
            continue
        alias = prefix + info.name[len("openpyxl") :]
        if alias in sys.modules:
            continue
        sys.modules[alias] = module
        registered.append(alias)
        top = alias[len(prefix) + 1 :]
        if "." not in top and not hasattr(package, top):
            setattr(package, top, module)
    return registered
