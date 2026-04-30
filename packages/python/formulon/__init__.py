# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
"""Formulon -- Excel 365 calculation engine, Python binding.

Public API:
  * :func:`eval_formula` -- one-shot formula evaluation.
  * :func:`library_version` -- query the underlying ``libformulon`` build.
  * :class:`Workbook` -- full workbook lifecycle.
  * :class:`Value`, :class:`ValueKind` -- cell value POD.
  * :class:`FormulonError` -- host-side failure exception.

Cell-level Excel errors (``#DIV/0!`` and friends) surface as ``Value``
instances with ``kind == ValueKind.ERROR``; only host-side failures
(NULL handle, parse errors during ``Workbook.load``, OOM) raise
:class:`FormulonError`.
"""

from __future__ import annotations

from ._c import LIB, ValueKind, decode_cstr
from .workbook import (
    Cell,
    DefinedName,
    FormulonError,
    PassthroughPart,
    Table,
    Value,
    Workbook,
)

# IMPORTANT: keep this in sync with the [project] version in
# packages/python/pyproject.toml. There is no portable way to read the
# wheel metadata at runtime without `importlib.metadata`, which is fine
# for installed packages but breaks for source-tree imports.
__version__ = "0.1.0"

__all__ = [
    "Cell",
    "DefinedName",
    "FormulonError",
    "PassthroughPart",
    "Table",
    "Value",
    "ValueKind",
    "Workbook",
    "__version__",
    "eval_formula",
    "library_version",
    "version_string",
]


def library_version() -> str:
    """Return the version string baked into the loaded ``libformulon``.

    Returns:
      The result of ``fm_version_string()``. Always non-empty.
    """
    return decode_cstr(LIB.fm_version_string())


# Backward-compat alias mirroring the npm binding's ``versionString`` name.
version_string = library_version


def eval_formula(formula: str) -> Value:
    """Evaluate a single formula against a fresh, default workbook.

    The formula is written to ``Sheet1!A1``, the workbook is recalculated,
    and the resulting cell value is returned.

    Cell-level Excel errors (``#DIV/0!``, ``#VALUE!``, etc.) surface as a
    :class:`Value` with ``kind == ValueKind.ERROR``; only host-side
    failures (parser crash, OOM, NULL handle) raise :class:`FormulonError`.

    Args:
      formula: the formula text, with or without a leading ``=``.

    Returns:
      The :class:`Value` cached at ``Sheet1!A1`` after recalc.
    """
    with Workbook.create_default() as wb:
        wb.set_formula(0, 0, 0, formula)
        wb.recalc()
        return wb.get_value(0, 0, 0)
