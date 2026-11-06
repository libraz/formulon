"""Resolve a table, pivot, CF-rule or protected-range reference to its index.

A reference is a 0-based index or a name: a table's display name, a pivot's
name, a CF rule's id, a protected range's name. A name matches ignoring case
under Unicode simple case folding, as sheet names do, and the first entry
carrying it wins; a numeric-looking string is still a name.
The collection is listed afresh on every call, so resolving is O(n).
"""

from __future__ import annotations

import operator
from typing import Callable, Iterable, Union

from ._c import FormulonError
from ._status import StatusCode


def _fold(text: str) -> str:
    """``text`` under Unicode simple case folding (one code point per code point)."""
    out = []
    for char in text:
        folded = char.casefold()
        if len(folded) != 1:
            # A full folding expands; the simple one is the one-code-point lowercase, else none.
            lowered = char.lower()
            folded = lowered if len(lowered) == 1 else char
        out.append(folded)
    return "".join(out)


def resolve_ref(ref: Union[int, str], names: Callable[[], Iterable[str]], what: str) -> int:
    """Return the index ``ref`` addresses.

    Args:
      ref: an index, returned unchanged so the caller's index checks apply,
        or a name looked up in ``names()``, ignoring case.
      names: lists the collection's names in index order; called only for a
        ``str`` reference.
      what: the collection's noun for the error message (``"table"``).

    Raises:
      TypeError: when ``ref`` is a ``bool`` or neither an integer nor a ``str``.
      NotFoundError: when no entry carries the name ``ref``.
    """
    if isinstance(ref, str):
        wanted = _fold(ref)
        for index, name in enumerate(names()):
            if _fold(name) == wanted:
                return index
        raise FormulonError(StatusCode.NOT_FOUND, _diagnostic_override=(f"no {what} is named {ref!r}", ""))
    if isinstance(ref, bool):
        raise TypeError(f"formulon: {what} must be an int or a str, got bool")
    try:
        return operator.index(ref)
    except TypeError:
        raise TypeError(f"formulon: {what} must be an int or a str, got {type(ref).__name__}") from None
