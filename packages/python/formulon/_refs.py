"""Resolve a table, pivot, CF-rule or protected-range reference to its index.

A reference is a 0-based index, a name or, for a table or a pivot, its
handle. A name is a table's display name, a pivot's name, a CF rule's id or
a protected range's name; it matches ignoring case under Unicode simple case
folding, as sheet names do, and the first entry carrying it wins; a
numeric-looking string is still a name. The collection is listed afresh on
every call, so resolving is O(n). A handle resolves to its object's current
index.
"""

from __future__ import annotations

import operator
from typing import Any, Callable, Iterable, Optional, Union

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


def _handle_index(ref: Any, owner: Any, what: str) -> Optional[int]:
    """The current index of ``ref`` when it is the handle ``owner`` resolves, else ``None``.

    A :class:`Sheet` owner resolves its :class:`Pivot` handles, a
    :class:`Workbook` owner its :class:`Table` handles.

    Raises:
      InvalidArgumentError: when ``ref`` is such a handle of another sheet or workbook.
    """
    from ._pivot import Pivot
    from ._sheet import Sheet
    from ._table import Table

    if isinstance(owner, Sheet):
        if not isinstance(ref, Pivot):
            return None
        parent, noun, place = ref.sheet, "Pivot", "sheet"
    else:
        if not isinstance(ref, Table):
            return None
        parent, noun, place = ref.workbook, "Table", "workbook"
    if parent is not owner:
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT, _diagnostic_override=(f"{what} is a {noun} of another {place}", "")
        )
    return ref.index


def resolve_ref(ref: Union[int, str, Any], names: Callable[[], Iterable[str]], what: str, owner: Any = None) -> int:
    """Return the index ``ref`` addresses.

    Args:
      ref: an index, returned unchanged so the caller's index checks apply;
        a name looked up in ``names()``, ignoring case; or a :class:`Table`
        or :class:`Pivot` handle, resolved to its object's current index.
      names: lists the collection's names in index order; called only for a
        ``str`` reference.
      what: the collection's noun for the error message (``"table"``).
      owner: the :class:`Workbook` whose :class:`Table` handles, or the
        :class:`Sheet` whose :class:`Pivot` handles, ``ref`` may be; ``None``
        accepts no handle.

    Raises:
      TypeError: when ``ref`` is a ``bool`` or neither an integer, a ``str``
        nor a handle.
      NotFoundError: when no entry carries the name ``ref``, or the handle's
        object was removed.
      InvalidArgumentError: when ``ref`` is a handle of another workbook or sheet.
    """
    if isinstance(ref, str):
        wanted = _fold(ref)
        for index, name in enumerate(names()):
            if _fold(name) == wanted:
                return index
        raise FormulonError(StatusCode.NOT_FOUND, _diagnostic_override=(f"no {what} is named {ref!r}", ""))
    if isinstance(ref, bool):
        raise TypeError(f"formulon: {what} must be an int, a str or a handle, got bool")
    index = None if owner is None else _handle_index(ref, owner, what)
    if index is not None:
        return index
    try:
        return operator.index(ref)
    except TypeError:
        raise TypeError(f"formulon: {what} must be an int, a str or a handle, got {type(ref).__name__}") from None
