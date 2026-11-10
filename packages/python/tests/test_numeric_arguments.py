"""Numeric arguments: a ``float`` or ``int`` parameter refuses a string (and ``None``) with a labelled ``TypeError``.

The cases are derived from the public signatures, so a new number parameter is covered without a hand-written test.
"""

from __future__ import annotations

import inspect
import typing
import unittest
from typing import Any, Dict, List, Tuple

import formulon
from formulon import Workbook

_NUMERIC = (float, int)

# Placeholder values for the other parameters of a call, by annotation. A required parameter whose annotation is
# not listed makes the method unreachable without a hand-built argument, and that method is skipped.
_PLACEHOLDERS: Dict[Any, Any] = {int: 0, float: 1.0, str: "A1", bool: False}

# Public classes with an instance a fresh workbook can hand over. The other handle classes (charts, pivots,
# drawings, tables, slicers) need a populated sheet to exist at all and are not constructed here.
_SUBJECTS = ("Workbook", "Sheet", "WorkbookOoxml", "SheetOoxml")

# Parameters that are only valid together: while one is under test the others of its group take a placeholder.
_COMPANIONS: Dict[Tuple[str, str], Dict[str, Any]] = {("Workbook", "parse_reference"): {"base_row": 0, "base_col": 0}}

# Methods left out of the sweep. None needed today; an entry names a method whose arguments the placeholders
# cannot make valid, with the reason beside it.
_SKIP: set = set()

# Coordinate pairs also take an A1 string in the first slot, so a string there selects the other form rather than
# failing as a number; the pair is exercised through its sibling parameters only.
_PAIRS = (("row", "col"), ("first", "last"), ("anchor_row", "anchor_col"))

# Cases that must exist, so a change to the discovery cannot empty the test silently.
_REQUIRED = {
    ("Sheet", "set_number", "value"),
    ("Sheet", "set_row_height", "height_pt"),
    ("Sheet", "set_column_width", "width_chars"),
}


def _subject(name: str) -> Any:
    wb = Workbook.create_default()
    sheet = wb.sheet(0)
    return {"Workbook": wb, "Sheet": sheet, "WorkbookOoxml": wb.ooxml, "SheetOoxml": sheet.ooxml}[name]


def _resolve(parameter: inspect.Parameter, owner: type) -> Any:
    """The annotation of ``parameter``, evaluated from its string form in the owner's module."""
    annotation = parameter.annotation
    if not isinstance(annotation, str):
        return annotation
    namespace = dict(vars(formulon))
    namespace.update(vars(inspect.getmodule(owner)))
    try:
        return eval(annotation, namespace)  # noqa: S307 - the package's own annotations
    except Exception:  # noqa: BLE001 - an unresolvable annotation leaves the parameter out of the sweep
        return None


def _unwrap(annotation: Any) -> Tuple[Any, bool]:
    """``(base, optional)`` of an annotation: ``Optional[float]`` is ``(float, True)``."""
    if typing.get_origin(annotation) is typing.Union:
        members = [m for m in typing.get_args(annotation) if m is not type(None)]
        if len(members) == 1 and len(members) < len(typing.get_args(annotation)):
            return members[0], True
    return annotation, False


def _cases() -> List[Tuple[str, str, str, bool, Dict[str, Any]]]:
    """``(class, method, param, optional, placeholders)`` for every number parameter that can be exercised."""
    cases = []
    for class_name in _SUBJECTS:
        cls = getattr(formulon, class_name)
        for name, function in vars(cls).items():
            if name.startswith("_") or not inspect.isfunction(function) or (class_name, name) in _SKIP:
                continue
            params = [p for p in inspect.signature(function).parameters.values() if p.name != "self"]
            companions = _COMPANIONS.get((class_name, name), {})
            for target in params:
                base, optional = _unwrap(_resolve(target, cls))
                if base not in _NUMERIC or target.kind in (target.VAR_POSITIONAL, target.VAR_KEYWORD):
                    continue
                if any(target.name in pair and set(pair) <= {p.name for p in params} for pair in _PAIRS):
                    continue
                others: Dict[str, Any] = {}
                for other in params:
                    if other is target or other.kind in (other.VAR_POSITIONAL, other.VAR_KEYWORD):
                        continue
                    if other.name in companions:
                        others[other.name] = companions[other.name]
                    elif other.default is inspect.Parameter.empty:
                        placeholder = _PLACEHOLDERS.get(_unwrap(_resolve(other, cls))[0], _MISSING)
                        if placeholder is _MISSING:
                            break
                        others[other.name] = placeholder
                else:
                    cases.append((class_name, name, target.name, optional, others))
    return cases


def _call(subject: Any, method: str, param: str, value: Any, others: Dict[str, Any]) -> None:
    """Call ``method`` with ``value`` for ``param``: by position up to the last given parameter, keyword-only by name."""
    function = getattr(type(subject), method)
    params = [p for p in inspect.signature(function).parameters.values() if p.name != "self"]
    given = {**others, param: value}
    positional = [p for p in params if p.kind is p.POSITIONAL_OR_KEYWORD]
    last = max((i for i, p in enumerate(positional) if p.name in given), default=-1)
    args = [given.get(p.name, p.default) for p in positional[: last + 1]]
    kwargs = {p.name: given[p.name] for p in params if p.kind is p.KEYWORD_ONLY and p.name in given}
    getattr(subject, method)(*args, **kwargs)


_MISSING = object()


class NumericArgumentTests(unittest.TestCase):
    def test_required_cases_are_discovered(self) -> None:
        found = {(c, m, p) for c, m, p, _, _ in _cases()}
        self.assertLessEqual(_REQUIRED, found)

    def test_number_parameters_refuse_a_string(self) -> None:
        for class_name, method, param, optional, others in _cases():
            bad_values = ["1"] if optional else ["1", None]
            for bad in bad_values:
                with self.subTest(call=f"{class_name}.{method}", param=param, value=bad):
                    with self.assertRaises(TypeError) as ctx:
                        _call(_subject(class_name), method, param, bad, others)
                    self.assertTrue(
                        str(ctx.exception).startswith(f"{class_name}.{method} {param}"),
                        str(ctx.exception),
                    )


if __name__ == "__main__":
    unittest.main()
