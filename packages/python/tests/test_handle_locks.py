"""Every handle class serializes its whole operations on the shared engine instance."""

from __future__ import annotations

import unittest

import formulon
from formulon._c import _NoCopy


def _handle_classes() -> "list[type]":
    out, todo = [], list(_NoCopy.__subclasses__())
    while todo:
        cls = todo.pop()
        out.append(cls)
        todo.extend(cls.__subclasses__())
    return [cls for cls in out if cls.__module__.startswith("formulon.")]


def _unlocked(cls: type) -> "list[str]":
    names = []
    for name, descriptor in cls.__dict__.items():
        if isinstance(descriptor, property):
            functions = [f for f in (descriptor.fget, descriptor.fset, descriptor.fdel) if f is not None]
        elif isinstance(descriptor, (classmethod, staticmethod)):
            functions = [descriptor.__func__]
        elif callable(descriptor) and not isinstance(descriptor, type):
            functions = [descriptor]
        else:
            continue
        if any(not hasattr(f, "__wrapped__") for f in functions):
            names.append(name)
    return names


class HandleLockTests(unittest.TestCase):
    def test_every_handle_class_holds_the_engine_lock_through_each_operation(self) -> None:
        classes = _handle_classes()
        for expected in (formulon.Workbook, formulon.Sheet, formulon.Table, formulon.PivotCache, formulon.Pivot):
            self.assertIn(expected, classes)
        for cls in classes:
            with self.subTest(cls=cls.__name__):
                self.assertEqual(_unlocked(cls), [])


if __name__ == "__main__":
    unittest.main()
