"""The type stub describes the same constructors and enum members as the runtime.

``py.typed`` makes ``__init__.pyi`` the type surface a checker sees, so a
dataclass the stub declares without its synthesized ``__init__`` rejects valid
calls, and an enum written with annotation-only members has no members at all.
"""

from __future__ import annotations

import ast
import dataclasses
import enum
import inspect
import unittest
from pathlib import Path

import formulon

_STUB = Path(formulon.__file__).resolve().with_name("__init__.pyi")


def _stub_classes() -> dict:
    tree = ast.parse(_STUB.read_text(encoding="utf-8"))
    return {node.name: node for node in tree.body if isinstance(node, ast.ClassDef)}


def _is_stub_dataclass(node: ast.ClassDef) -> bool:
    return any(ast.unparse(d).startswith("dataclass") for d in node.decorator_list)


def _stub_init_params(node: ast.ClassDef):
    """``[(name, has_default), ...]`` of the constructor the stub declares, or None."""
    for item in node.body:
        if isinstance(item, ast.FunctionDef) and item.name == "__init__":
            args = item.args
            positional = args.args[1:]
            defaults = [False] * (len(positional) - len(args.defaults)) + [True] * len(args.defaults)
            params = [(a.arg, d) for a, d in zip(positional, defaults)]
            params += [(a.arg, d is not None) for a, d in zip(args.kwonlyargs, args.kw_defaults)]
            return params
    if _is_stub_dataclass(node):
        return [
            (item.target.id, item.value is not None)
            for item in node.body
            if isinstance(item, ast.AnnAssign) and isinstance(item.target, ast.Name)
        ]
    return None


def _runtime_init_params(cls) -> list:
    return [
        (name, param.default is not inspect.Parameter.empty)
        for name, param in inspect.signature(cls).parameters.items()
    ]


def _int_literal(node: ast.AST):
    if isinstance(node, ast.Constant):
        return node.value
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub) and isinstance(node.operand, ast.Constant):
        return -node.operand.value
    return ast.unparse(node)


class StubRuntimeParityTests(unittest.TestCase):
    def test_dataclass_constructors_match(self) -> None:
        checked = 0
        for name, node in sorted(_stub_classes().items()):
            runtime = getattr(formulon, name, None)
            if runtime is None or not dataclasses.is_dataclass(runtime):
                continue
            checked += 1
            with self.subTest(cls=name):
                self.assertEqual(_stub_init_params(node), _runtime_init_params(runtime))
        self.assertGreater(checked, 0)

    def test_enum_members_match(self) -> None:
        checked = 0
        for name, node in sorted(_stub_classes().items()):
            runtime = getattr(formulon, name, None)
            if not (isinstance(runtime, type) and issubclass(runtime, enum.Enum)):
                continue
            checked += 1
            stub_members = {
                item.targets[0].id: _int_literal(item.value)
                for item in node.body
                if isinstance(item, ast.Assign) and isinstance(item.targets[0], ast.Name)
            }
            with self.subTest(cls=name):
                self.assertEqual(stub_members, {member.name: member.value for member in runtime})
        self.assertGreater(checked, 0)

    def test_no_annotation_names_the_field_it_annotates(self) -> None:
        # `bytes: bytes` makes the annotation resolve to the field itself.
        for name, node in sorted(_stub_classes().items()):
            for item in ast.walk(node):
                if isinstance(item, ast.AnnAssign) and isinstance(item.target, ast.Name):
                    with self.subTest(cls=name, field=item.target.id):
                        self.assertNotEqual(ast.unparse(item.annotation), item.target.id)
                if isinstance(item, ast.arg) and item.annotation is not None:
                    with self.subTest(cls=name, param=item.arg):
                        self.assertNotEqual(ast.unparse(item.annotation), item.arg)


if __name__ == "__main__":
    unittest.main()
