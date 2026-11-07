#!/usr/bin/env python3
"""Compare the methods the ledger gives each binding class with the Python class that implements it.

For each class named on the command line, the ledger side is every op whose `handle.class` lists the class
(a list counts for each class in it; a method a base class carries counts for the base only), and for
`Workbook` / `Sheet` also the ops of that scope the class publishes. The Python side is the public methods
and properties of the class of that name in packages/python/formulon/*.py; an identity property the facade
builds itself (`index`, `object_id`, ...) counts only when the ledger names it.

Prints `Class ledger=N python=M` per class and exits 1 when any pair differs.
"""

from __future__ import annotations

import argparse
import ast
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT / "tools" / "codegen"))

import facade_scope  # noqa: E402

from tools.dev.binding_drift.api_conventions import (  # noqa: E402
    HANDLE_IDENTITY_PROPERTIES,
    handle_classes,
    handle_name,
)
from tools.dev.binding_drift.api_conventions_bindings import FACADE_EXTRA  # noqa: E402

PY_PKG = REPO_ROOT / "packages" / "python" / "formulon"
SCOPE_OF_CLASS = {"Workbook": "workbook", "Sheet": "sheet"}


def ledger_names(ledger: facade_scope.Ledger, cls: str) -> Set[str]:
    """The Python names the ledger publishes on `cls`."""
    out: Set[str] = set()
    for _key, entry in sorted(ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") not in ("op", "list"):
            continue
        if cls in handle_classes(entry) and facade_scope._on_language(entry, "py"):
            name = handle_name(entry, "py", ledger.collections)
            if name:
                out.add(name)
    scope = SCOPE_OF_CLASS.get(cls)
    if scope is not None:
        out.update(facade_scope.methods(ledger, scope, "py"))
        out.update(FACADE_EXTRA["py"].get(cls, frozenset()))
    if cls == "Workbook":
        for _key, entry in ledger.entries.items():
            name = facade_scope.public_name(entry, "py") if isinstance(entry, dict) else None
            if entry.get("kind") == "factory" and name:
                out.add(name.rpartition(".")[2])
    return out


def python_names(cls: str, package: Path = PY_PKG) -> Optional[Set[str]]:
    """Public method and property names defined in the body of class `cls`, or None when no module defines it."""
    for path in sorted(package.glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in tree.body:
            if isinstance(node, ast.ClassDef) and node.name == cls:
                return {
                    item.name
                    for item in node.body
                    if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)) and not item.name.startswith("_")
                }
    return None


def compare(cls: str, ledger: facade_scope.Ledger, package: Path = PY_PKG) -> Dict[str, object]:
    """`{"ledger": N, "python": M, "missing": [...], "extra": [...]}` of one class."""
    want = ledger_names(ledger, cls)
    found = python_names(cls, package)
    have = set() if found is None else {n for n in found if n in want or n not in HANDLE_IDENTITY_PROPERTIES}
    return {
        "ledger": len(want),
        "python": len(have),
        "missing": sorted(want - have),
        "extra": sorted(have - want),
        "absent": found is None,
    }


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("classes", nargs="+", metavar="Class")
    parser.add_argument("-v", "--verbose", action="store_true", help="list the names that differ")
    args = parser.parse_args(argv)
    ledger = facade_scope.load()
    failed = False
    for cls in args.classes:
        result = compare(cls, ledger)
        print(f"{cls} ledger={result['ledger']} python={result['python']}")
        differs = result["ledger"] != result["python"] or result["missing"] or result["extra"]
        if differs:
            failed = True
            if args.verbose or result["absent"]:
                lines: List[str] = []
                if result["absent"]:
                    lines.append(f"  no class {cls} in {PY_PKG.relative_to(REPO_ROOT)}")
                lines += [f"  missing: {n}" for n in result["missing"]]
                lines += [f"  extra: {n}" for n in result["extra"]]
                print("\n".join(lines))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
