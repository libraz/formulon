"""Which patch keys admit `null`: the ledger, the `.d.ts` patch types and the Python `update_*` keywords agree."""

from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from .api_conventions import Ledger, load_ledger
from .surface_files import PYTHON_PKG_DIR, WASM_DTS, _read

# Ledger patch records whose op names a `*Properties` snapshot instead of `<Name>Patch`; `None` is a family
# of shared arguments that has no patch record.
_PATCH_RECORD_FOR_SNAPSHOT: Dict[str, Optional[str]] = {
    "TableProperties": "record.table_patch",
    "ShapeProperties": "record.shape_patch",
    "SlicerProperties": "record.slicer_patch",
    "ChartProperties": "record.chart_display_patch",
    "DrawingProperties": None,
}
# `description` takes `None` on every drawing-family update by design (an empty description).
_DRAWING_FAMILY_SNAPSHOTS = frozenset({"DrawingProperties", "ShapeProperties", "SlicerProperties"})
_DRAWING_FAMILY_CLEARS = frozenset({"description"})
# Class-name prefixes an op drops from its method name on its handle class.
_HANDLE_PREFIXES = ("pivot_cache", "pivot", "chart", "table", "shape", "slicer", "drawing")


def _snake(name: str) -> str:
    return re.sub(r"(?<=[a-z0-9])([A-Z])", lambda m: "_" + m.group(1), name).lower()


def _strip_comments(text: str) -> str:
    text = re.sub(r"/\*.*?\*/", "", text, flags=re.S)
    return re.sub(r"(?m)^\s*//.*$", "", text)


def _split_top(text: str, seps: str) -> List[str]:
    """`text` cut at the separators that sit outside every bracket pair."""
    parts, depth, start, i = [], 0, 0, 0
    while i < len(text):
        ch = text[i]
        if ch == "=" and text[i + 1 : i + 2] == ">":
            i += 2
            continue
        if ch in "<{([":
            depth += 1
        elif ch in ">})]":
            depth -= 1
        elif ch in seps and depth == 0:
            parts.append(text[start:i])
            start = i + 1
        i += 1
    parts.append(text[start:])
    return [p.strip() for p in parts if p.strip()]


def _balanced(text: str, open_at: int) -> str:
    """The text between the bracket at `open_at` and its partner."""
    depth, i = 0, open_at
    while i < len(text):
        ch = text[i]
        if ch == "=" and text[i + 1 : i + 2] == ">":
            i += 2
            continue
        if ch in "<{([":
            depth += 1
        elif ch in ">})]":
            depth -= 1
            if depth == 0:
                return text[open_at + 1 : i]
        i += 1
    return text[open_at + 1 :]


def _member_nulls(body: str) -> Set[str]:
    """Names of the members of an object-type body whose type admits `null`."""
    out: Set[str] = set()
    for member in _split_top(body, ";,"):
        m = re.match(r"(?:readonly\s+)?(\w+)\??\s*:\s*(.*)$", member, re.S)
        if m and re.search(r"\bnull\b", m.group(2)):
            out.add(m.group(1))
    return out


def dts_null_keys(dts_text: str, name: str) -> Optional[Set[str]]:
    """camelCase keys of the exported patch type `name` that admit `null`; `None` when it is not declared.

    A `Patch<T, 'a' | 'b'>` admits `null` on the literals of its second argument; an object type
    (interface, or a brace block outside `Patch<>`) on each member whose type mentions `null`.
    """
    text = _strip_comments(dts_text)
    m = re.search(r"export interface %s(?:\s+extends\s+[^{]+?)?\s*\{" % re.escape(name), text)
    if m:
        return _member_nulls(_balanced(text, m.end() - 1))
    m = re.search(r"export type %s\s*=" % re.escape(name), text)
    if not m:
        return None
    expr = text[m.end() :]
    end = len(_split_top(expr + ";", ";")[0])
    expr = expr[:end]
    out: Set[str] = set()
    i = 0
    while i < len(expr):
        if expr.startswith("Patch<", i):
            args = _split_top(_balanced(expr, i + len("Patch")), ",")
            if len(args) > 1:
                out |= set(re.findall(r"'(\w+)'", args[1]))
            i += len("Patch") + len(_balanced(expr, i + len("Patch"))) + 2
        elif expr[i] == "{":
            body = _balanced(expr, i)
            out |= _member_nulls(body)
            i += len(body) + 2
        else:
            i += 1
    return out


def _python_update_methods(package: Path) -> Dict[Tuple[str, str], Dict[str, bool]]:
    """`(class, method)` -> {keyword-only parameter: annotation admits None} for each `update*` method."""
    out: Dict[Tuple[str, str], Dict[str, bool]] = {}
    for path in sorted(package.glob("*.py")):
        for cls in ast.walk(ast.parse(_read(path))):
            if not isinstance(cls, ast.ClassDef):
                continue
            for fn in cls.body:
                if isinstance(fn, ast.FunctionDef) and fn.name.startswith("update"):
                    out[(cls.name, fn.name)] = {
                        a.arg: _admits_none(a.annotation) for a in fn.args.kwonlyargs if a.annotation is not None
                    }
    return out


def _admits_none(annotation: ast.AST) -> bool:
    for node in ast.walk(annotation):
        if isinstance(node, ast.Constant) and (
            node.value is None or (isinstance(node.value, str) and re.search(r"\bNone\b", node.value))
        ):
            return True
        if isinstance(node, ast.Name) and node.id == "Optional":
            return True
    return False


def _method_names(op: dict) -> List[str]:
    handle = op.get("handle") or {}
    names = [handle["py"]] if handle.get("py") else []
    py = op["py"]
    names.append(py)
    if py.startswith("update_"):
        rest = py[len("update_") :]
        for prefix in _HANDLE_PREFIXES:
            if rest == prefix:
                names.append("update")
            elif rest.startswith(prefix + "_"):
                names.append("update_" + rest[len(prefix) + 1 :])
    return names


def _op_class(op: dict) -> str:
    handle = op.get("handle") or {}
    return handle.get("class") or ("Sheet" if op.get("scope") == "sheet" else "Workbook")


def check_patch_nullability(
    ledger: Optional[Ledger] = None, dts_text: Optional[str] = None, package: Optional[Path] = None
) -> List[str]:
    """Each patch record with a `clearable:` list agrees with its `.d.ts` type and Python `update_*` keywords."""
    ledger = load_ledger() if ledger is None else ledger
    if not ledger.entries:
        return list(ledger.problems)
    dts = _read(WASM_DTS) if dts_text is None else dts_text
    methods = _python_update_methods(PYTHON_PKG_DIR if package is None else package)
    problems: List[str] = []
    by_js = {
        v["js"]: (k, v)
        for k, v in ledger.entries.items()
        if v.get("kind") == "record" and v.get("role") == "patch" and "clearable" in v
    }
    for key, record in sorted(by_js.values()):
        clearable = set(record["clearable"])
        declared = dts_null_keys(dts, record["js"])
        if declared is None:
            problems.append(f"patch-nullability: {key}: {record['js']} is not declared in the WASM d.ts")
        elif {_snake(k) for k in declared} != clearable:
            got = sorted(_snake(k) for k in declared)
            problems.append(
                f"patch-nullability: {key}: d.ts {record['js']} admits null on {got}, ledger clearable is {sorted(clearable)}"
            )
    # A patch record without a `clearable:` list declares no clear rule to compare.
    unlisted = {
        v["js"]
        for v in ledger.entries.values()
        if v.get("kind") == "record" and v.get("role") == "patch" and "clearable" not in v
    }
    seen: Set[Tuple[str, str]] = set()
    for opkey, op in sorted(ledger.entries.items()):
        if op.get("kind") != "op" or not op.get("patch") or "py" not in op.get("surfaces", []):
            continue
        snapshot = op["patch_record"]
        if snapshot in _PATCH_RECORD_FOR_SNAPSHOT:
            link = _PATCH_RECORD_FOR_SNAPSHOT[snapshot]
            clearable = set(ledger.entries[link]["clearable"]) if link else set()
        elif snapshot + "Patch" in by_js:
            clearable = set(by_js[snapshot + "Patch"][1]["clearable"])
        elif snapshot + "Patch" in unlisted:
            continue
        else:
            problems.append(f"patch-nullability: {opkey}: no patch record named {snapshot}Patch in the ledger")
            continue
        if snapshot in _DRAWING_FAMILY_SNAPSHOTS:
            clearable |= _DRAWING_FAMILY_CLEARS
        found = [(_op_class(op), n) for n in _method_names(op) if (_op_class(op), n) in methods]
        if not found:
            problems.append(f"patch-nullability: {opkey}: Python has no method for {op['py']}")
            continue
        cls, name = found[0]
        if (cls, name) in seen:
            continue
        seen.add((cls, name))
        params = methods[(cls, name)]
        for param in sorted(params):
            if params[param] != (param in clearable):
                state = "admits" if params[param] else "excludes"
                problems.append(
                    f"patch-nullability: {cls}.{name}({param}) {state} None, "
                    f"ledger clearable is {'set' if param in clearable else 'not set'} for it"
                )
        for missing in sorted(clearable - set(params)):
            problems.append(f"patch-nullability: {cls}.{name} has no keyword `{missing}`, which the ledger lets clear")
    return problems
