#!/usr/bin/env python3
"""Formulon function-coverage reporter.

Derives "implemented" from a static scan of C++ sources (authoritative
strings: `FunctionDef{"NAME"` inside `src/eval/builtins/*.cpp` and entries
in the `kLazyDispatch` table inside `src/eval/tree_walker.cpp`). Derives
"targeted" from the canonical catalog at `tools/catalog/functions.txt`.

Prints a per-category coverage report by default. Supports a few
pipe-friendly flags for CI / debugging:

    tools/catalog/status.py
    tools/catalog/status.py --missing           # every un-implemented name
    tools/catalog/status.py --category "Text"   # filter to one section
    tools/catalog/status.py --orphans           # registered-but-uncatalogued

Stdlib only. Works on CPython 3.8+.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
CATALOG_PATH = REPO_ROOT / "tools" / "catalog" / "functions.txt"
EVAL_DIR = REPO_ROOT / "src" / "eval"
TREE_WALKER_PATH = EVAL_DIR / "tree_walker.cpp"
SPECIAL_FORMS_PATH = EVAL_DIR / "special_forms_catalog.cpp"

# Matches `FunctionDef{"SUM"`, `FunctionDef def{"AND"`, etc. We require at
# least one uppercase letter at the start so we don't pick up free-form
# strings like `"foo"`.
FUNCTION_DEF_RE = re.compile(r'FunctionDef(?:\s+\w+)?\s*\{\s*"([A-Z][A-Z0-9_.]*)"')

# Inside `constexpr LazyEntry kLazyDispatch[] = { ... };` entries look like
# `{"IF", &eval_if_lazy},`. We anchor on the leading `{` and `"` to avoid
# catching arbitrary strings in code comments.
LAZY_ENTRY_RE = re.compile(r'\{\s*"([A-Z][A-Z0-9_.]*)"\s*,\s*&eval_')

# Inside `special_forms_catalog.cpp` the sole source of truth is a static
# array initialiser of the form
# `static constexpr const char* kNames[] = {"LET", nullptr};`. We scan the
# file for every string literal inside a `kNames[] = { ... }` block so
# future additions (LAMBDA, ...) are picked up without edits here.
SPECIAL_FORMS_BLOCK_RE = re.compile(
    r"kNames\s*\[\s*\]\s*=\s*\{([^}]*)\}", re.DOTALL
)
SPECIAL_FORMS_NAME_RE = re.compile(r'"([A-Z][A-Z0-9_.]*)"')


# ---- ANSI helpers --------------------------------------------------------

def _supports_color() -> bool:
    if os.environ.get("NO_COLOR"):
        return False
    return sys.stdout.isatty()


_ANSI = _supports_color()


def _c(code: str, text: str) -> str:
    if not _ANSI:
        return text
    return f"\033[{code}m{text}\033[0m"


def green(s: str) -> str:
    return _c("32", s)


def red(s: str) -> str:
    return _c("31", s)


def yellow(s: str) -> str:
    return _c("33", s)


def bold(s: str) -> str:
    return _c("1", s)


# ---- Catalog parsing -----------------------------------------------------

def load_catalog(path: Path) -> Tuple[List[Tuple[str, List[str]]], Set[str]]:
    """Parses `functions.txt`. Returns (sections, all_names).

    `sections` is a list of (section_title, [names_in_section]) preserving
    file order. `all_names` is the deduped set of every name.

    Lines starting with `#` are treated as either the preamble (discarded)
    or a section title. A new section begins at each `# 11.3.x ...` line —
    they are the only `#` lines that delimit groups in the current file.
    Blank lines are tolerated within a section.
    """
    sections: List[Tuple[str, List[str]]] = []
    current_title: Optional[str] = None
    current_names: List[str] = []
    all_names: Set[str] = set()

    def flush() -> None:
        nonlocal current_title, current_names
        if current_title is not None:
            sections.append((current_title, current_names))
        current_title = None
        current_names = []

    with path.open("r", encoding="utf-8") as fh:
        for raw in fh:
            line = raw.rstrip("\n")
            stripped = line.strip()
            if not stripped:
                continue
            if stripped.startswith("#"):
                # Section titles look like `# 11.3.1 Math & Trig (約 85)`.
                # Anything else is preamble / prose and just ends the
                # previous section if one was open.
                body = stripped.lstrip("#").strip()
                if re.match(r"^11\.\d+(?:\.\d+)?\s", body):
                    flush()
                    current_title = body
                    current_names = []
                continue
            name = stripped
            if current_title is None:
                # Names outside any section are still collected so the
                # invariant test covers them; we synthesise a holding
                # bucket called "(uncategorised)".
                current_title = "(uncategorised)"
                current_names = []
            current_names.append(name)
            all_names.add(name)
    flush()
    return sections, all_names


# ---- Source scanning -----------------------------------------------------

def scan_registered_names(eval_dir: Path) -> Set[str]:
    """Returns every name appearing inside a `FunctionDef{"NAME"` literal
    under `src/eval/` (recursively). Covers builtins + any host extensions
    that follow the same registration pattern."""
    names: Set[str] = set()
    for path in sorted(eval_dir.rglob("*.cpp")):
        text = path.read_text(encoding="utf-8", errors="replace")
        for m in FUNCTION_DEF_RE.finditer(text):
            names.add(m.group(1))
    return names


def scan_lazy_names(tree_walker_path: Path) -> Set[str]:
    """Returns every name appearing as a `{"NAME", &eval_..._lazy}` entry
    inside `kLazyDispatch`. The broader `&eval_` anchor deliberately catches
    any future lazy impl that follows the naming convention."""
    if not tree_walker_path.exists():
        return set()
    text = tree_walker_path.read_text(encoding="utf-8", errors="replace")
    return {m.group(1) for m in LAZY_ENTRY_RE.finditer(text)}


def scan_special_form_names(path: Path) -> Set[str]:
    """Returns every UPPERCASE string literal inside the `kNames[] = { ... }`
    initialiser in `special_forms_catalog.cpp`. These are the parser-
    integrated special forms (LET today, LAMBDA later) that don't reach the
    registry or the lazy-dispatch table. Falls back to the hard-coded
    baseline {"LET"} if the file is missing or the initialiser cannot be
    located, so `make function-status` still reports sensibly on a tree
    that's been half-rebased."""
    fallback = {"LET"}
    if not path.exists():
        return fallback
    text = path.read_text(encoding="utf-8", errors="replace")
    block = SPECIAL_FORMS_BLOCK_RE.search(text)
    if block is None:
        return fallback
    found = {m.group(1) for m in SPECIAL_FORMS_NAME_RE.finditer(block.group(1))}
    return found or fallback


def scan_implemented(repo_root: Path) -> Set[str]:
    eval_dir = repo_root / "src" / "eval"
    tree_walker = eval_dir / "tree_walker.cpp"
    special_forms = eval_dir / "special_forms_catalog.cpp"
    return (
        scan_registered_names(eval_dir)
        | scan_lazy_names(tree_walker)
        | scan_special_form_names(special_forms)
    )


# ---- Reporting -----------------------------------------------------------

def _count_line(impl: int, total: int) -> str:
    pct = 0.0 if total == 0 else 100.0 * impl / total
    raw = f"{impl}/{total} ({pct:.1f}%)"
    if impl == total and total > 0:
        return green(raw)
    if impl == 0:
        return red(raw)
    return yellow(raw)


def print_full_report(
    sections: Sequence[Tuple[str, Sequence[str]]],
    implemented: Set[str],
    total_names: Set[str],
    section_filter: Optional[str],
) -> None:
    total = len(total_names)
    done = len(implemented & total_names)
    pct = 0.0 if total == 0 else 100.0 * done / total
    header = f"Formulon function coverage: {done} / {total} implemented ({pct:.1f}%)"
    print(bold(header))
    print()

    for title, names in sections:
        if section_filter and section_filter.lower() not in title.lower():
            continue
        name_set = set(names)
        impl_here = sorted(n for n in name_set if n in implemented)
        miss_here = sorted(n for n in name_set if n not in implemented)
        print(bold(f"## {title}"))
        print(f"  Implemented {_count_line(len(impl_here), len(name_set))}:")
        print("    " + (" ".join(impl_here) if impl_here else red("(none)")))
        print(f"  Missing {len(miss_here)}:")
        if miss_here:
            print("    " + " ".join(miss_here))
        else:
            print("    " + green("(none)"))
        print()


def print_missing_only(
    sections: Sequence[Tuple[str, Sequence[str]]],
    implemented: Set[str],
    section_filter: Optional[str],
) -> None:
    for title, names in sections:
        if section_filter and section_filter.lower() not in title.lower():
            continue
        for name in names:
            if name not in implemented:
                print(name)


def print_orphans(catalog_names: Set[str], implemented: Set[str]) -> None:
    orphans = sorted(implemented - catalog_names)
    for name in orphans:
        print(name)


# ---- main ----------------------------------------------------------------

def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--missing",
        action="store_true",
        help="Print only names that are in the catalog but not implemented, "
        "one per line.",
    )
    parser.add_argument(
        "--orphans",
        action="store_true",
        help="Print names that are implemented in source but NOT listed in "
        "the catalog (the unit-test invariant should keep this empty).",
    )
    parser.add_argument(
        "--category",
        metavar="NAME",
        default=None,
        help="Substring filter (case-insensitive) applied to section titles.",
    )
    args = parser.parse_args(argv)

    if not CATALOG_PATH.exists():
        print(f"catalog not found: {CATALOG_PATH}", file=sys.stderr)
        return 2

    sections, catalog_names = load_catalog(CATALOG_PATH)
    implemented = scan_implemented(REPO_ROOT)

    if args.orphans:
        print_orphans(catalog_names, implemented)
        return 0
    if args.missing:
        print_missing_only(sections, implemented, args.category)
        return 0

    print_full_report(sections, implemented, catalog_names, args.category)
    return 0


if __name__ == "__main__":
    sys.exit(main())
