#!/usr/bin/env python3
"""Formulon function-coverage reporter.

Derives "implemented" from a static scan of C++ sources: builtin entries
inside `FunctionDef` array initialisers (any `{"NAME", <arity>u, ...}`
literal under `src/eval/`), entries in the `kLazyDispatch` table inside
`src/eval/tree_walker_lazy_table.cpp`, and special-form names inside the
`kNames[] = { ... }` block in `src/eval/special_forms_catalog.cpp`.
Derives "targeted" from the canonical catalog at
`tools/catalog/functions.txt`.

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
STATUS_PATH = REPO_ROOT / "tools" / "catalog" / "function_status.tsv"
EVAL_DIR = REPO_ROOT / "src" / "eval"
SPECIAL_FORMS_PATH = EVAL_DIR / "special_forms_catalog.cpp"
C_API_PATH = REPO_ROOT / "src" / "c_api" / "parts" / "function_catalog.cpp"

# Matches builtin entries inside a `FunctionDef` array initialiser, of
# the form `{"SUM", 1u, kVariadic, &Sum, ...}`. The arity column is the
# disambiguator — every builtin entry has an unsigned-literal arity right
# after the name, which is unique enough to keep the regex from matching
# arbitrary string-keyed maps elsewhere in the tree. Catches single-letter
# names like `N`/`T` too (the `*` quantifier allows zero suffix chars).
FUNCTION_DEF_RE = re.compile(r'\{\s*"([A-Z][A-Z0-9_.]*)"\s*,\s*\d+u\s*,')

# Inside `constexpr LazyEntry kLazyDispatch[] = { ... };` entries look like
# `{"IF", &eval_if_lazy},`. We anchor on the leading `{` and `"` to avoid
# catching arbitrary strings in code comments.
LAZY_ENTRY_RE = re.compile(r'\{\s*"([A-Z][A-Z0-9_.]*)"\s*,\s*&eval_')

# Inside `special_forms_catalog.cpp` the sole source of truth is a static
# array initialiser of the form
# `static constexpr const char* kNames[] = {"LET", nullptr};`. We scan the
# file for every string literal inside a `kNames[] = { ... }` block so
# future additions (LAMBDA, ...) are picked up without edits here.
SPECIAL_FORMS_BLOCK_RE = re.compile(r"kNames\s*\[\s*\]\s*=\s*\{([^}]*)\}", re.DOTALL)
SPECIAL_FORMS_NAME_RE = re.compile(r'"([A-Z][A-Z0-9_.]*)"')
C_API_AVAILABILITY_RE = re.compile(r'\{\s*"([A-Z][A-Z0-9_.]*)"\s*,\s*FM_FUNCTION_([A-Z_]+)\s*\}')


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


def load_function_status(path: Path) -> Dict[str, str]:
    """Parses the optional function availability override table.

    Omitted names default to `implemented`. The parser intentionally stays
    TSV + stdlib-only so `make function-status` has no PyYAML dependency.
    """
    statuses: Dict[str, str] = {}
    if not path.exists():
        return statuses
    valid = {
        "implemented",
        "implemented_unverified",
        "environment_bound",
        "unavailable_stub",
    }
    with path.open("r", encoding="utf-8") as fh:
        for lineno, raw in enumerate(fh, start=1):
            line = raw.rstrip("\n")
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            cols = line.split("\t")
            if len(cols) < 2:
                raise ValueError(f"{path}:{lineno}: expected NAME<TAB>STATUS")
            name = cols[0].strip()
            status = cols[1].strip()
            if status not in valid:
                raise ValueError(f"{path}:{lineno}: unknown status {status!r}")
            statuses[name] = status
    return statuses


def load_c_api_availability(path: Path) -> Dict[str, str]:
    """Parses the explicit non-default availability table in parts/function_catalog.cpp."""
    if not path.exists():
        return {}
    enum_to_status = {
        "IMPLEMENTED_UNVERIFIED": "implemented_unverified",
        "ENVIRONMENT_BOUND": "environment_bound",
        "UNAVAILABLE_STUB": "unavailable_stub",
    }
    out: Dict[str, str] = {}
    text = path.read_text(encoding="utf-8", errors="replace")
    for m in C_API_AVAILABILITY_RE.finditer(text):
        status = enum_to_status.get(m.group(2))
        if status is not None:
            out[m.group(1)] = status
    return out


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


def scan_lazy_names(eval_dir: Path) -> Set[str]:
    """Returns every name appearing as a `{"NAME", &eval_..._lazy}` entry
    in any `src/eval/**/*.cpp`. The historical `kLazyDispatch` table was
    split out of `tree_walker.cpp` into `tree_walker_lazy_table.cpp`, so a
    file-specific scan would miss it; the regex anchor (`{"NAME", &eval_`)
    is specific enough that a tree-wide rglob is safe."""
    names: Set[str] = set()
    for path in sorted(eval_dir.rglob("*.cpp")):
        text = path.read_text(encoding="utf-8", errors="replace")
        for m in LAZY_ENTRY_RE.finditer(text):
            names.add(m.group(1))
    return names


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
    special_forms = eval_dir / "special_forms_catalog.cpp"
    return scan_registered_names(eval_dir) | scan_lazy_names(eval_dir) | scan_special_form_names(special_forms)


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
    statuses: Dict[str, str],
    section_filter: Optional[str],
) -> None:
    total = len(total_names)
    done = len(implemented & total_names)
    unavailable = sum(1 for n in total_names if statuses.get(n, "implemented") == "unavailable_stub")
    unverified = sum(1 for n in total_names if statuses.get(n, "implemented") == "implemented_unverified")
    env_bound = sum(1 for n in total_names if statuses.get(n, "implemented") == "environment_bound")
    real_impl = done - unavailable
    pct = 0.0 if total == 0 else 100.0 * done / total
    real_pct = 0.0 if total == 0 else 100.0 * real_impl / total
    header = f"Formulon function recognition: {done} / {total} catalogued ({pct:.1f}%)"
    print(bold(header))
    print(
        "Availability: "
        f"{real_impl}/{total} real implementations ({real_pct:.1f}%), "
        f"{unavailable} unavailable stubs, "
        f"{unverified} implemented-unverified, "
        f"{env_bound} environment-bound"
    )
    print()

    for title, names in sections:
        if section_filter and section_filter.lower() not in title.lower():
            continue
        name_set = set(names)
        impl_here = sorted(n for n in name_set if n in implemented)
        miss_here = sorted(n for n in name_set if n not in implemented)
        stub_here = sorted(n for n in name_set if statuses.get(n, "implemented") == "unavailable_stub")
        unverified_here = sorted(n for n in name_set if statuses.get(n, "implemented") == "implemented_unverified")
        env_here = sorted(n for n in name_set if statuses.get(n, "implemented") == "environment_bound")
        print(bold(f"## {title}"))
        print(f"  Recognized {_count_line(len(impl_here), len(name_set))}:")
        print("    " + (" ".join(impl_here) if impl_here else red("(none)")))
        if stub_here or unverified_here or env_here:
            if stub_here:
                print(f"  Unavailable stubs {len(stub_here)}:")
                print("    " + " ".join(stub_here))
            if unverified_here:
                print(f"  Implemented but unverified {len(unverified_here)}:")
                print("    " + " ".join(unverified_here))
            if env_here:
                print(f"  Environment-bound {len(env_here)}:")
                print("    " + " ".join(env_here))
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


def check_c_api_availability(statuses: Dict[str, str]) -> int:
    """Checks that function_status.tsv and the C ABI availability table agree.

    The C ABI table only lists non-default statuses; omitted names mean
    `implemented`, matching function_status.tsv's defaulting rule.
    """
    c_api = load_c_api_availability(C_API_PATH)
    interesting_statuses = {
        "implemented_unverified",
        "environment_bound",
        "unavailable_stub",
    }
    wanted = {name: status for name, status in statuses.items() if status in interesting_statuses}
    drift = []
    for name in sorted(set(wanted) | set(c_api)):
        want = wanted.get(name, "implemented")
        got = c_api.get(name, "implemented")
        if want != got:
            drift.append((name, want, got))
    if not drift:
        return 0
    print("C API availability drift:", file=sys.stderr)
    for name, want, got in drift:
        print(f"  {name}: function_status.tsv={want}, c_api={got}", file=sys.stderr)
    return 1


# ---- main ----------------------------------------------------------------


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--missing",
        action="store_true",
        help="Print only names that are in the catalog but not implemented, one per line.",
    )
    parser.add_argument(
        "--orphans",
        action="store_true",
        help="Print names that are implemented in source but NOT listed in "
        "the catalog (the unit-test invariant should keep this empty).",
    )
    parser.add_argument(
        "--availability",
        choices=["implemented", "implemented_unverified", "environment_bound", "unavailable_stub"],
        default=None,
        help="Print catalog names with the selected availability status.",
    )
    parser.add_argument(
        "--check-c-api-availability",
        action="store_true",
        help="Exit 1 if function_status.tsv and the C ABI availability table drift.",
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

    try:
        sections, catalog_names = load_catalog(CATALOG_PATH)
        statuses = load_function_status(STATUS_PATH)
    except ValueError as exc:
        print(f"catalog status error: {exc}", file=sys.stderr)
        return 2
    implemented = scan_implemented(REPO_ROOT)

    unknown_status_names = sorted(set(statuses) - catalog_names)
    if unknown_status_names:
        print(
            "catalog status error: function_status.tsv contains names not in functions.txt: "
            + ", ".join(unknown_status_names),
            file=sys.stderr,
        )
        return 2

    if args.orphans:
        print_orphans(catalog_names, implemented)
        return 0
    if args.check_c_api_availability:
        return check_c_api_availability(statuses)
    if args.missing:
        print_missing_only(sections, implemented, args.category)
        return 0
    if args.availability:
        for name in sorted(catalog_names):
            if statuses.get(name, "implemented") == args.availability:
                print(name)
        return 0

    print_full_report(sections, implemented, catalog_names, statuses, args.category)
    return 0


if __name__ == "__main__":
    sys.exit(main())
