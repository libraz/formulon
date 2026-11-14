#!/usr/bin/env python3
"""Source-level seam guards.

Some invariants in this codebase are enforced by "there is exactly one place
that does this", and none is expressible in the type system. A second
implementation is not a compile error -- it is a silent loss of whatever the
single implementation was there to guarantee. This script fails when one
appears. It also guards the include layering of `src/`.

Checks:
  array-alloc      Every `ArrayValue` header is allocated by
                   `eval::allocate_array_value`.

                   An `ArrayValue` is a `(rows, cols)` header over a flat
                   `Value[]` of length `rows * cols`. On wasm32 `size_t` is
                   32 bits, so a bare product of two `uint32_t` axes wraps:
                   the buffer comes back short while the caller's fill loop
                   still walks the full logical extent, writing past the
                   arena block. The seam is the one place that does the
                   bounds-checked multiplication, the grid-extent check, the
                   cell ceiling and the arena-exhaustion check together.

                   The check targets the header rather than the `Value[]`
                   buffer: a header cannot exist without one of these calls,
                   whereas plain `Value[]` scratch buffers (argument vectors,
                   for instance) are unrelated to array shape and legitimately
                   allocate on their own.

  xml-passthrough  Raw-XML retention goes through `io::append_raw_xml` /
                   `io::raw_xml` / `io::capture_unknown_children`.

                   A part the reader consumes drops out of the unknown-part
                   passthrough sweep, so anything in it the model does not
                   represent is lost on the next save unless it is retained
                   verbatim. Every reader that consumes a part needs the same
                   retention format; before it was shared, seven of them had
                   grown their own copy, and each new copy is a chance to get
                   the format -- or the "which children are unmodelled"
                   question -- subtly different.

  layers           `src/` includes only downward through the layer graph.

                   Bottom to top: utils < value < parser < model < pivot
                   engine < eval < {cf engine, print}. `io/` sits beside
                   `eval/` and never includes it, `workbook.cpp` (with its
                   `workbook_ref_rewrite` helper) is the facade over all of
                   them, and the app surfaces (c_api, cli, wasm, node_addon)
                   sit on top. Each layer's allowed
                   includes are listed in `LAYER_ALLOWED`. A second rule bans
                   the `io::` namespace token (outside comments) below `io/`,
                   so a model or engine type never depends on a reader or
                   writer. Known exceptions are listed in `LAYER_BASELINE`
                   with their reason.

  limit-origin     Every `kMax*` ceiling carries an origin tag and is refused accordingly.

                   `// @limit: engine|excel|format|internal` sits on the line
                   above the definition or trails it (`# @limit:` in the
                   Python facade). An `engine` ceiling is refused with
                   SecurityResourceLimit and never clamps silently; an
                   `excel` or `format` ceiling is never SecurityResourceLimit.
                   Not checked: ceilings reached through a variable, codes
                   chosen in a callee, and whether a tag is semantically true.

  all              Run every check above (default).

Stdlib only; no build artifacts or network access required.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SRC_ROOT = REPO_ROOT / "src"

SOURCE_SUFFIXES = (".cpp", ".h", ".cc")


def _scan(pattern: re.Pattern[str], allowed: set[Path]) -> list[str]:
    """Returns `path:line: text` for every match outside `allowed`."""
    hits: list[str] = []
    for path in sorted(SRC_ROOT.rglob("*")):
        if path.suffix not in SOURCE_SUFFIXES or path in allowed:
            continue
        for lineno, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
            if pattern.search(line):
                hits.append(f"{path.relative_to(REPO_ROOT)}:{lineno}: {line.strip()}")
    return hits


def _report(name: str, hits: list[str], remedy: str) -> bool:
    if not hits:
        print(f"{name}: ok")
        return True
    print(f"{name}: FAILED", file=sys.stderr)
    for hit in hits:
        print(f"  {hit}", file=sys.stderr)
    print(f"\n{remedy}", file=sys.stderr)
    return False


# Layer graph. Each group maps to the groups whose headers it may include
# (itself is always allowed).
LAYER_ALLOWED: dict[str, frozenset[str]] = {
    "utils": frozenset({"utils"}),
    "value": frozenset({"utils", "value"}),
    "parser": frozenset({"utils", "value", "parser"}),
    "model": frozenset({"utils", "value", "model"}),
    "pivot": frozenset({"utils", "value", "model", "pivot"}),
    "eval": frozenset({"utils", "value", "parser", "model", "pivot", "eval"}),
    "cf": frozenset({"utils", "value", "parser", "model", "eval", "cf"}),
    "print": frozenset({"utils", "value", "model", "print"}),
    "io": frozenset({"utils", "value", "parser", "model", "pivot", "io"}),
    "drawing": frozenset({"utils", "value", "model", "print", "io", "drawing"}),
    "facade": frozenset(
        {"utils", "value", "parser", "model", "pivot", "eval", "cf", "print", "io", "drawing", "facade"}
    ),
    "app": frozenset(
        {"utils", "value", "parser", "model", "pivot", "eval", "cf", "print", "io", "drawing", "facade", "app"}
    ),
}

# Groups in which the `io::` namespace token is banned outside comments.
LAYER_NO_IO = frozenset({"value", "parser", "model", "pivot", "eval", "cf", "print"})

# Top-level files that belong to the value layer; every other top-level file
# except the facade files is model. The Excel profile and locale tables are
# leaf data the parser reads, so they sit beside Value.
_VALUE_FILES = ("value.", "sheet_name", "phonetic.h", "value_sort_order.h", "excel_profile.", "excel_locale.")

# Top-level files that implement the `Workbook` facade.
_FACADE_FILES = frozenset(
    {
        "workbook.cpp",
        "workbook_internal.h",
        "workbook_defined_names.cpp",
        "workbook_error_detail.cpp",
        "workbook_external_links.cpp",
        "workbook_hyperlinks.cpp",
        "workbook_merges_arrays.cpp",
        "workbook_formula_index.h",
        "workbook_formula_index.cpp",
        "workbook_ref_rewrite.h",
        "workbook_ref_rewrite.cpp",
        "workbook_copy_sheet.cpp",
        "workbook_data_table.cpp",
        "workbook_pivot_blocked.cpp",
        "workbook_pivot_details.cpp",
        "workbook_pivot_projection.cpp",
        "workbook_pivot_refresh.cpp",
        "workbook_row_col_edit.cpp",
        "workbook_tables.cpp",
        "workbook_theme_package.h",
        "workbook_theme_package.cpp",
        "workbook_table_insert.cpp",
        "workbook_table_header.cpp",
        "workbook_table_loss.h",
        "workbook_table_loss.cpp",
        "table_calc_column.h",
        "table_calc_column.cpp",
        "table_totals_sync.h",
        "table_totals_sync.cpp",
        "workbook_sheet_mutation.h",
        "workbook_sheets.cpp",
        "workbook_auto_filter.cpp",
        "workbook_cell_edit.cpp",
    }
)

# Headers under `pivot/` and `cf/` that are model types rather than engine.
_MODEL_HEADERS = frozenset(
    {
        "pivot/pivot_types.h",
        "pivot/pivot_table.h",
        "pivot/pivot_cache.h",
        "pivot/pivot_result.h",
        "cf/cf_types.h",
        "cf/cf_x14_id.h",
    }
)

_APP_DIRS = frozenset({"c_api", "cli", "wasm", "node_addon"})

# (includer, included), both relative to `src/`, mapped to the reason the
# upward include stays.
LAYER_BASELINE: dict[tuple[str, str], str] = {
    ("value.cpp", "eval/lambda_value.h"): "Value owns the LAMBDA payload.",
    ("pivot/value_order.h", "eval/collation.h"): "Pivot item ordering shares the eval text collation.",
    ("pivot/value_order.h", "eval/eval_profile_scope.h"): "Pivot collation reads the evaluating profile.",
    ("pivot/filter_engine.cpp", "eval/collation.h"): "Label filters collate like the items they filter.",
    ("pivot/pivot_locale.cpp", "eval/collation.h"): "Default data-field names collide under the profile collation.",
    (
        "pivot/field_patch.cpp",
        "eval/collation.h",
    ): "A renamed field or item collides with a sibling under the profile collation.",
    ("pivot/filter_engine.cpp", "eval/collated_wildcard.h"): "Caption begins/ends/contains match like SEARCH.",
    ("pivot/filter_engine.h", "eval/collated_wildcard.h"): "The record filter keeps one caption matcher per filter.",
    ("pivot/filter_engine.cpp", "eval/eval_profile_scope.h"): "Pivot collation reads the evaluating profile.",
    ("pivot/pivot_evaluator.cpp", "eval/eval_profile_scope.h"): "Pivot evaluation scopes the collation profile.",
    ("pivot/hierarchy_builder.cpp", "eval/text_format/display_text.h"): (
        "Numeric axis items are labelled in their field's number format, like the cells Excel draws."
    ),
    ("pivot/field_grouping.cpp", "eval/text_format/display_text.h"): (
        "Date group items show their bounds as the locale's short date, as Excel writes them."
    ),
    ("pivot/pivot_refresh.cpp", "eval/text_format/display_text.h"): (
        "A refreshed field is named by its header cell's displayed text, as Excel names it."
    ),
    ("pivot/pivot_source.cpp", "eval/a1_parse.h"): (
        "A worksheet source ref is A1 text read with the grammar INDIRECT and names use."
    ),
    ("io/ooxml_writer.cpp", "eval/eval_profile_scope.h"): "Pivot write plans collate shared items like evaluation.",
    ("sheet_name.cpp", "eval/collation.h"): "Name identity is the measured Mac '=' text collation equality.",
    ("cell_style_save_plan.cpp", "eval/collation.h"): "Saved cell styles merge and sort under the profile collation.",
}

_INCLUDE_RE = re.compile(r'^\s*#\s*include\s+"([^"]+)"')
_IO_TOKEN_RE = re.compile(r"\bio::")


def _layer_of(rel: Path) -> str | None:
    """Returns the layer group of a path relative to `src/`, or None."""
    top = rel.parts[0]
    if len(rel.parts) == 1:
        if rel.name in _FACADE_FILES:
            return "facade"
        return "value" if rel.name.startswith(_VALUE_FILES) else "model"
    if top in ("pivot", "cf"):
        if rel.as_posix() in _MODEL_HEADERS:
            return "model"
        return "pivot" if top == "pivot" else "cf"
    if top in _APP_DIRS:
        return "app"
    return top if top in LAYER_ALLOWED else None


def _blank_comments_and_strings(text: str) -> str:
    """Replaces comments and string/char literal bodies with spaces, keeping newlines."""
    out: list[str] = []
    i, n = 0, len(text)
    while i < n:
        two = text[i : i + 2]
        if two == "//":
            while i < n and text[i] != "\n":
                out.append(" ")
                i += 1
        elif two == "/*":
            end = text.find("*/", i + 2)
            end = n if end < 0 else end + 2
            out.extend(c if c == "\n" else " " for c in text[i:end])
            i = end
        elif text[i] in "\"'":
            quote = text[i]
            out.append(quote)
            i += 1
            while i < n and text[i] != quote and text[i] != "\n":
                if text[i] == "\\":
                    out.append(" ")
                    i += 1
                out.append(" ")
                i += 1
            if i < n and text[i] == quote:
                out.append(quote)
                i += 1
        else:
            out.append(text[i])
            i += 1
    return "".join(out)


def _resolve_include(includer: Path, spelled: str) -> Path | None:
    """Resolves a quoted include as the compiler does: beside the includer first, then from src/."""
    for base in (includer.parent, SRC_ROOT):
        candidate = Path(os.path.normpath(base / spelled))
        if candidate.is_file() and SRC_ROOT in candidate.parents:
            return candidate.relative_to(SRC_ROOT)
    return None


def check_layers() -> bool:
    hits: list[str] = []
    for path in sorted(SRC_ROOT.rglob("*")):
        if path.suffix not in SOURCE_SUFFIXES:
            continue
        rel = path.relative_to(SRC_ROOT)
        group = _layer_of(rel)
        if group is None:
            continue
        text = path.read_text(encoding="utf-8")
        for lineno, line in enumerate(text.splitlines(), start=1):
            match = _INCLUDE_RE.match(line)
            if not match:
                continue
            target = _resolve_include(path, match.group(1))
            if target is None:
                continue
            target_group = _layer_of(target)
            if target_group is None or target_group in LAYER_ALLOWED[group]:
                continue
            if (rel.as_posix(), target.as_posix()) in LAYER_BASELINE:
                continue
            hits.append(f"src/{rel}:{lineno}: {group} includes {target.as_posix()} ({target_group})")
        if group in LAYER_NO_IO:
            code = _blank_comments_and_strings(text)
            for lineno, line in enumerate(code.splitlines(), start=1):
                if _IO_TOKEN_RE.search(line):
                    hits.append(f"src/{rel}:{lineno}: {group} names the io:: namespace")
    return _report(
        "layers",
        hits,
        "Include only downward (utils < value < parser < model < pivot < eval < cf/print; io beside eval).\n"
        "Move the shared type into the lower layer instead of including upward; a\n"
        "deliberate exception goes in LAYER_BASELINE with its reason.",
    )


def check_array_alloc() -> bool:
    seam = SRC_ROOT / "eval" / "array_alloc.cpp"
    hits = _scan(re.compile(r"create\s*<\s*ArrayValue\s*>"), {seam})
    return _report(
        "array-alloc",
        hits,
        "Call `eval::allocate_array_value(rows, cols, arena, out_buffer, max_cells)`\n"
        f"instead; only {seam.relative_to(REPO_ROOT)} may allocate the header directly.",
    )


def check_xml_passthrough() -> bool:
    seam = SRC_ROOT / "io" / "xml_utils.cpp"
    hits = _scan(re.compile(r"\bformat_raw\b"), {seam})
    return _report(
        "xml-passthrough",
        hits,
        "Call `io::append_raw_xml` / `io::raw_xml` for one element, or\n"
        "`io::capture_unknown_children` for the unmodelled children of a parent;\n"
        f"only {seam.relative_to(REPO_ROOT)} may drive a pugixml writer directly.",
    )


LIMIT_TAGS = frozenset({"engine", "excel", "format", "internal"})

# Facade sources that define limit constants outside src/.
_LIMIT_JS_FILES = ("packages/npm/common.mjs",)
_LIMIT_PY_DIR = "packages/python/formulon"

_LIMIT_CPP_DEF_RE = re.compile(r"^(?:(?:inline )?constexpr\b|  static constexpr\b).*?\b(kMax\w+)\s*=")
_LIMIT_JS_DEF_RE = re.compile(r"^(?:export )?const\s+(MAX_\w+)\s*=")
_LIMIT_PY_DEF_RE = re.compile(r"^(_\w*_CAP)\s*=")
_LIMIT_TAG_RE = re.compile(r"(?://|#)\s*@limit:\s*(\S*)")
_LIMIT_CPP_NAME_RE = re.compile(r"\bkMax\w+\b")
_LIMIT_JS_NAME_RE = re.compile(r"\bMAX_\w+\b")
_LIMIT_PY_NAME_RE = re.compile(r"\b_\w*_CAP\b")
_LIMIT_ERROR_HELPER_RE = re.compile(r"\b(?:make_error|invalid_argument|invalid|set_binding_error|group_error)\s*\(")
_LIMIT_RESOURCE_CODE_RE = re.compile(r"\bkSecurityResourceLimit\b|\bFM_STATUS_SECURITY_RESOURCE_LIMIT\b")
_LIMIT_FACADE_CODE_RE = re.compile(r"StatusCode\.(?:SecurityResourceLimit|SECURITY_RESOURCE_LIMIT)\b")
_LIMIT_BUDGET_RE = re.compile(r"\bResourceBudget(?:\s+\w+)?\s*[({]")
_LIMIT_CPP_WINDOW = 6
_LIMIT_FACADE_WINDOW = 8


def _limit_tag_for(lines: list[str], idx: int) -> tuple[str | None, bool]:
    """Returns (tag value, present) from the definition line or the one above."""
    above = lines[idx - 1].lstrip() if idx > 0 else ""
    for text in (lines[idx], above if above.startswith(("//", "#")) else ""):
        match = _LIMIT_TAG_RE.search(text)
        if match:
            return match.group(1), True
    return None, False


def _limit_call_args(code: list[str], idx: int, start: int) -> tuple[str, int]:
    """Returns the text between the bracket at (idx, start) and its match, and the arg count."""
    text = "\n".join(code[idx : idx + 4])[start:]
    depth, commas, body = 0, 0, []
    for ch in text:
        if ch in "({[":
            depth += 1
            if depth == 1:
                continue
        elif ch in ")}]":
            depth -= 1
            if depth == 0:
                break
        elif ch == "," and depth == 1:
            commas += 1
        body.append(ch)
    return "".join(body), commas + 1


def limit_origin_violations(sources: dict[str, str]) -> list[str]:
    """Checks the origin-tag rules over `{relative path: text}`; the language follows the suffix."""
    kinds: dict[str, str] = {}
    for rel in sources:
        suffix = Path(rel).suffix
        kinds[rel] = "cpp" if suffix in SOURCE_SUFFIXES else "js" if suffix == ".mjs" else "py"
    def_res = {"cpp": _LIMIT_CPP_DEF_RE, "js": _LIMIT_JS_DEF_RE, "py": _LIMIT_PY_DEF_RE}
    name_res = {"cpp": _LIMIT_CPP_NAME_RE, "js": _LIMIT_JS_NAME_RE, "py": _LIMIT_PY_NAME_RE}

    hits: list[str] = []
    raw: dict[str, list[str]] = {}
    local: dict[str, dict[str, str]] = {}
    shared: dict[str, str] = {}
    definitions: set[tuple[str, int]] = set()
    for rel, text in sources.items():
        raw[rel] = text.splitlines()
        local[rel] = {}
        for idx, line in enumerate(raw[rel]):
            match = def_res[kinds[rel]].match(line)
            if not match:
                continue
            definitions.add((rel, idx))
            name = match.group(1)
            tag, present = _limit_tag_for(raw[rel], idx)
            if not present:
                hits.append(f"{rel}:{idx + 1}: {name} has no `@limit:` tag")
            elif tag not in LIMIT_TAGS:
                hits.append(f"{rel}:{idx + 1}: {name} has unknown `@limit: {tag}` (use engine|excel|format|internal)")
            else:
                local[rel][name] = tag
                shared.setdefault(name, tag)

    for rel, lines in raw.items():
        kind = kinds[rel]
        code = _blank_comments_and_strings("\n".join(lines)).splitlines() if kind == "cpp" else lines
        for idx, line in enumerate(code):
            if (rel, idx) in definitions or (kind != "cpp" and line.lstrip().startswith(("//", "#"))):
                continue
            for name in dict.fromkeys(name_res[kind].findall(line)):
                tag = local[rel].get(name) or shared.get(name)
                if tag not in ("engine", "excel", "format"):
                    continue
                where = f"{rel}:{idx + 1}: {name}"
                esc = re.escape(name)
                comparison = re.search(rf"(?<![<>\-])[<>]=?\s*{esc}\b|\b{esc}\s*(?:[<>]=?)(?![<>])", line)
                if kind != "cpp":
                    window = "\n".join(lines[idx : idx + _LIMIT_FACADE_WINDOW + 1])
                    if tag == "engine" and comparison and not _LIMIT_FACADE_CODE_RE.search(window):
                        hits.append(
                            f"{where} (engine) is compared without StatusCode.SecurityResourceLimit "
                            f"within {_LIMIT_FACADE_WINDOW} lines"
                        )
                    continue
                clamp = re.search(
                    rf"\b(?:std::)?(?:min|clamp)\s*(?:<[^>()]*>)?\s*\([^;]*\b{esc}\b|\.resize\(\s*{esc}\b", line
                )
                budget = _LIMIT_BUDGET_RE.search(line)
                if budget and name not in line[budget.end() - 1 :]:
                    budget = None
                window = "\n".join(code[idx : idx + _LIMIT_CPP_WINDOW + 1])
                if tag == "engine":
                    if clamp:
                        hits.append(f"{where} (engine) clamps instead of refusing with kSecurityResourceLimit")
                    if budget:
                        _, argc = _limit_call_args(code, idx, budget.end() - 1)
                        if argc > 2:
                            hits.append(f"{where} (engine) is a ResourceBudget ceiling with a third argument")
                    elif comparison and not clamp:
                        helper = None
                        for j in range(idx, min(idx + _LIMIT_CPP_WINDOW + 1, len(code))):
                            helper = _LIMIT_ERROR_HELPER_RE.search(code[j])
                            if helper:
                                call = "\n".join(code[j : j + 4])[helper.start() :].split(";", 1)[0]
                                if not _LIMIT_RESOURCE_CODE_RE.search(call):
                                    hits.append(
                                        f"{where} (engine) is followed by a refusal other than "
                                        f"kSecurityResourceLimit (line {j + 1})"
                                    )
                                break
                elif (comparison or budget or clamp) and _LIMIT_RESOURCE_CODE_RE.search(window):
                    hits.append(f"{where} ({tag}) is followed by kSecurityResourceLimit")
    return hits


def check_limit_origin() -> bool:
    sources: dict[str, str] = {}
    for path in sorted(SRC_ROOT.rglob("*")):
        if path.suffix in SOURCE_SUFFIXES:
            sources[path.relative_to(REPO_ROOT).as_posix()] = path.read_text(encoding="utf-8")
    for rel in _LIMIT_JS_FILES:
        sources[rel] = (REPO_ROOT / rel).read_text(encoding="utf-8")
    for path in sorted((REPO_ROOT / _LIMIT_PY_DIR).glob("*.py")):
        sources[path.relative_to(REPO_ROOT).as_posix()] = path.read_text(encoding="utf-8")
    return _report(
        "limit-origin",
        limit_origin_violations(sources),
        "Tag every ceiling `// @limit: engine|excel|format|internal` (`# @limit:` in Python).\n"
        "engine: refused with kSecurityResourceLimit, never clamped; excel/format: never\n"
        "kSecurityResourceLimit. Fix the code or the tag; there is no allow-list.",
    )


CHECKS = {
    "array-alloc": check_array_alloc,
    "xml-passthrough": check_xml_passthrough,
    "layers": check_layers,
    "limit-origin": check_limit_origin,
}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("check", nargs="?", default="all", choices=[*CHECKS, "all"])
    args = parser.parse_args()

    selected = CHECKS.values() if args.check == "all" else [CHECKS[args.check]]
    return 0 if all([fn() for fn in selected]) else 1


if __name__ == "__main__":
    sys.exit(main())
