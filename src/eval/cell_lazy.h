// Copyright 2026 libraz. Licensed under the MIT License.
//
// MVP implementation of Excel's `CELL(info_type, [reference])` function.
//
// CELL is registered as a lazy builtin because it must inspect the
// un-evaluated AST of its optional `reference` argument: when the
// reference is omitted the function answers about the formula cell
// itself (read off `EvalContext::formula_row()` / `formula_col()`); when
// a multi-cell range is supplied only the top-left cell is consulted.
// The eager dispatcher would flatten both shapes to a `Value` before
// the impl runs.
//
// Coverage in this MVP:
//   * "address", "col", "row", "contents", "type" - fully implemented
//     against the bound cell or the reference's top-left.
//   * "filename", "format", "color", "parentheses", "prefix", "protect",
//     "width" - return safe fixed stubs because the engine does not yet
//     carry style / format / column-width / lock metadata. Each stub is
//     called out at its return site with a `// no <metadata> yet` comment
//     so it is unambiguous that the value is intentional, not a bug.
//
// info_type matching is ASCII case-insensitive (Excel folds the key
// before lookup). Errors in info_type / reference propagate; an empty
// or unknown info_type surfaces `#VALUE!`.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// dispatch-table contract in `tree_walker.cpp`.

#ifndef FORMULON_EVAL_CELL_LAZY_H_
#define FORMULON_EVAL_CELL_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `CELL(info_type, [reference])` - returns various pieces of metadata
/// about the top-left cell of `reference` (or the formula cell itself
/// when `reference` is omitted).
///
/// Supported `info_type` keys (ASCII case-insensitive):
///   * `"address"`   - A1 absolute address with `$` anchors. Cross-sheet
///                     references include the sheet prefix
///                     (`Sheet1!$A$1`). The `[Workbook]` filename wrapper
///                     is intentionally omitted because Workbook has no
///                     path field.
///   * `"col"`       - 1-based column of the top-left cell.
///   * `"row"`       - 1-based row of the top-left cell.
///   * `"contents"`  - top-left cell's value. Blank cells return `0`
///                     (matching Excel's blank-as-zero rule); all other
///                     shapes (Number / Bool / Text / Error) pass
///                     through unchanged.
///   * `"type"`      - `"b"` for blank cells or an empty-string text
///                     value (Mac folds `""` to `"b"`), `"l"` for any
///                     non-empty text, `"v"` for everything else
///                     (number, bool, error).
///   * `"filename"`  - always `""` (Mac returns blank); the workbook has
///                     no filesystem path. Empty text avoids the
///                     top-level blank-as-zero collapse; the divergence
///                     is skip-oracle until the workbook gains a path
///                     field.
///   * `"format"`    - always `"G"`; no style subsystem yet.
///   * `"color"`     - always `0`; no negative-number color flag yet.
///   * `"parentheses"` - always `0`; no parenthesis-format flag yet.
///   * `"prefix"`    - always `""` (Mac returns blank); no text-alignment
///                     metadata yet. Empty text avoids the top-level
///                     blank-as-zero collapse; skip-oracle until the
///                     style subsystem lands.
///   * `"protect"`   - always `1`; cells are default-locked until the
///                     style subsystem lands.
///   * `"width"`     - always a 1x2 array `{8, TRUE}`; no column-width
///                     metadata yet.
///
/// Error / arity rules:
///   * Unknown info_type (after lowercase fold) -> `#VALUE!`.
///   * Empty info_type string -> `#VALUE!`.
///   * info_type is an Error -> propagate (via `coerce_to_text`).
///   * info_type is not text-coercible (e.g. Array) -> `#VALUE!`.
///   * `reference` evaluates to an error -> propagate that error.
///   * Wrong arity (0 or >=3) -> `#VALUE!`.
///   * `reference` omitted and no formula cell anchor on the context
///     (e.g. ad-hoc CLI eval) -> `#REF!` defensively for the
///     "address" / "col" / "row" / "contents" / "type" keys.
Value eval_cell_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_CELL_LAZY_H_
