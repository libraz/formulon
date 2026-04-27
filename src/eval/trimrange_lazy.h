// Copyright 2026 libraz. Licensed under the MIT License.
//
// TRIMRANGE — trims fully-blank leading / trailing rows and columns from a
// rectangular reference, range, or array literal. The Excel 365 dynamic-array
// helper used to harvest a "tight" sub-rectangle from a sparsely populated
// region.
//
// `TRIMRANGE(range, [trim_rows]=3, [trim_cols]=3)`
//
// Trim modes (independent for the row and column axes):
//   * 0 -> trim no edges
//   * 1 -> trim only the leading edge (top rows / left columns)
//   * 2 -> trim only the trailing edge (bottom rows / right columns)
//   * 3 -> trim both edges (DEFAULT)
//
// What counts as blank: only the `Blank` value variant. The empty string
// (`""`), whitespace-only text (`" "`), the number `0`, the boolean `FALSE`,
// and any error value (e.g. `#N/A`) are NOT trimmable. This matches Mac
// Excel's observed behaviour and mirrors the `skip_blanks` path in TOCOL /
// TOROW (see `eval/dynamic_array_lazy.cpp`).
//
// A row is considered trimmable iff every cell in it is `Blank`. The same
// rule applies to columns. Each axis is scanned from the appropriate edge
// inward and stops at the first non-trimmable row / column. The output is a
// contiguous sub-rectangle of the input.
//
// Output shape: always an `ArrayValue`. A single-cell scalar input degenerates
// to a 1x1 array pass-through. If trimming would consume every row or every
// column the call surfaces `#CALC!` (Mac Excel: an empty spill is rejected).

#ifndef FORMULON_EVAL_TRIMRANGE_LAZY_H_
#define FORMULON_EVAL_TRIMRANGE_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `TRIMRANGE(range, [trim_rows]=3, [trim_cols]=3)` — returns the sub-rectangle
/// of `range` that remains after dropping fully-blank leading and / or
/// trailing rows and columns, controlled independently per axis by
/// `trim_rows` and `trim_cols`.
///
/// Arguments:
///   * `range` — any expression evaluable in array context. Single-cell
///     scalars are treated as 1x1 arrays. Errors propagate verbatim.
///   * `trim_rows` — integer in `{0, 1, 2, 3}` (optional, default `3`).
///     `0` keeps every row, `1` trims only leading blank rows, `2` trims
///     only trailing blank rows, `3` trims both.
///   * `trim_cols` — same shape as `trim_rows` (optional, default `3`),
///     applied to the column axis independently.
///
/// Errors:
///   * arity outside `[1, 3]` -> `#VALUE!`;
///   * any argument evaluates to an error -> propagate verbatim;
///   * `trim_rows` / `trim_cols` not coercible to a number -> `#VALUE!`;
///   * trim-mode integer outside `{0, 1, 2, 3}` -> `#VALUE!`;
///   * trimming would leave zero rows or zero columns -> `#CALC!`.
Value eval_trimrange_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TRIMRANGE_LAZY_H_
