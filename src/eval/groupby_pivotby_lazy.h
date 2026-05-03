// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lazy impls for Excel 365's GROUPBY (and, in a follow-up, PIVOTBY) dynamic-
// array functions. Both produce a 2D spilled result that summarises a tabular
// input by one or more grouping keys, applying a user-supplied aggregator
// function per group.
//
// GROUPBY signature (Mac Excel 365):
//
//   GROUPBY(row_fields, values, function,
//           [field_headers=0], [total_depth=-1], [sort_order=0],
//           [filter_array])
//
//   * `row_fields` is a 1D or 2D rectangle whose rows are the group keys.
//     A multi-column shape produces composite keys (column-major tuple).
//   * `values` is a 1D or 2D rectangle whose columns are aggregated
//     independently; the output preserves one aggregated column per input
//     column.
//   * `function` is one of three accepted shapes:
//       - Form A: an inline `LAMBDA(v, ...)` literal of arity 1.
//       - Form B: a name-bound LAMBDA value from a surrounding `LET`.
//       - Form C: a bare function name (`SUM`, `AVERAGE`, `COUNT`, ...) —
//         resolved against the eval-time `FunctionRegistry` and invoked with
//         the group's column slice as a single Value-array argument. No
//         synthetic LambdaValue is constructed (the bare name has no AST
//         body), so the per-group invocation path branches on whether the
//         resolved aggregator is a Lambda value or a registry-defined
//         function.
//   * `field_headers` ∈ {0,1,2,3} controls header handling (0=none,
//     1=inputs only, 2=synthesise output, 3=both).
//   * `total_depth` ∈ {-2,-1,0,1,2} controls grand-total + subtotal
//     placement. The sign governs position (negative = above, positive =
//     below). Subtotals (±2) only meaningful when `row_fields` has >= 2
//     columns; with single-column keys ±2 silently degrade to ±1.
//   * `sort_order` is 0 (preserve first-occurrence order), N>0 (sort
//     ascending by N-th aggregated value column, 1-based), or N<0
//     (descending by |N|-th column).
//   * `filter_array`, when supplied, is a column-shaped boolean mask over
//     the data rows (header row excluded if `field_headers` ∈ {1,3}). TRUE
//     keeps the row; FALSE drops it.
//
// Per-group error isolation: unlike MAP / BYROW, an error returned by the
// aggregator for ONE group does NOT abort the entire result. The error is
// stored verbatim in that group's output cell and the remaining groups are
// still computed. Errors elsewhere (in `function`, `field_headers`, the
// filter mask cells, etc.) follow the standard short-circuit rules.
//
// Group-key equality:
//   * Text keys are compared after `jp_fold` normalisation, matching Mac
//     Excel ja-JP COUNTIF / VLOOKUP behaviour (half-width katakana folds to
//     full-width, hiragana folds to katakana, ...).
//   * Numeric keys compare bit-exact via `==`.
//   * Blank cells form their own "blank" group.
//   * Errors in `row_fields` cells become the row's group key (the error
//     itself); error groups sort after all valid keys.
//
// PIVOTBY (signature: row_fields, col_fields, values, function, ...) is the
// 2D analogue of GROUPBY. It introduces a column-grouping axis: `col_fields`
// drives a horizontal grouping in addition to the vertical grouping driven
// by `row_fields`, and the output is a 2D pivot whose body cells are
// per-(row-group, col-group) aggregates. PIVOTBY shares the aggregator
// resolution, group-key equality, per-group invocation, and sort/total
// machinery with GROUPBY (the helpers live in this TU's anonymous namespace
// and are reused verbatim).
//
// Differences vs. GROUPBY argument defaults:
//   * `field_headers` defaults to `3` (vs. `0`) — pivot output typically
//     wants both the input row to be treated as a header AND a header to
//     be emitted on the output's left/top edges.
//   * `col_total_depth` defaults to `1` (vs. row_total_depth's `-1`) —
//     grand totals on rows live at the top by default, on columns they
//     live on the right by default.
//
// First-pass scope (this commit): single-column `row_fields`, single-column
// `col_fields`, single-column `values`, `row_total_depth ∈ {-1, 0, 1}` and
// `col_total_depth ∈ {-1, 0, 1}` (subtotal placement at ±2 is deferred
// until the GROUPBY oracle pass settles), and `row_sort_order` /
// `col_sort_order` ∈ {-1, 0, 1}. Multi-column inputs surface `#VALUE!`.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// `eval_node` recursion entry point. The GROUPBY impl is wired into the
// central `kLazyDispatch` table in `tree_walker.cpp` between `FREQUENCY`
// and `GROWTH` (alphabetical); PIVOTBY between `PERCENTRANK.INC` and
// `PROB`.

#ifndef FORMULON_EVAL_GROUPBY_PIVOTBY_LAZY_H_
#define FORMULON_EVAL_GROUPBY_PIVOTBY_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `GROUPBY(row_fields, values, function, [field_headers], [total_depth],
///          [sort_order], [filter_array])` — see the header preamble for the
/// full contract.
Value eval_groupby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

/// `PIVOTBY(row_fields, col_fields, values, function, [field_headers=3],
///          [row_total_depth=-1], [row_sort_order=0], [col_total_depth=1],
///          [col_sort_order=0], [filter_array])` — see the header preamble
/// for the full contract.
Value eval_pivotby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_GROUPBY_PIVOTBY_LAZY_H_
