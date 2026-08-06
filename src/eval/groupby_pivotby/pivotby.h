//
// Lazy impl for Excel 365's PIVOTBY dynamic-array function. PIVOTBY is the
// 2D analogue of GROUPBY: it introduces a column-grouping axis
// (`col_fields`) in addition to the vertical grouping axis (`row_fields`)
// and emits a 2D pivot whose body cells are per-(row-group, col-group)
// aggregates. PIVOTBY shares the aggregator resolution, group-key
// equality, per-group invocation, and sort/total machinery with GROUPBY;
// the shared helpers live in `groupby_pivotby/common.h`.
//
// PIVOTBY signature (Mac Excel 365):
//
//   PIVOTBY(row_fields, col_fields, values, function,
//           [field_headers=3], [row_total_depth=-1], [row_sort_order=0],
//           [col_total_depth=1], [col_sort_order=0], [filter_array])
//
// Differences vs. GROUPBY argument defaults:
//   * `field_headers` defaults to `3` (vs. `0`) — pivot output typically
//     wants both the input row to be treated as a header AND a header to
//     be emitted on the output's left/top edges.
//   * `col_total_depth` defaults to `1` (vs. row_total_depth's `-1`) —
//     grand totals on rows live at the top by default, on columns they
//     live on the right by default.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature. The PIVOTBY
// impl is wired into the central `kLazyDispatch` table in
// `tree_walker_lazy_table.cpp`.

#ifndef FORMULON_EVAL_GROUPBY_PIVOTBY_PIVOTBY_H_
#define FORMULON_EVAL_GROUPBY_PIVOTBY_PIVOTBY_H_

#include "eval/lazy_impls.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `PIVOTBY(row_fields, col_fields, values, function, [field_headers=3],
///          [row_total_depth=-1], [row_sort_order=0], [col_total_depth=1],
///          [col_sort_order=0], [filter_array])` — see the header preamble
/// for the full contract.
Value eval_pivotby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

// Compile-time guard: the lazy impl declared above must convert implicitly
// to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`.
inline constexpr LazyImpl kPivotbyLazySignatureWitness = &eval_pivotby_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_GROUPBY_PIVOTBY_PIVOTBY_H_
