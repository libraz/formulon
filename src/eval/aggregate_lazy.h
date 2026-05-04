// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// AGGREGATE — Excel 365's ergonomic generalization of SUBTOTAL. Dispatches on
// a leading numeric `function_num` (1..19) and accepts a second `options`
// argument that controls which cells inside the data range are filtered out
// before the aggregator runs:
//
//   1  AVERAGE      11  VAR.P
//   2  COUNT        12  MEDIAN
//   3  COUNTA       13  MODE.SNGL
//   4  MAX          14  LARGE          (needs trailing `k`)
//   5  MIN          15  SMALL          (needs trailing `k`)
//   6  PRODUCT      16  PERCENTILE.INC (needs trailing probability)
//   7  STDEV.S      17  QUARTILE.INC   (needs trailing quart in {0..4})
//   8  STDEV.P      18  PERCENTILE.EXC (needs trailing probability)
//   9  SUM          19  QUARTILE.EXC   (needs trailing quart in {1..3})
//  10  VAR.S
//
// The `options` byte (0..7) controls three independent "ignore" bits in
// Excel:
//
//   bit 0 (mask 1)  ignore hidden rows
//   bit 1 (mask 2)  ignore errors
//   bit 2 (mask 4)  reserved by Excel as the "ignore none" sentinel; the
//                   public encoding pairs (0,4), (1,5), (2,6), (3,7) so that
//                   each pair toggles only the nested-SUBTOTAL/AGGREGATE
//                   filter. Codes 0..3 *also* skip nested SUBTOTAL/AGGREGATE
//                   results inside the range; codes 4..7 do not.
//
// Two intentional simplifications relative to Mac Excel 365 (the same
// shortfall SUBTOTAL has documented in `eval/builtins/subtotal.cpp`):
//
//   * Nested-SUBTOTAL/AGGREGATE filtering: Excel ignores cells whose source
//     formula is itself a SUBTOTAL or AGGREGATE call so a column of
//     subtotals can be aggregated without double-counting. Formulon does
//     not yet expose per-cell formula text from inside a builtin, so this
//     filter is omitted. The nested-call bit is therefore unobservable and
//     options codes 0/4, 1/5, 2/6, 3/7 each collapse to one pair of
//     behaviors.
//   * Hidden-row filtering: codes 1, 3, 5, 7 fold to the same treatment as
//     0, 2, 4, 6 respectively because Formulon has no row-visibility state
//     to consult.
//
// What we *do* honor today is the error-ignore bit (mask 2):
//
//   * options ∈ {2, 3, 6, 7} -> range-sourced error cells are silently
//     dropped before the aggregator runs.
//   * options ∈ {0, 1, 4, 5} -> the first range-sourced error short-
//     circuits to that error (canonical row-major scan order, matching
//     SUBTOTAL).
//
// AGGREGATE is registered through the central `kLazyDispatch` table in
// `tree_walker.cpp` rather than via `FunctionRegistry::register_function`.
// The eager path concatenates every flattened argument into a single
// `args[]` vector before invoking the impl, but for codes 14..19 the LAST
// positional argument is the `k` parameter to LARGE/SMALL/PERCENTILE/
// QUARTILE — and "the last cell of the flattened range vs. the trailing
// `k`" is indistinguishable from the eager view. Lazy dispatch lets us
// peek at the un-flattened AST and route the trailing scalar differently
// from the data range itself. This is the same architectural reason
// PERCENTOF and SUMPRODUCT live on the lazy path.

#ifndef FORMULON_EVAL_AGGREGATE_LAZY_H_
#define FORMULON_EVAL_AGGREGATE_LAZY_H_

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

/// `AGGREGATE(function_num, options, ref1, [ref2|k], ...)` — multi-mode
/// aggregator with a configurable error / nested-call filter.
///
/// Arguments:
///   * `function_num` (required, 1..19, integer-truncated): selects the
///     underlying aggregator. Values outside [1, 19] yield `#VALUE!`.
///   * `options` (required, 0..7, integer-truncated): ignore flags. Values
///     outside [0, 7] yield `#VALUE!`. Today only the error-ignore bit
///     (mask 2) is observable; see the file header for the rest.
///   * For `function_num` 1..13: the remaining positional args are the
///     data range(s). Each may be a range, a Ref, an inline `{...}` array
///     literal, or a direct scalar. Per-cell provenance: range-sourced
///     non-numeric cells are dropped (except for code 3, COUNTA, which
///     counts every non-blank).
///   * For `function_num` 14..19: exactly one data argument, followed by a
///     trailing scalar `k` / probability / quart. Excess args yield
///     `#VALUE!`. The constraints on `k` are per-mode:
///       - LARGE / SMALL: integer in `[1, n]`; out-of-range yields `#NUM!`.
///       - PERCENTILE.INC: `0 <= p <= 1`; out-of-range yields `#NUM!`.
///       - PERCENTILE.EXC: `0 < p < 1` AND `p*(n+1)` lies in `[1, n]`;
///         either failure yields `#NUM!`.
///       - QUARTILE.INC: integer in `{0, 1, 2, 3, 4}`; otherwise `#NUM!`.
///       - QUARTILE.EXC: integer in `{1, 2, 3}`; otherwise `#NUM!`.
///
/// Errors:
///   * Errors in `function_num` or `options` always propagate verbatim
///     (the options bit governs the data range only, not the metadata).
///   * Errors in the data range obey the error-ignore bit as documented
///     in the file header.
Value eval_aggregate_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx);

// Compile-time guard: `eval_aggregate_lazy` must convert implicitly to
// the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`, otherwise the dispatch table in `tree_walker.cpp`
// would silently break. Witness flagged at the header so a parameter
// drift surfaces here, not five files away.
inline constexpr LazyImpl kAggregateLazySignatureWitness = &eval_aggregate_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_AGGREGATE_LAZY_H_
