// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the workday-arithmetic family: `NETWORKDAYS`,
// `NETWORKDAYS.INTL`, `WORKDAY`, and `WORKDAY.INTL`. All four accept an
// optional trailing `holidays` argument that may be a range/Ref
// (e.g. `A1:A10`) or an inline array literal (`{45292,45293}`); the eager
// dispatcher would erase the per-argument AST shape by flattening every
// arg to a single `Value`, which is why these impls live on the
// lazy-dispatch seam alongside SUMPRODUCT / INDEX / MATCH. The `.INTL`
// variants additionally take a `weekend` selector (either an integer code
// or a 7-character `01` mask string) that extends the fixed Sat/Sun
// weekend to any customisable pair or single day.
//
// The central dispatch table in `tree_walker.cpp` references these
// externs by unqualified name. See `eval/lazy_impls.h` for the shared
// `LazyImpl` signature and the `eval_node` entry point these impls
// recurse through.

#ifndef FORMULON_EVAL_WORKDAYS_LAZY_H_
#define FORMULON_EVAL_WORKDAYS_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `NETWORKDAYS(start, end, [holidays])` - counts Monday..Friday weekdays
/// in `[min(start, end), max(start, end)]` inclusive, excluding any days
/// that match an entry in `holidays`. `holidays` may be a range, a bare
/// cell reference, or an inline array literal; non-numeric holiday cells
/// are silently ignored. Errors propagate. When `start > end`, Excel
/// returns the negated positive count — this impl mirrors that.
Value eval_networkdays_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx);

/// `WORKDAY(start, days, [holidays])` - returns the serial date that is
/// `days` working days before or after `start`, skipping weekends and any
/// dates in `holidays`. `days` is truncated toward zero; a negative
/// value walks backward. Errors propagate.
Value eval_workday_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

/// `NETWORKDAYS.INTL(start, end, [weekend=1], [holidays])` - like
/// NETWORKDAYS but with a customisable weekend pattern. `weekend` is
/// either an integer code (1..7, 11..17) selecting a pre-defined pattern
/// or a 7-character string of '0' and '1' (position 0 = Monday,
/// 1 = weekend). The all-weekend string "1111111" is rejected as
/// `#VALUE!` to match Excel. An invalid integer selector yields `#NUM!`.
Value eval_networkdays_intl_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx);

/// `WORKDAY.INTL(start, days, [weekend=1], [holidays])` - like WORKDAY
/// with the same custom `weekend` selector as NETWORKDAYS.INTL. `days`
/// is truncated toward zero; a negative value walks backward.
Value eval_workday_intl_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_WORKDAYS_LAZY_H_
