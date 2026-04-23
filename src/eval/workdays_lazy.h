// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the workday-arithmetic family: `NETWORKDAYS` and
// `WORKDAY`. Both accept an optional trailing `holidays` argument that
// may be a range/Ref (e.g. `A1:A10`) or an inline array literal
// (`{45292,45293}`); the eager dispatcher would erase the per-argument
// AST shape by flattening every arg to a single `Value`, which is why
// these two impls live on the lazy-dispatch seam alongside SUMPRODUCT /
// INDEX / MATCH.
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

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_WORKDAYS_LAZY_H_
