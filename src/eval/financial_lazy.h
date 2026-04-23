// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the financial family. `IRR` is the sole resident: it
// needs the un-flattened AST of its `values` argument so a bare Ref, a
// RangeOp, or an inline ArrayLiteral can each be walked cell-by-cell
// and fed into Newton-Raphson iteration. The eager path in
// `tree_walker::dispatch_call` would flatten every argument to a single
// `Value` before invoking the impl, erasing the range shape required
// here — which is why IRR rides the lazy dispatch seam alongside
// SUMPRODUCT / INDEX / MATCH / NETWORKDAYS.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// `eval_node` entry point this impl recurses through.

#ifndef FORMULON_EVAL_FINANCIAL_LAZY_H_
#define FORMULON_EVAL_FINANCIAL_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `IRR(values, [guess=0.1])` - internal rate of return via Newton-
/// Raphson. `values` must be a range, a single cell reference, or an
/// inline array literal; non-numeric cells inside a range are skipped
/// silently (matching Excel). The sequence must contain at least one
/// positive and one negative cash flow or the call returns `#NUM!`.
/// Iteration uses the classic IRR period-0 convention (first cash flow
/// at time 0), which differs from NPV's period-1 indexing.
Value eval_irr_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FINANCIAL_LAZY_H_
