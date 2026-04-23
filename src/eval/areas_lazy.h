// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impl for the `AREAS` built-in. `AREAS(reference)` counts the number
// of disjoint rectangles in its single reference argument. The count is an
// inherently structural property of the AST: evaluating the argument to a
// scalar `Value` first would flatten a `(A1:B2, C3:D4)` union into an
// indistinguishable error value, so this impl peeks at the un-evaluated
// argument shape directly.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// dispatch-table contract in `tree_walker.cpp`.

#ifndef FORMULON_EVAL_AREAS_LAZY_H_
#define FORMULON_EVAL_AREAS_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `AREAS(reference)` - returns the number of disjoint rectangles in the
/// `reference` argument. `AREAS(A1)` and `AREAS(A1:B2)` both return 1;
/// a parenthesised comma-union such as `AREAS((A1:B2, C3:D4))` returns 2.
/// Any non-reference argument (a literal, an arithmetic expression, a
/// function call) surfaces `#VALUE!`. The impl is lazy so the argument's
/// AST shape (specifically `UnionOp` / `RangeOp` / `Ref`) is preserved
/// through dispatch.
Value eval_areas_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_AREAS_LAZY_H_
