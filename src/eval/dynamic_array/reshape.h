//
// Lazy impls for the dynamic-array "reshape" family: `HSTACK`, `VSTACK`,
// `EXPAND`. These builtins assemble an `ArrayValue` whose footprint is
// derived from the *combined* shapes of multiple inputs (HSTACK / VSTACK)
// or from explicit caller-supplied dimensions (EXPAND), padding any
// uncovered cells with `#N/A` or the caller-supplied scalar.
//
// `TRANSPOSE` is also a dynamic-array reshape function but it lives in
// `src/eval/shape_ops_lazy.{h,cpp}` together with the SUMPRODUCT-side
// helpers it pre-dates, so this header does NOT declare it.
//
// They are lazy (rather than eager like ordinary builtins) because every
// stack input may be a range-shaped argument whose 2D shape the eager
// dispatcher would flatten away, and because the EXPAND `pad_with`
// argument must not be evaluated unless the impl needs it. The central
// dispatch table in `tree_walker_lazy_table.cpp` references these
// externs by unqualified name; see `eval/lazy_impls.h` for the shared
// `LazyImpl` signature and `eval_node` entry point.

#ifndef FORMULON_EVAL_DYNAMIC_ARRAY_RESHAPE_H_
#define FORMULON_EVAL_DYNAMIC_ARRAY_RESHAPE_H_

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

Value eval_hstack_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);
Value eval_vstack_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);
Value eval_expand_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

// Compile-time guard: every lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`. Picking `eval_hstack_lazy` as a witness is
// sufficient because every sibling shares the same parameter list.
inline constexpr LazyImpl kDynamicArrayReshapeLazySignatureWitness = &eval_hstack_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DYNAMIC_ARRAY_RESHAPE_H_
