//
// Lazy impls for the dynamic-array "indexing" family: `CHOOSECOLS`,
// `CHOOSEROWS`, `TAKE`, `DROP`. These builtins return a sub-array
// addressed either by enumerated 1-based / negative axis indices
// (CHOOSECOLS / CHOOSEROWS) or by signed edge-counted slice arguments
// (TAKE / DROP). They share `resolve_choose_index`, `resolve_take_drop_range`,
// `materialise_selected_lanes`, and `materialise_slice` with the rest of
// the dynamic-array family (see `dynamic_array/common.h`).
//
// They are lazy (rather than eager like ordinary builtins) because the
// array argument must reach the impl as an AST node so a bare single-cell
// `Ref` can still be treated as a 1-cell range. The central dispatch
// table in `tree_walker_lazy_table.cpp` references these externs by
// unqualified name; see `eval/lazy_impls.h` for the shared `LazyImpl`
// signature and `eval_node` entry point.

#ifndef FORMULON_EVAL_DYNAMIC_ARRAY_INDEXING_H_
#define FORMULON_EVAL_DYNAMIC_ARRAY_INDEXING_H_

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

Value eval_choosecols_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);
Value eval_chooserows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);
Value eval_take_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);
Value eval_drop_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

// Compile-time guard: every lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`. Picking `eval_choosecols_lazy` as a witness is
// sufficient because every sibling shares the same parameter list.
inline constexpr LazyImpl kDynamicArrayIndexingLazySignatureWitness = &eval_choosecols_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DYNAMIC_ARRAY_INDEXING_H_
