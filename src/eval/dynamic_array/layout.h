// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lazy impls for the dynamic-array "layout" family: `TOCOL`, `TOROW`,
// `WRAPROWS`, `WRAPCOLS`. TOCOL / TOROW flatten a 2D array into a single
// column / row with an optional skip-blanks / skip-errors bitmask;
// WRAPROWS / WRAPCOLS perform the inverse, wrapping a 1D vector into a
// 2D rectangle and padding the trailing edge. They share
// `eval_truncated_number_arg`, `allocate_array_value`, and
// `materialise_vector` with the rest of the dynamic-array family (see
// `dynamic_array/common.h`).
//
// They are lazy (rather than eager like ordinary builtins) because the
// vector argument may be a range-shaped expression whose 2D shape the
// eager dispatcher would flatten, and because the WRAPROWS / WRAPCOLS
// `pad_with` argument must not be evaluated unless padding actually
// occurs. The central dispatch table in `tree_walker_lazy_table.cpp`
// references these externs by unqualified name; see `eval/lazy_impls.h`
// for the shared `LazyImpl` signature and `eval_node` entry point.

#ifndef FORMULON_EVAL_DYNAMIC_ARRAY_LAYOUT_H_
#define FORMULON_EVAL_DYNAMIC_ARRAY_LAYOUT_H_

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

Value eval_tocol_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);
Value eval_torow_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);
Value eval_wraprows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);
Value eval_wrapcols_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

// Compile-time guard: every lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`. Picking `eval_tocol_lazy` as a witness is
// sufficient because every sibling shares the same parameter list.
inline constexpr LazyImpl kDynamicArrayLayoutLazySignatureWitness = &eval_tocol_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DYNAMIC_ARRAY_LAYOUT_H_
