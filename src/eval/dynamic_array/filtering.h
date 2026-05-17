// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lazy impls for the dynamic-array "filtering" family: `FILTER`, `UNIQUE`,
// `SORT`, `SORTBY`. These builtins all consume a single `array` argument
// (a row-major Value::Array carrying the 2D shape) and produce an
// `ArrayValue` whose footprint is a subset / reordering of the input
// along one axis. They share cell- and lane-comparison primitives with
// the rest of the dynamic-array family (see `dynamic_array/common.h`).
//
// They are lazy (rather than eager like ordinary builtins) because the
// array argument must reach the impl as an AST node so a bare single-cell
// `Ref` can still be treated as a 1-cell range, and because the FILTER
// `if_empty` and SORT / SORTBY `sort_order` arguments must not be
// flattened to a Value before the impl decides how to interpret them.
// The central dispatch table in `tree_walker_lazy_table.cpp` references
// these externs by unqualified name; see `eval/lazy_impls.h` for the
// shared `LazyImpl` signature and `eval_node` entry point.

#ifndef FORMULON_EVAL_DYNAMIC_ARRAY_FILTERING_H_
#define FORMULON_EVAL_DYNAMIC_ARRAY_FILTERING_H_

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

Value eval_filter_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);
Value eval_unique_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);
Value eval_sort_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);
Value eval_sortby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

// Compile-time guard: every lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`. Picking `eval_filter_lazy` as a witness is
// sufficient because every sibling shares the same parameter list.
inline constexpr LazyImpl kDynamicArrayFilteringLazySignatureWitness = &eval_filter_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DYNAMIC_ARRAY_FILTERING_H_
