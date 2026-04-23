// Copyright 2026 libraz. Licensed under the MIT License.
//
// Shared declarations for the tree-walk evaluator's "lazy" function impls.
//
// A lazy impl receives the un-evaluated `Call` AST node and decides which
// of its argument subtrees actually need to be evaluated. This is how
// short-circuit semantics (`IF` / `IFERROR` / `IFNA`), the conditional
// aggregators (`*IF` / `*IFS`), and the range-aware lookups (`VLOOKUP`,
// `INDEX`, `MATCH`, `XLOOKUP`, ...) get expressed: the eager dispatch
// path in `tree_walker.cpp` flattens every argument to a `Value` before
// calling the impl, which would defeat short-circuiting AND would erase
// the range/Ref shape needed by the lookups.
//
// This header exists as a seam so each family of lazy impls can live in
// its own translation unit while the central dispatch table in
// `tree_walker.cpp` keeps a single source of truth for name-to-impl
// routing. It publishes only the shared vocabulary: the `LazyImpl`
// function-pointer signature and the `eval_node` recursion entry point.
// Each family of lazy impls declares its own externs in a sibling header
// (`special_forms_lazy.h`, `conditional_aggregates.h`, and the
// forthcoming lookup headers), and the dispatch table in
// `tree_walker.cpp` includes all of them.

#ifndef FORMULON_EVAL_LAZY_IMPLS_H_
#define FORMULON_EVAL_LAZY_IMPLS_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Recursive entry point of the tree-walk evaluator. Defined in
/// `tree_walker.cpp`; published here so lazy-impl translation units can
/// evaluate selected argument subtrees on demand.
Value eval_node(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx);

/// Signature of every lazy (short-circuit / range-aware) function impl.
/// The central dispatch table in `tree_walker.cpp` stores pointers of
/// this type.
using LazyImpl = Value (*)(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LAZY_IMPLS_H_
