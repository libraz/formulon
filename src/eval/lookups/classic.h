// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the "classic" lookup family: `CHOOSE`, `INDEX`, `MATCH`,
// `VLOOKUP`, and `HLOOKUP`. These builtins share the legacy row/column
// scan model (linear walk along a chosen axis with approximate /
// exact-with-wildcard modes) and a single private helper, `lookup_scan`,
// which lives alongside the impls in `lookups/classic.cpp`.
//
// They are lazy (rather than eager like ordinary builtins) because the
// lookup array argument must reach the impl as an AST node so a bare
// single-cell `Ref` can still be treated as a 1-cell range, and because
// `CHOOSE` must select exactly one of its value arguments and leave the
// rest un-evaluated. The central dispatch table in `tree_walker.cpp`
// references these externs by unqualified name; see `eval/lazy_impls.h`
// for the shared `LazyImpl` signature and `eval_node` entry point.

#ifndef FORMULON_EVAL_LOOKUPS_CLASSIC_H_
#define FORMULON_EVAL_LOOKUPS_CLASSIC_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

Value eval_choose_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);
Value eval_index_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);
Value eval_match_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);
Value eval_vlookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);
Value eval_hlookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);
Value eval_lookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LOOKUPS_CLASSIC_H_
