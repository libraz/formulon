//
// Private seam between the tree-walker's recursive node visitor
// (`tree_walker/walker.cpp`) and the function-call dispatch path
// (`tree_walker/dispatch.cpp`). The two translation units were split out
// of the original monolithic `tree_walker.cpp` to keep compile units
// digestible; this header publishes the two entry points walker.cpp
// needs from dispatch.cpp.
//
// `dispatch_call` is invoked from `eval_node` for every `Call` AST node:
// it handles name-bound lambda dispatch, lazy/special-form routing
// (IF / IFERROR / *IFS / lookups / ...), and the eager-arg path that
// expands range-shaped arguments before forwarding to the registered
// `FunctionDef::impl`.
//
// `invoke_lambda` is shared between the `LambdaCall` AST case (handled
// in walker.cpp) and the name-bound dispatch path (in dispatch.cpp).
//
// This header is internal to the tree-walker family and is not part of
// the public evaluator surface — production callers reach the evaluator
// through `eval/tree_walker.h`.

#ifndef FORMULON_EVAL_TREE_WALKER_DISPATCH_H_
#define FORMULON_EVAL_TREE_WALKER_DISPATCH_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;
struct LambdaValue;

// Special-cased function-call dispatch. Routes `Call` AST nodes through
// (in order):
//   1. Name-bound lambda lookup against `EvalContext::name_env`.
//   2. The lazy dispatch table (`find_lazy_impl`).
//   3. The eager `FunctionRegistry` path, with range-aware argument
//      expansion for `accepts_ranges` entries.
//
// Unknown names yield `#NAME?`; arity violations yield `#VALUE!`;
// argument errors propagate left-to-right unless the function opted out
// of `propagate_errors` (the IS* type-predicate family).
Value dispatch_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx);

// Invokes a runtime `LambdaValue` with the given argument-AST accessor
// and arity. Shared between the `LambdaCall` AST case (a parser-emitted
// IIFE or curried call) and the name-bound dispatch path inside
// `dispatch_call` (where the user wrote `f(x)` and `f` resolves through
// `NameEnv` to a Lambda).
//
// Arity check: required slots = `param_count - optional_count`; the call
// must satisfy `required <= arity <= param_count`. Anything else surfaces
// `#VALUE!`. Trailing optional slots that the caller did not supply bind
// to an "omitted" sentinel that `ISOMITTED` detects via `lookup_omitted`.
// Argument evaluation is eager and left-to-right in the *caller's* scope;
// the first error short-circuits.
Value invoke_lambda(const LambdaValue* lv, std::uint32_t arity, const parser::AstNode* const* call_args, Arena& arena,
                    const FunctionRegistry& registry, const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TREE_WALKER_DISPATCH_H_
