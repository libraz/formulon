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
// Its already-evaluated-arguments siblings `invoke_lambda_values` /
// `invoke_lambda_values_with_ast` are the single lambda-invocation entry
// point for the bytecode VM and the lazy lambda helpers, so the arity and
// omitted-parameter rules are stated once.
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

/// Invokes an AST-backed runtime lambda with arguments that have already
/// been evaluated. This is the safe bridge used by the bytecode VM when a
/// direct `Call` names a workbook-defined LAMBDA, and by the lazy lambda
/// helpers (`MAP` / `BYROW` / `BYCOL` / `REDUCE` / `SCAN` / `MAKEARRAY`),
/// which have cell payloads rather than argument AST nodes. Argument values
/// are copied into the lambda environment and the AST body is then evaluated
/// by the tree walker. The helper deliberately accepts only the normal
/// AST-backed LambdaValue representation; VM-internal closure records never
/// pass through this interface.
///
/// Arity follows the one rule published on `LambdaValue`:
/// `param_count - optional_count <= arity <= param_count`. Trailing params
/// the caller did not supply are bound to the omitted sentinel, so
/// `ISOMITTED` inside the body sees them. Anything outside that window is
/// `#VALUE!`; a null `body` is `#NAME?`.
Value invoke_lambda_values(const LambdaValue* lv, std::uint32_t arity, const Value* args, Arena& arena,
                           const FunctionRegistry& registry, const EvalContext& ctx);

/// `invoke_lambda_values` with an optional parallel array of AST nodes
/// (length `arity`) recorded alongside each binding. When non-null, the AST
/// node lets range-aware consumers inside the lambda body see the binding as
/// a range-shaped expression — the seam `BYROW` / `BYCOL` need so `SUM(r)`
/// flattens a row slice instead of receiving an opaque `Value::Array`. Pass
/// `nullptr` for scalar-only bindings.
Value invoke_lambda_values_with_ast(const LambdaValue* lv, std::uint32_t arity, const Value* args,
                                    const parser::AstNode* const* ast_args, Arena& arena,
                                    const FunctionRegistry& registry, const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TREE_WALKER_DISPATCH_H_
