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
// point for the lazy lambda helpers, so the arity and
// omitted-parameter rules are stated once.
//
// This header is internal to the tree-walker family and is not part of
// the public evaluator surface — production callers reach the evaluator
// through `eval/tree_walker.h`.

#ifndef FORMULON_EVAL_TREE_WALKER_DISPATCH_H_
#define FORMULON_EVAL_TREE_WALKER_DISPATCH_H_

#include <cstdint>
#include <string_view>

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
/// Strips the xlsx-only `_xlfn.` / `_xlfn._xlws.` storage prefixes from a
/// function name (ASCII case-insensitively). xlsx tags post-2007 functions
/// with them and Excel strips the tag on load, so the bare name is the only
/// one the registry and every name-keyed table know.
std::string_view strip_future_prefix(std::string_view name) noexcept;

Value dispatch_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx);

/// True for an AST written as a reference -- a cell, a `:` range, a union, an
/// intersection, a 3-D or spill reference -- which Excel refuses to call
/// (`A1(1)`, `(A1:A2)(1)`) with #REF!.
bool is_reference_shape(const parser::AstNode& node) noexcept;

/// Returns the static `Ref` / `RangeOp` a reference-valued `expr` denotes (`A1`,
/// `A1:B2`, a reference-returning call resolved to its rectangle, a name bound
/// to one), or nullptr for any other source.
const parser::AstNode* resolve_binding_reference(const parser::AstNode& expr, Arena& arena,
                                                 const FunctionRegistry& registry, const EvalContext& ctx);

/// Evaluates the source of a LET binding or a LAMBDA argument, the one rule
/// both follow. A reference source also yields in `*out_ast` the AST from
/// `resolve_binding_reference` for the binding to record, so the bound name
/// stays a reference; array literals and `A1#` spill references are recorded
/// as written. Any other source binds by value alone and `*out_ast` is nullptr.
/// A recorded AST never names a binding, so it reads the same inside a lambda
/// body whose scope is not the caller's.
Value eval_binding_source(const parser::AstNode& expr, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx, const parser::AstNode** out_ast);

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

/// Resolves the reference a call of `lv` with `call_args` returns
/// (`LAMBDA(x,x)(A1:A3)` names A1:A3), binding the parameters as a call does.
/// False with `*out_err` when the body yields no reference.
bool resolve_lambda_reference(const LambdaValue* lv, std::uint32_t arity, const parser::AstNode* const* call_args,
                              Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                              std::string_view* out_sheet, std::uint32_t* out_top_row, std::uint32_t* out_left_col,
                              std::uint32_t* out_bottom_row, std::uint32_t* out_right_col, ErrorCode* out_err);

/// Invokes an AST-backed runtime lambda with arguments that have already
/// been evaluated. This is the bridge used by the lazy lambda
/// helpers (`MAP` / `BYROW` / `BYCOL` / `REDUCE` / `SCAN` / `MAKEARRAY`),
/// which have cell payloads rather than argument AST nodes. Argument values
/// are copied into the lambda environment and the AST body is then evaluated
/// by the tree walker.
///
/// Arity follows the one rule published on `LambdaValue`:
/// `param_count - optional_count <= arity <= param_count`. Trailing params
/// the caller did not supply are bound to the omitted sentinel, so
/// `ISOMITTED` inside the body sees them. Anything outside that window is
/// `#VALUE!`; a null `body` is `#NAME?`.
Value invoke_lambda_values(const LambdaValue* lv, std::uint32_t arity, const Value* args, Arena& arena,
                           const FunctionRegistry& registry, const EvalContext& ctx);

/// Resolves the callable argument of a lambda helper (MAP, BYROW, BYCOL,
/// REDUCE, SCAN, MAKEARRAY, GROUPBY, PIVOTBY) to a lambda that accepts
/// `call_arity` arguments. Accepts an inline or name-bound `LAMBDA`, and a
/// bare built-in function name, which Excel reads as the eta-reduced
/// `LAMBDA(p1, ..., pn, FN(p1, ..., pn))` and is expanded to exactly that.
/// On failure returns nullptr and writes the scalar error to `*out_err`: the
/// argument's own error, `#VALUE!` for a non-lambda or an arity the lambda
/// cannot accept, `#NAME?` for an unknown name or a body-less lambda.
const LambdaValue* resolve_callable(const parser::AstNode& arg, std::uint32_t call_arity, Arena& arena,
                                    const FunctionRegistry& registry, const EvalContext& ctx, Value* out_err);

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
