// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tree-walk evaluator for the Formulon AST.
//
// This is the simple recursive interpreter used during early development,
// before the bytecode VM lands, and as the fallback evaluation path for
// LAMBDA bodies. It supports scalar arithmetic, comparison, concatenation,
// unary operators, and Excel error propagation as described in
// `backup/plans/02-calc-engine.md` §2.1.1 and §2.5.
//
// AST node kinds that need a `FunctionRegistry` or `Workbook` (calls,
// references, ranges, arrays, lambdas, structured refs, defined names) are
// evaluated to the appropriate Excel error sentinel: `#NAME?` for unresolved
// names / functions / refs, `#VALUE!` for range-producing operators and
// inline arrays. Once the surrounding infrastructure exists the evaluator
// will be extended; the public API here is intentionally minimal so that
// promotion will be a strict superset.
//
// `NodeKind::Ref` resolution is delegated to an `EvalContext`: when a
// context bound to a sheet is supplied, local A1 references resolve to the
// target cell's cached value (or to the appropriate Excel error sentinel
// per `EvalContext::resolve_ref`). The one- and two-argument overloads
// construct a default `EvalContext{}` with no bound sheet, so unqualified
// references still surface as `#NAME?` — preserving the prior observable
// behaviour of callers that have not yet wired a context in.

#ifndef FORMULON_EVAL_TREE_WALKER_H_
#define FORMULON_EVAL_TREE_WALKER_H_

#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

class FunctionRegistry;
class EvalContext;

/// Evaluates the AST `node` to a scalar `Value` using the process-wide
/// default function registry.
///
/// `arena` backs any new text values produced by the evaluator (e.g. the
/// result of `="a"&"b"`). It must outlive the returned `Value`. Literal
/// text values are returned without re-interning, so the AST arena that
/// owns the literal payload must also stay alive for the result to remain
/// readable.
///
/// Numeric overflow, division by zero, and other runtime arithmetic faults
/// are reported as Excel error sentinels (`#NUM!`, `#DIV/0!`, ...). Errors
/// in either operand of a binary operator propagate left-to-right per the
/// Excel rule documented in `backup/plans/02-calc-engine.md` §2.1.1.
///
/// Unsupported node kinds return `Value::error` with an appropriate Excel
/// error code: `#NAME?` for refs / lambdas / let bindings, and `#VALUE!`
/// for range / union / intersect operators and array literals. Function
/// calls are dispatched through the registry: unknown names yield `#NAME?`,
/// arity violations yield `#VALUE!`.
Value evaluate(const parser::AstNode& node, Arena& arena);

/// Same as the single-arg form, but uses the caller-supplied `registry`
/// for function dispatch. Useful for tests that want to inject a custom or
/// minimal function set.
Value evaluate(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry);

/// Full evaluate: caller supplies both the function registry and an
/// `EvalContext`. The context anchors local A1 references (see
/// `EvalContext::resolve_ref` for the resolution rules). When the context
/// is unbound, references continue to surface as `#NAME?` exactly as the
/// shorter overloads do.
Value evaluate(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx);

/// Returns a pointer to an array of canonical UPPERCASE names for the
/// evaluator's lazy / special-form dispatch table (IF, IFERROR, IFS,
/// CHOOSE, SUMIF, VLOOKUP, OFFSET, INDIRECT, ...). These are routed by the
/// tree walker before the `FunctionRegistry` is consulted, so they do NOT
/// appear in `default_registry()`. Pair with `FunctionRegistry::for_each_name`
/// to enumerate every function name the engine recognizes.
///
/// The returned array is terminated by a nullptr sentinel, has static
/// storage duration, and must not be freed.
const char* const* lazy_form_names();

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TREE_WALKER_H_
