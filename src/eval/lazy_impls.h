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

#include "eval/tail_array.h"
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

/// `eval_node` without the final densification: a whole column / row
/// (`A:A`, `1:1`, `A:C`), an operator or scalar function applied to one, a
/// LET-bound one, comes back as a `TailArray` of its declared size. Every other
/// result is a plain `Shaped{value}`. `eval_node` is `densify` of this.
Shaped eval_node_shaped(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

/// The node as a reference where it denotes one: a cell, range or whole-axis
/// token and a 3-D span become a `Ref` (no cell is read), a call (a LAMBDA call included) or a literal
/// yielding a `Ref` returns it as it is. Any other node is its `eval_node`
/// value. A `Ref` from here is read through `eval/reference/deref.h`.
Value eval_node_ref(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx);

/// Whether `node` evaluates to a `Ref` when it names a reference, so `eval_node_ref` reads it once: a cell,
/// range or whole-axis token, a 3-D span, a `:` / ` ` / `,` over such operands, a call that returns its
/// reference (a lazy entry whose `ref_result` is neither `kNone` nor `kProjected`), a call that keeps the grid of
/// references such an argument may give (a `kReshape` or `kKeepRef` position), a name (a LET / LAMBDA binding or a
/// defined name, which may stand for a value instead), or a literal holding a `Ref`.
bool evaluates_to_reference(const parser::AstNode& node);

/// Whether an argument `node` is read as the `Ref` it evaluates to rather than spelled as a static reference: a
/// call returning a reference, a `:` over one (`A1:OFFSET(...)`), a name, a LET or lambda call, `@` over a
/// reference, or a literal holding a `Ref`. A name or
/// call that yields a value instead reads as that value. Static tokens and `Ref:Ref` ranges are excluded; their
/// callers expand them directly.
bool is_returned_reference(const parser::AstNode& node, const EvalContext& ctx);

/// Whether `node` is a literal holding a `Ref`, as an element of a reference grid is when a call runs once per element.
bool is_literal_reference(const parser::AstNode& node);

/// Whether `node` is known to stand for a reference only once evaluated, so a function taking a reference reads it
/// through `eval_node_ref`: a name, a table reference, a literal holding a `Ref`, or a LET, lambda call or `@` whose
/// result may be one. Reference-returning calls are excluded; their callers branch on the call first.
bool is_evaluated_reference(const parser::AstNode& node);

/// Signature of every lazy (short-circuit / range-aware) function impl.
/// The central dispatch table in `tree_walker.cpp` stores pointers of
/// this type.
using LazyImpl = Value (*)(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LAZY_IMPLS_H_
