// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the short-circuit "special form" family: `IF`,
// `IFERROR`, and `IFNA`. Each impl owns its own arity check and decides
// which argument subtrees actually need to be evaluated, preserving
// Excel's short-circuit semantics.
//
// The central dispatch table in `tree_walker.cpp` references these
// externs by unqualified name. See `eval/lazy_impls.h` for the shared
// `LazyImpl` signature and the `eval_node` entry point these impls
// recurse through.

#ifndef FORMULON_EVAL_SPECIAL_FORMS_LAZY_H_
#define FORMULON_EVAL_SPECIAL_FORMS_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

Value eval_if_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx);
Value eval_iferror_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);
Value eval_ifna_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

// IFS(cond1, val1, cond2, val2, ...) - multi-branch short-circuit. Each
// condition is evaluated in turn; the first TRUE wins and the paired value
// is returned. Untaken value subtrees are never evaluated. Returns #N/A
// when no condition matches (including when the argument count is zero or
// odd). Errors in any evaluated condition propagate.
Value eval_ifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx);

// SWITCH(expr, case1, val1, ..., [default]) - evaluates `expr` once and
// returns the value paired with the first case that equals it. An
// unpaired trailing argument (odd number of remaining args after `expr`)
// is used as the default. Returns #N/A when no match and no default.
// Comparison semantics match the `=` operator: ASCII case-insensitive for
// text, ordinary equality for numbers and bools; cross-type pairs never
// match (but are not errors).
Value eval_switch_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

// COUNT is lazy because Excel's "direct-arg bool counts, range-sourced bool
// doesn't" rule requires per-arg AST inspection: once a range has been
// flattened into a `Value` vector, the provenance of each Bool is lost.
Value eval_count_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SPECIAL_FORMS_LAZY_H_
