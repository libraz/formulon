// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

Value eval_if_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx);
Value eval_iferror_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);
Value eval_ifna_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);
Value eval_and_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx);
Value eval_or_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
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

// `ISOMITTED(arg)` — TRUE iff `arg` is a bare name reference resolving
// to an omitted trailing-optional LAMBDA parameter. Lazy because the
// query is about the AST shape and the binding's `is_omitted` flag, not
// the bound value. Outside of a LAMBDA call (or for any arg that isn't
// a NameRef into the active environment) the result is FALSE.
Value eval_isomitted_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx);

// Compile-time guard: every lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`. If a future edit drifts a parameter list the
// dispatch table in `tree_walker.cpp` would have to coerce the symbol
// through a non-matching cast — this `static_cast` makes that drift a
// header-level compile error instead. Picking `eval_if_lazy` as a witness
// is sufficient because every sibling in this header shares the same
// parameter list.
inline constexpr LazyImpl kSpecialFormsLazySignatureWitness = &eval_if_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SPECIAL_FORMS_LAZY_H_
