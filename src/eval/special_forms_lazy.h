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

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SPECIAL_FORMS_LAZY_H_
