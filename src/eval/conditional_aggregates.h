// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for Excel's `*IF` / `*IFS` conditional aggregators:
// `COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`,
// `MAXIFS`, and `MINIFS`.
//
// These cannot ride on the eager `accepts_ranges` path because arg 0
// (or arg 1 in the single-criterion aggregators) must reach the impl as
// AST so a bare single-cell `Ref` can be treated as a 1-cell range, and
// the parallel result / additional criteria ranges must iterate in
// lockstep rather than being flattened into one values vector. The
// central dispatch table in `tree_walker.cpp` references these externs
// by unqualified name; see `eval/lazy_impls.h` for the shared
// `LazyImpl` signature and `eval_node` entry point.

#ifndef FORMULON_EVAL_CONDITIONAL_AGGREGATES_H_
#define FORMULON_EVAL_CONDITIONAL_AGGREGATES_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

Value eval_countif_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);
Value eval_sumif_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);
Value eval_averageif_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx);
Value eval_countifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);
Value eval_sumifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);
Value eval_averageifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);
Value eval_maxifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);
Value eval_minifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_CONDITIONAL_AGGREGATES_H_
