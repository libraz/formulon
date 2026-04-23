// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the Excel 365 XLOOKUP / XMATCH family. These builtins
// modernise the classic lookup model: caller-selectable match modes
// (Exact / Smaller / Larger / Wildcard) and search modes (FirstToLast /
// LastToFirst / BinaryAsc / BinaryDesc), with binary-search-aware
// scanning and their own wildcard / comparison semantics. They share no
// helpers with the classic lookups, so they live in a sibling translation
// unit (`lookups_xlookup.cpp`).
//
// They are lazy (rather than eager like ordinary builtins) because the
// lookup and return array arguments must reach the impl as AST so single-
// cell refs can still be treated as 1-cell ranges, and because the
// `if_not_found` argument of XLOOKUP must be evaluated only on a miss.
// The central dispatch table in `tree_walker.cpp` references these
// externs by unqualified name; see `eval/lazy_impls.h` for the shared
// `LazyImpl` signature and `eval_node` entry point.

#ifndef FORMULON_EVAL_LOOKUPS_XLOOKUP_H_
#define FORMULON_EVAL_LOOKUPS_XLOOKUP_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

Value eval_xlookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);
Value eval_xmatch_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LOOKUPS_XLOOKUP_H_
