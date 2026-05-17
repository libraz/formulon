// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lazy impl entry point for `INDIRECT(ref_text, [a1])`. Declared in its
// own header so the central dispatch table (`tree_walker_lazy_table.cpp`)
// can include it directly without pulling in the OFFSET / intersection
// surfaces. The accompanying text-to-rectangle decoder lives in
// `eval/a1_parse.h` (public) and `reference/common.h` (intra-subdir).

#ifndef FORMULON_EVAL_REFERENCE_INDIRECT_H_
#define FORMULON_EVAL_REFERENCE_INDIRECT_H_

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

/// `INDIRECT(ref_text, [a1])` — evaluates `ref_text`, parses it as a
/// single-cell A1 reference, and resolves that target through the bound
/// context. Range text (`"A1:B2"`) returns `#REF!` in this MVP because
/// `Value::Array` is not yet implemented; R1C1 style (`a1=FALSE`) is also
/// deferred and surfaces as `#REF!`. Empty / malformed text -> `#REF!`.
Value eval_indirect_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

// Compile-time guard: the lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`.
inline constexpr LazyImpl kIndirectLazySignatureWitness = &eval_indirect_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REFERENCE_INDIRECT_H_
