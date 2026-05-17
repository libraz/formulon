// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lazy impl entry point for `OFFSET(reference, rows, cols, [height],
// [width])`. The matching range-expander surface
// (`expand_offset_call` / `expand_choose_call` / `expand_if_call` /
// `expand_row_call` / `expand_column_call`) is declared in
// `eval/range_expanders.h` and lives in `reference/offset.cpp` because
// those helpers all flow through the same OFFSET rectangle-construction
// pipeline.

#ifndef FORMULON_EVAL_REFERENCE_OFFSET_H_
#define FORMULON_EVAL_REFERENCE_OFFSET_H_

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

/// `OFFSET(reference, rows, cols, [height], [width])` — offsets `reference`
/// by `(rows, cols)` and returns either the single cell at the shifted
/// position (when height = width = 1) or `#VALUE!` when the resulting
/// rectangle is multi-cell. Multi-cell OFFSET is visible to lazy range
/// consumers (SUM/AVERAGE/COUNTIF/…) through `expand_offset_call` (declared
/// in `eval/range_expanders.h`).
Value eval_offset_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

// Compile-time guard: the lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`.
inline constexpr LazyImpl kOffsetLazySignatureWitness = &eval_offset_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REFERENCE_OFFSET_H_
