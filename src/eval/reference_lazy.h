// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lazy impls for Excel's reference-manipulation builtins `INDIRECT` and
// `OFFSET`.
//
//   * `INDIRECT(ref_text, [a1])` evaluates `ref_text` to a string, parses
//     it as an A1 reference, and resolves that target on the current
//     context. It is lazy only so the caller owns arity handling and can
//     branch on `a1` before deciding how to interpret the text; the
//     text-shape -> reference decoding uses `refs_internal::parse_a1_ref`
//     declared in `eval/a1_parse.h`.
//
//   * `OFFSET(reference, rows, cols, [height], [width])` takes a literal
//     Ref / RangeOp AST node as its first argument, applies the numeric
//     offsets, and produces either a single-cell `Value` (when the
//     resulting rectangle is 1x1) or a synthetic reference that the lazy
//     aggregator dispatch in `range_args.cpp` can expand as if it were a
//     `RangeOp`. Outside that aggregator context a multi-cell OFFSET
//     degrades to `#VALUE!`, matching Excel's scalar-context behaviour
//     prior to dynamic-array spill.
//
// The accompanying range-shape adaptations live in three sibling headers
// to keep the responsibility split clean:
//
//   * `eval/a1_parse.h`          -- `refs_internal::parse_a1_ref`,
//                                   `column_letters`, the `A1Parse`
//                                   carrier struct.
//   * `eval/range_expanders.h`   -- `expand_offset_call` /
//                                   `expand_choose_call` /
//                                   `expand_if_call` / `expand_row_call`
//                                   / `expand_column_call`.
//   * `eval/range_resolvers.h`   -- `resolve_reference_call` /
//                                   `resolve_range_endpoint` /
//                                   `compute_intersect_rect`.
//
// Most TUs only need one of the three sibling headers; this header is
// only useful to the lazy-dispatch table.

#ifndef FORMULON_EVAL_REFERENCE_LAZY_H_
#define FORMULON_EVAL_REFERENCE_LAZY_H_

#include "eval/a1_parse.h"
#include "eval/range_expanders.h"
#include "eval/range_resolvers.h"
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

/// `OFFSET(reference, rows, cols, [height], [width])` — offsets `reference`
/// by `(rows, cols)` and returns either the single cell at the shifted
/// position (when height = width = 1) or `#VALUE!` when the resulting
/// rectangle is multi-cell. Multi-cell OFFSET is visible to lazy range
/// consumers (SUM/AVERAGE/COUNTIF/…) through `expand_offset_call` (declared
/// in `eval/range_expanders.h`).
Value eval_offset_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REFERENCE_LAZY_H_
