// Copyright 2026 libraz. Licensed under the MIT License.
//
// Range-argument resolution helper shared by the conditional aggregators
// (`COUNTIF` / `SUMIF` / `AVERAGEIF` / `COUNTIFS` / ...) and the lookup
// family (`VLOOKUP` / `HLOOKUP` / `INDEX` / `MATCH` / `XLOOKUP`). These
// builtins all need to turn a lazily-parsed AST argument — which may be
// either a `RangeOp` (`A1:B2`) or a bare `Ref` — into a flat row-major
// vector of cell `Value`s plus an optional (rows, cols) shape.
//
// Hoisting this helper out of `tree_walker.cpp` lets the conditional and
// lookup impls be split into their own translation units without each TU
// pulling in the rest of the evaluator's internals.

#ifndef FORMULON_EVAL_RANGE_ARGS_H_
#define FORMULON_EVAL_RANGE_ARGS_H_

#include <cstdint>
#include <vector>

#include "utils/error.h"
#include "value.h"

namespace formulon {

class Arena;

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Resolves `arg_node` to a flat vector of cell `Value`s in row-major order.
///
/// Accepts either a `RangeOp` (`A1:B2`) whose endpoints are both `Ref`s, or
/// a bare `Ref` (treated as a 1-cell range). Any other shape — a literal, a
/// function call, a structured reference — fails with `#VALUE!`. An
/// expansion error from `EvalContext::expand_range` (e.g. `#REF!` for a
/// missing sheet) is surfaced via `*out_err_code` and the return value is
/// `false`.
///
/// `out_rows` and `out_cols`, when non-null, receive the rectangle's shape
/// (1-cell `Ref` -> 1x1, `RangeOp` -> computed from normalised endpoints).
/// Callers that only need the flat vector may pass `nullptr` for both.
bool resolve_range_arg(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                       std::uint32_t* out_rows = nullptr, std::uint32_t* out_cols = nullptr);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RANGE_ARGS_H_
