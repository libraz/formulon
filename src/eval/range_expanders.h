// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Range expanders: turn a range-shaped function call AST into a flat
// row-major vector of `Value`s for consumers that aggregate across the
// expansion (SUM, AVERAGE, COUNTIF, INDEX, MATCH, ...).
//
// Each expander mirrors what `EvalContext::expand_range` does for a
// literal `RangeOp`, but with awareness of one specific call shape that
// would otherwise collapse to a scalar `Value`:
//
//   * `expand_offset_call`  -> `OFFSET(reference, ...)`
//   * `expand_choose_call`  -> `CHOOSE(idx, value1, value2, ...)`
//   * `expand_if_call`      -> `IF(cond, then, [else])`
//   * `expand_row_call`     -> `ROW(arg)`
//   * `expand_column_call`  -> `COLUMN(arg)`
//
// Each returns `true` on success and writes the flattened cells +
// rectangle dimensions through out-params; on failure returns `false`
// and writes the Excel error code through `*out_err_code`. Split out of
// the original `reference_lazy.h` so range-aware aggregator TUs don't
// have to pull in the INDIRECT / OFFSET impl declarations. The bodies
// live in `src/eval/reference/offset.cpp`.

#ifndef FORMULON_EVAL_RANGE_EXPANDERS_H_
#define FORMULON_EVAL_RANGE_EXPANDERS_H_

#include <cstdint>
#include <vector>

#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Expands `Call("OFFSET", …)` into a flat row-major vector of cell
/// `Value`s, mirroring what `EvalContext::expand_range` would produce for a
/// literal `RangeOp`. Used by `resolve_range_arg` so aggregator-family
/// builtins (SUM, AVERAGE, COUNTIF, INDEX, MATCH, …) can consume an OFFSET
/// range without a spilled `Value::Array`. Returns `true` on success and
/// fills `*out_cells` / `*out_rows` / `*out_cols`; on failure returns
/// `false` and writes the Excel error code to `*out_err_code`.
///
/// `call` must be a `NodeKind::Call` whose callee name is `"OFFSET"`
/// (case-insensitive); the caller is expected to have already verified
/// that shape.
bool expand_offset_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                        std::uint32_t* out_rows, std::uint32_t* out_cols);

/// Expands `Call("CHOOSE", index, value1, value2, …)` into a flat row-major
/// vector of cell `Value`s by evaluating `index`, validating it against the
/// `[1, arity-1]` range with `floor`-truncation (mirroring
/// `eval_choose_lazy`), and then recursively flattening the chosen child:
///
///   * a nested `OFFSET(...)` call is forwarded to `expand_offset_call`,
///   * a nested `CHOOSE(...)` call is forwarded back here recursively,
///   * any other shape (`Ref`, `RangeOp`, …) flows through
///     `resolve_range_arg`'s existing branches.
///
/// Returns `true` on success and fills `*out_cells` / `*out_rows` /
/// `*out_cols`; on failure returns `false` and writes the Excel error code
/// to `*out_err_code`. Errors produced by evaluating `index` propagate
/// unchanged; an out-of-range `index` surfaces `#VALUE!`. Scalar children
/// (literals, BinaryOp, INDIRECT, …) reach `resolve_range_arg`'s
/// "anything else -> #VALUE!" fallthrough — Excel only allows range-or-
/// scalar children to CHOOSE in this aggregator context, but the scalar
/// path is unsupported here until a `Value::Array` runtime lands.
///
/// `call` must be a `NodeKind::Call` whose callee name is `"CHOOSE"`
/// (case-insensitive); the caller is expected to have already verified
/// that shape.
bool expand_choose_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                        std::uint32_t* out_rows, std::uint32_t* out_cols);

/// Expands `Call("IF", cond, then, [else])` into a flat row-major vector of
/// cell `Value`s. Mirrors `expand_choose_call`'s contract but short-circuits
/// on `cond` rather than indexing: TRUE picks `then`, FALSE picks `else`
/// (or surfaces `#VALUE!` for the two-arity `IF(FALSE, then)` shape, since
/// Excel returns boolean FALSE — not a reference — there). The chosen
/// branch is then recursively flattened: nested `OFFSET` / `CHOOSE` / `IF`
/// calls forward to their dedicated expanders, anything else flows through
/// `resolve_range_arg`'s existing branches. Errors propagate left-to-right
/// (cond first, then the chosen branch). Used by the eager dispatcher in
/// `tree_walker.cpp` so that `=LET(r, IF(TRUE, A1:A3, B1:B3), SUM(r))`
/// aggregates the 3-cell range rather than collapsing `r` to a scalar.
///
/// `call` must be a `NodeKind::Call` whose callee name is `"IF"`
/// (case-insensitive); the caller is expected to have already verified
/// that shape.
bool expand_if_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                    std::vector<Value>* out_cells, ErrorCode* out_err_code, std::uint32_t* out_rows,
                    std::uint32_t* out_cols);

/// Expands `Call("ROW", arg)` into a vertical column of 1-based row indices
/// when `arg` resolves to a multi-row reference. ROW() with no argument
/// returns the formula cell's 1-based row as a 1x1 (or `#VALUE!` when no
/// formula cell is bound). Mirrors `expand_offset_call`'s contract: returns
/// `true` on success, fills `*out_cells` / `*out_rows` / `*out_cols`; on
/// failure returns `false` and writes the Excel error code to
/// `*out_err_code`. `*out_cols` is always 1.
///
/// Used by aggregator-family callers (SUM, AVERAGE, SUMPRODUCT, …) so
/// `=SUM(ROW(A1:A5))` aggregates `{1;2;3;4;5}` to 15. The `eval_row_lazy`
/// scalar path is intentionally left collapsing to the first row, matching
/// Mac Excel's scalar-context output for a bare `=ROW(A1:A5)`.
bool expand_row_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                     std::uint32_t* out_rows, std::uint32_t* out_cols);

/// Expands `Call("COLUMN", arg)` into a horizontal row of 1-based column
/// indices when `arg` resolves to a multi-column reference. Mirrors
/// `expand_row_call` but along the column axis: `*out_rows` is always 1.
bool expand_column_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                        std::uint32_t* out_rows, std::uint32_t* out_cols);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RANGE_EXPANDERS_H_
