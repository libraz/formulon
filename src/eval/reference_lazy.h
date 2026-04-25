// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls and shared helpers for Excel's reference-manipulation
// builtins `INDIRECT` and `OFFSET`. Both need AST-level inspection of
// their arguments:
//
//   * `INDIRECT(ref_text, [a1])` evaluates `ref_text` to a string, parses
//     it as an A1 reference, and resolves that target on the current
//     context. It is lazy only so the caller owns arity handling and can
//     branch on `a1` before deciding how to interpret the text; the
//     text-shape -> reference decoding uses the shared `parse_a1_ref`
//     helper below.
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
// `expand_offset_call` is the seam consumed by `resolve_range_arg` so
// `SUM(OFFSET(A1,0,0,3,3))` iterates the offset rectangle without going
// through the scalar `Value` surface. See `range_args.cpp` for the
// dispatcher hook.

#ifndef FORMULON_EVAL_REFERENCE_LAZY_H_
#define FORMULON_EVAL_REFERENCE_LAZY_H_

#include <cstdint>
#include <string_view>
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
/// consumers (SUM/AVERAGE/COUNTIF/…) through `expand_offset_call` below.
Value eval_offset_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

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

/// Attempts to resolve `node` as a reference-returning call, producing a
/// rectangular reference without dereferencing. Recognises INDIRECT (A1-style)
/// and nested OFFSET. On success writes the rectangle (0-based, inclusive)
/// into `*out_top_row`/`*out_left_col`/`*out_bottom_row`/`*out_right_col` with
/// the sheet qualifier (empty = bound sheet) in `*out_sheet`. Returns true
/// on success. On failure returns false and sets `*out_err` to the Excel
/// error code to surface. `*out_is_range` is set to true when the resolved
/// reference covers more than one cell.
///
/// The return value does NOT carry a Value; callers decide how to map the
/// rectangle (ROW returns top row, ROWS returns height, OFFSET-base treats
/// it as its input rectangle, etc.).
bool resolve_reference_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                            std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                            bool* out_is_range, ErrorCode* out_err);

/// Resolves a `:` operator endpoint into a rectangle. `node` may be a
/// plain `Ref` (1x1 rectangle) or a `Call` to `OFFSET` / `INDIRECT`,
/// in which case `resolve_reference_call` produces the rectangle.
/// Returns `true` on success and writes the rectangle (0-based,
/// inclusive) into the out parameters; the sheet qualifier (empty =
/// bound sheet) is written to `*out_sheet`. On failure returns `false`
/// and writes the Excel error code to `*out_err`.
///
/// Used by the range-aware dispatcher (`tree_walker.cpp`) and by
/// `resolve_range_arg` so `SUM(A1:OFFSET(...))` and
/// `COUNTIF(A1:INDIRECT(...), ...)` produce the union rectangle that
/// Mac Excel 365 produces. Whole-column / whole-row endpoints surface
/// `#VALUE!` because the resulting union would be unbounded; that
/// matches `expand_range`'s existing degradation for those shapes.
bool resolve_range_endpoint(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                            std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                            ErrorCode* out_err);

namespace refs_internal {

/// Output of `parse_a1_ref`: sheet qualifier (empty if unqualified),
/// 0-based row/col, and a `valid` flag. When `valid` is false the other
/// fields are meaningless. `is_range` is true when the source text
/// contained a `:` separator; the second endpoint populates
/// `row2` / `col2`.
///
/// Full-column (`D:D`, `$FF:FG`) and full-row (`5:5`, `$12:$23`) shapes
/// set `is_full_col` / `is_full_row` respectively; in those cases
/// `is_range` is also true and `row`/`row2`/`col`/`col2` are populated
/// with the resulting rectangle (full-column: rows span
/// `0..Sheet::kMaxRows-1`; full-row: cols span `0..Sheet::kMaxCols-1`).
struct A1Parse {
  bool valid = false;
  bool is_range = false;
  bool is_full_col = false;
  bool is_full_row = false;
  std::string_view sheet;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t row2 = 0;
  std::uint32_t col2 = 0;
};

/// Parses `text` as an A1-style reference (with optional sheet qualifier
/// and optional `:` range). Dollar signs in `$A$1` are accepted and
/// ignored for evaluation purposes (INDIRECT does not preserve
/// absolute/relative distinction because the return path does not need
/// them). Supports single-quoted sheet names (`'Sheet 1'!A1`) with
/// doubled-quote escaping (`'O''Brien'!A1`). Also recognises full-column
/// (`D:D`, `$FF:FG`) and full-row (`5:5`, `$12:$23`) shapes, setting
/// `is_full_col` / `is_full_row` and expanding to the implied rectangle.
/// Returns an `A1Parse` with `valid = false` for any malformed input.
A1Parse parse_a1_ref(std::string_view text);

/// Writes the uppercase A1 column letters for the 1-based column `col`
/// to `out`. Returns the number of letters written (1..3); `out` must
/// have room for at least 3 chars. `col` must satisfy
/// `1 <= col <= 16384`; callers are expected to range-check first.
std::size_t column_letters(std::uint32_t col, char* out);

}  // namespace refs_internal

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REFERENCE_LAZY_H_
