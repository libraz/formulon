//
// Range resolvers: turn a reference-shaped AST node (Ref / RangeOp /
// OFFSET-call / INDIRECT-call / IntersectOp) into a rectangular
// (sheet, top, left, bottom, right) tuple WITHOUT dereferencing the
// cells. Used by:
//
//   * `tree_walker.cpp`           - intersect operator, range endpoints
//   * `range_args.cpp`            - dispatch into expand_* helpers
//   * `cell_lazy.cpp`             - CELL("address",...) anchor recovery
//   * `areas_lazy.cpp`            - AREAS() shape introspection
//   * `dynamic_array/anchor.cpp`  - spill anchor resolution
//   * `shape_ops_lazy.cpp`        - ROWS / COLUMNS shape queries
//
// Split out of the original `reference_lazy.h` so consumers that only
// need rectangle resolution don't have to pull in the INDIRECT / OFFSET
// impl declarations or the `expand_*_call` family. The bodies live in
// `src/eval/reference/intersection.cpp`.

#ifndef FORMULON_EVAL_RANGE_RESOLVERS_H_
#define FORMULON_EVAL_RANGE_RESOLVERS_H_

#include <cstdint>
#include <string_view>

#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"  // formulon::ErrorCode

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

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

/// Computes the rectangular intersection of an `IntersectOp`'s two
/// operands. Each operand is resolved via `resolve_range_endpoint`
/// (`Ref` / reference-returning `Call`) or, for `RangeOp`, by unioning
/// its two endpoints into a single rectangle. On success the inclusive
/// 0-based intersection rectangle is written to the out parameters and
/// the resolved sheet qualifier (empty = bound sheet) to `*out_sheet`.
///
/// `*out_disjoint` is set to true when the two rectangles do not
/// overlap; in that case the rectangle fields are left untouched and
/// callers translate it into Excel's `#NULL!`. The disjoint case is NOT
/// a hard error so AREAS can distinguish "no intersection" from a
/// resolution failure.
///
/// Returns false on hard error (cross-sheet mismatch -> `#REF!`,
/// whole-column / whole-row endpoint -> `#VALUE!`, non-reference shape
/// -> `#REF!`); writes the Excel error code to `*out_err`.
bool compute_intersect_rect(const parser::AstNode& lhs, const parser::AstNode& rhs, Arena& arena,
                            const FunctionRegistry& registry, const EvalContext& ctx, std::string_view* out_sheet,
                            std::uint32_t* out_top_row, std::uint32_t* out_left_col, std::uint32_t* out_bottom_row,
                            std::uint32_t* out_right_col, bool* out_disjoint, ErrorCode* out_err);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RANGE_RESOLVERS_H_
