// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for dynamic-array spilling builtins that need per-argument AST
// shape inspection: `FILTER`, `UNIQUE`, `SORT`, `SORTBY`, `HSTACK`, `VSTACK`,
// `CHOOSECOLS`, `CHOOSEROWS`, `DROP`, `TAKE`, `EXPAND`, `TOCOL`, `TOROW`,
// `WRAPCOLS`, `WRAPROWS`. These functions either produce an `ArrayValue`
// whose footprint depends on the input array's 2D shape, or accept a range
// argument they must keep as a 2D rectangle (the eager dispatcher would
// flatten range cells into a 1D vector and lose the shape).
//
// Sibling of `eval/shape_ops_lazy.h` (which hosts the SUMPRODUCT-side helpers
// and TRANSPOSE) and `eval/builtins/dynamic_array.cpp` (the eager-arg
// SEQUENCE impl). Externs registered in the central `kLazyDispatch` table in
// `tree_walker.cpp`; see `eval/lazy_impls.h` for the shared signature.

#ifndef FORMULON_EVAL_DYNAMIC_ARRAY_LAZY_H_
#define FORMULON_EVAL_DYNAMIC_ARRAY_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `FILTER(array, include, [if_empty])` — returns a subset of `array`
/// determined by a parallel boolean mask `include`.
///
/// Shape rules:
///   * `include` must be a 1D array matching one of `array`'s axes:
///     either `(array.rows, 1)` to filter rows, or `(1, array.cols)` to
///     filter columns. Anything else surfaces `#VALUE!`.
///   * Each `include` cell coerces to bool via `coerce_to_bool`. Coercion
///     errors short-circuit and propagate (the entire result is the error).
///   * Cells in `include` evaluating to TRUE keep the corresponding
///     row / column in `array`; FALSE drops it.
///
/// Empty result handling:
///   * If no rows / cols are kept and `if_empty` is provided, `if_empty`
///     is returned scalar (Excel does not spill it; matches Mac).
///   * If no rows / cols are kept and `if_empty` is omitted, `#CALC!`.
///
/// Errors in `array` cells are preserved verbatim in the output; FILTER
/// does not coerce or evaluate cell contents (it only routes them).
/// Errors in `include` cells (other than coercion failures, which were
/// short-circuited above) are not currently distinguished from FALSE —
/// Mac Excel actually propagates per-cell errors here, but the conservative
/// "any error short-circuits" path matches the dominant case and avoids a
/// surprise spill of mixed bool / error cells.
Value eval_filter_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `UNIQUE(array, [by_col], [exactly_once])` — returns the distinct rows
/// (default) or columns (`by_col=TRUE`) of `array`, in first-occurrence
/// order.
///
/// Equality semantics (matches Mac Excel):
///   * Different value kinds never equal (Number 0 != Bool FALSE,
///     Blank != Text "").
///   * Numbers compare bit-exact via `==`.
///   * Booleans compare exactly.
///   * Errors compare by code.
///   * Text is ASCII case-insensitive (Excel-canonical), via
///     `strings::case_insensitive_eq`. This matches `=A1=B1` and the
///     `SWITCH` precedent in `special_forms_lazy`.
///   * Blank cells compare equal to each other.
///   * A row / column matches an earlier one only when ALL of its cells
///     match cellwise.
///
/// Modes:
///   * `exactly_once = FALSE` (default): return each distinct row / col
///     once, preserving the order of first occurrence.
///   * `exactly_once = TRUE`: return only rows / cols that occur exactly
///     once in the input. Duplicated rows are dropped entirely.
///
/// Empty input or zero matches in `exactly_once = TRUE` mode surfaces
/// `#CALC!` (Mac Excel's documented surface for "no values returned").
///
/// Errors in `array` cells are preserved verbatim in the output; UNIQUE
/// does not coerce or evaluate cell contents (it only routes them).
/// Equality across `Error`-typed cells uses error-code comparison so
/// `#N/A == #N/A` and they collapse into a single output row.
Value eval_unique_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `SORT(array, [sort_index], [sort_order], [by_col])` — returns `array`
/// with its rows (default) or columns (`by_col = TRUE`) reordered by the
/// values in the `sort_index`-th column / row.
///
/// Arguments:
///   * `sort_index` (1-based): column to sort by when `by_col = FALSE`,
///     or row to sort by when `by_col = TRUE`. Defaults to 1. Must be in
///     `[1, cols]` (rows version: `[1, rows]`); out-of-range -> `#VALUE!`.
///   * `sort_order`: `1` ascending (default) or `-1` descending. Any other
///     value -> `#VALUE!`.
///   * `by_col`: FALSE (default) sorts rows; TRUE sorts columns. Coerced
///     via `coerce_to_bool` like the UNIQUE flag.
///
/// Cell ordering (Excel-canonical, ascending):
///   1. Numbers (ascending numeric).
///   2. Text (ASCII case-insensitive lex via `strings::case_insensitive_eq`
///      / lowercase compare; matches the SWITCH / UNIQUE precedent).
///   3. FALSE.
///   4. TRUE.
///   5. Errors (by error-code value).
///   6. Blank cells (always last regardless of `sort_order`; matches Excel's
///      "blanks last" surface for SORT).
///
/// The sort is stable (`std::stable_sort`) so rows that compare equal in
/// the chosen key column retain their original input order. Errors in
/// `array` cells are preserved verbatim in the output; SORT does not
/// short-circuit on them.
Value eval_sort_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `SORTBY(array, by_array1, [order1], [by_array2, order2], ...)` — sorts
/// `array` by one or more parallel key vectors (out-of-band keys), in
/// contrast to SORT which keys off a column / row inside `array` itself.
///
/// Argument layout (variadic, alternating `by_array` / `order` after the
/// first slot):
///   * `arity = 2`: one key, ascending.
///   * `arity = 3`: one key, with explicit order.
///   * `arity = 4`: two keys, second uses default ascending.
///   * `arity = 5`: two keys, both with explicit orders.
///   * ... and so on. Maximum supported arity is 13 (six keys); deeper
///     calls return `#VALUE!`.
///
/// Axis inference: the first `by_array` decides the axis.
///   * `(N, 1)` column-vector with `array.rows == N` -> sort rows.
///   * `(1, N)` row-vector with `array.cols == N` -> sort columns.
///   * Anything else, or a subsequent `by_array_k` whose shape doesn't
///     match the first, surfaces `#VALUE!`.
///
/// Each `order_k` must be `1` or `-1` (matches SORT). The cell ordering
/// inside each key is the same Excel-canonical ranking used by SORT
/// (Number < Text < Bool FALSE < Bool TRUE < Error < Blank, with text
/// compared ASCII case-insensitively, and blanks always sinking to the
/// end regardless of `order_k`).
///
/// Multi-key compare walks keys left-to-right and decides at the first
/// non-equal key. Rows that compare equal under all keys retain their
/// input order (stable sort).
Value eval_sortby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `HSTACK(array1, array2, ...)` — concatenates the inputs horizontally
/// (column-wise). Output rows = `max(input.rows)`; output cols =
/// `sum(input.cols)`. Cells in a shorter input that are above the output
/// row count are filled with `#N/A` (Mac Excel's documented behaviour for
/// the stack family).
///
/// Scalar arguments are treated as `(1, 1)` arrays. Range / Ref /
/// RangeOp arguments preserve their 2D shape via `eval_node_as_array`.
/// Argument-level errors propagate verbatim. Arity 1..253 (Excel's
/// observed ceiling); anything else surfaces `#VALUE!`.
Value eval_hstack_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `VSTACK(array1, array2, ...)` — concatenates the inputs vertically
/// (row-wise). Output rows = `sum(input.rows)`; output cols =
/// `max(input.cols)`. Cells in a narrower input that are beyond its
/// column count are filled with `#N/A`.
///
/// Scalar / shape / error / arity rules match `HSTACK`.
Value eval_vstack_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `CHOOSECOLS(array, col_num1, [col_num2], ...)` — returns an array
/// composed of the columns of `array` named by the index arguments, in
/// the order given.
///
/// Indices are 1-based after `coerce_to_number` + truncate-toward-zero.
/// Negative indices count from the right end (`-1` is the last column,
/// `-array.cols` is the first). Index `0`, or any index whose absolute
/// value exceeds `array.cols`, surfaces `#VALUE!`. Indices may repeat;
/// e.g. `CHOOSECOLS(A, 2, 2, 1)` is valid and yields a 3-column result.
///
/// Argument-level errors (in `array` or any index) propagate verbatim.
/// Arity 2..254 (Excel's observed ceiling); anything else -> `#VALUE!`.
Value eval_choosecols_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);

/// `CHOOSEROWS(array, row_num1, [row_num2], ...)` — symmetric variant of
/// `CHOOSECOLS` operating on rows instead of columns. All argument
/// semantics, error policy, and arity bounds match.
Value eval_chooserows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);

/// `TAKE(array, rows, [columns])` — returns a sub-array taken from one
/// corner of `array`.
///
/// Sign-based corner selection (matches Mac Excel):
///   * Positive `rows` -> take from the TOP edge (first |rows| rows).
///   * Negative `rows` -> take from the BOTTOM edge (last |rows| rows).
///   * Positive `columns` -> take from the LEFT edge.
///   * Negative `columns` -> take from the RIGHT edge.
///   * `columns` omitted -> all columns.
///   * `|rows|` >= `array.rows` -> take all rows; `|columns|` >=
///     `array.cols` -> take all columns. (Excel does NOT error on this;
///     it clamps.)
///   * `rows == 0` -> resulting axis is 0-sized -> `#CALC!`. Same for
///     `columns == 0`. (`rows` omitted is fine; defaults to "all rows".)
///
/// `rows` / `columns` are coerced via `coerce_to_number` and truncated
/// toward zero. Argument-level errors propagate verbatim.
Value eval_take_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `DROP(array, rows, [columns])` — returns the sub-array remaining after
/// dropping `rows` and `columns` from one corner of `array`.
///
/// Sign-based corner selection mirrors `TAKE`:
///   * Positive `rows` -> drop the TOP |rows| rows.
///   * Negative `rows` -> drop the BOTTOM |rows| rows.
///   * Positive `columns` -> drop from the LEFT.
///   * Negative `columns` -> drop from the RIGHT.
///   * `columns` omitted -> drop no columns.
///
/// If `|rows|` >= `array.rows` or `|columns|` >= `array.cols` the
/// resulting axis is 0-sized and DROP surfaces `#CALC!` (matches Mac
/// Excel's documented behaviour for "nothing left").
Value eval_drop_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `EXPAND(array, rows, [columns], [pad_with])` — pads `array` with the
/// scalar `pad_with` (default `#N/A`) up to `rows` rows and `columns`
/// columns. The original cells stay anchored at the top-left.
///
/// Argument rules (matches Mac Excel):
///   * `rows` must be `>= array.rows`; `columns` must be `>= array.cols`
///     (or omitted, in which case the column count is unchanged).
///     A target smaller than the existing axis surfaces `#VALUE!`.
///   * Both dimensions are coerced via `coerce_to_number` and truncated
///     toward zero.
///   * `pad_with`, if omitted, is `#N/A`. A scalar; arrays / refs are
///     evaluated normally and only their scalar payload is used.
///   * Argument-level errors propagate verbatim.
///   * Arity 2..4.
Value eval_expand_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `TOCOL(array, [ignore], [scan_by_column])` — flattens a 2D array into
/// a single column.
///
/// Arguments:
///   * `ignore` (default 0): bitmask controlling which cells to skip.
///       0 -> keep all cells.
///       1 -> skip blank cells.
///       2 -> skip error cells.
///       3 -> skip both blanks and errors.
///     Coerced via `coerce_to_number` + truncate-toward-zero; any other
///     integer value surfaces `#VALUE!`.
///   * `scan_by_column` (default FALSE): FALSE iterates row-major
///     (left-to-right, top-to-bottom); TRUE iterates column-major
///     (top-to-bottom, left-to-right). Coerced via `coerce_to_bool`.
///
/// If every cell is filtered out, surfaces `#CALC!` (matches Mac Excel's
/// documented "no values to return" surface for the dynamic-array
/// reshape family).
Value eval_tocol_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

/// `TOROW(array, [ignore], [scan_by_column])` — symmetric variant of
/// `TOCOL`, returning a single row instead of a single column. The
/// `ignore` bitmask and `scan_by_column` flag have the same semantics
/// and the same error policy.
Value eval_torow_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

/// `WRAPROWS(vector, wrap_count, [pad_with])` — flattens `vector`
/// row-major and wraps it into a 2D array whose rows are `wrap_count`
/// wide. The final row, if short, is padded with the scalar `pad_with`
/// (default `#N/A`).
///
/// Argument rules:
///   * `vector` must be 1D (one row OR one column). 2D arrays surface
///     `#VALUE!` (matches Mac Excel's per-function 1D constraint).
///     Scalar args become `(1, 1)` and pass.
///   * `wrap_count` is coerced + truncated toward zero. `< 1` surfaces
///     `#NUM!`.
///   * `pad_with` defaults to `#N/A` if omitted; otherwise it is taken
///     as scalar (its first cell if an array).
///   * Argument-level errors propagate verbatim.
///   * Arity 2..3.
Value eval_wraprows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

/// `WRAPCOLS(vector, wrap_count, [pad_with])` — symmetric variant of
/// `WRAPROWS`, wrapping into columns `wrap_count` tall instead of rows
/// `wrap_count` wide. All argument and error semantics match.
Value eval_wrapcols_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DYNAMIC_ARRAY_LAZY_H_
