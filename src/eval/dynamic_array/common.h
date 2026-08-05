//
// Shared helpers for the dynamic-array spilling builtins
// (`FILTER`/`UNIQUE`/`SORT`/`SORTBY`/`HSTACK`/`VSTACK`/`CHOOSECOLS`/
// `CHOOSEROWS`/`TAKE`/`DROP`/`EXPAND`/`TOCOL`/`TOROW`/`WRAPROWS`/
// `WRAPCOLS`/`ANCHORARRAY`). The per-function impls live in sibling
// `dynamic_array/<group>.cpp` translation units; this header gathers the
// argument-resolution and array-materialisation primitives they share so
// no single helper is duplicated across the split.
//
// Cell-comparison helpers (`unique_cell_equal`, `sort_cell_less_asc`,
// `sort_lane_less`, `ascii_ci_less`) intentionally implement the
// Excel-canonical cross-kind ordering for the dynamic-array surface and
// are NOT migrated to the eval::coerce / eval::jp helpers: the W3.5
// matrix-strict / fold_jp_text APIs operate on numeric coercion or
// criterion-equality folding, which is broader than what UNIQUE / SORT
// need (ASCII case-insensitive Text compare plus Number / Bool / Error /
// Blank kind ordering).

#ifndef FORMULON_EVAL_DYNAMIC_ARRAY_COMMON_H_
#define FORMULON_EVAL_DYNAMIC_ARRAY_COMMON_H_

#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

namespace dynamic_array {

/// The dynamic-array family allocates through the shared evaluator seam; the
/// name is re-exported here so the family's call sites keep reading in terms
/// of their own namespace. See `eval/array_alloc.h` for the contract.
using formulon::eval::allocate_array_value;

/// Evaluate `node`, coerce to a finite number, and truncate toward zero.
/// On error or coercion failure writes the caller-visible error to
/// `error_out` and returns `false`. Shared by every helper that takes a
/// numeric scalar argument (counts, indices, ignore masks, etc.).
bool eval_truncated_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, double& out, Value& error_out);

/// Gather rows (`by_col == false`) or columns (`by_col == true`) of
/// `src` named by the 0-based indices in `indices`, preserving the index
/// order. Returns nullptr on arena OOM.
ArrayValue* materialise_selected_lanes(const ArrayValue& src, const std::vector<std::uint32_t>& indices, bool by_col,
                                       Arena& arena);

/// Materialise a 2D row-major slice `[row_lo, row_hi) x [col_lo, col_hi)`
/// of `src` into `arena`. Returns nullptr on arena OOM. Both half-open
/// intervals must satisfy `lo <= hi <= src axis size`.
ArrayValue* materialise_slice(const ArrayValue& src, std::uint32_t row_lo, std::uint32_t row_hi, std::uint32_t col_lo,
                              std::uint32_t col_hi, Arena& arena);

/// Build a `1xN` (when `as_column == false`) or `Nx1` (when
/// `as_column == true`) `ArrayValue` from the gathered cells. Returns
/// nullptr on arena OOM. Callers that need to surface a specific error
/// for "no cells" should pre-validate and not call this with
/// `cells.empty()`.
ArrayValue* materialise_vector(std::vector<Value>&& cells, bool as_column, Arena& arena);

/// Excel-canonical cell equality for UNIQUE. Mirrors `Value::operator==`
/// for Number / Bool / Error / Blank, but uses ASCII case-insensitive
/// compare for Text (matching `=A1=B1`, COUNTIF, and the SWITCH precedent
/// in `special_forms_lazy`). Cross-kind pairs are never equal.
bool unique_cell_equal(const Value& a, const Value& b);

/// Returns `true` iff the `i`-th and `j`-th rows (`by_col == false`) or
/// columns (`by_col == true`) of `arr` are cellwise equal under
/// `unique_cell_equal`.
bool unique_lane_equal(const ArrayValue& arr, std::uint32_t i, std::uint32_t j, bool by_col);

/// Strict weak ordering for SORT keys, ascending. Cross-kind ordering
/// follows Excel: `Number < Text < Bool < Error < Blank`. Text uses
/// ASCII case-insensitive lex; Bool uses `FALSE < TRUE`; Error uses
/// error-code value compare.
bool sort_cell_less_asc(const Value& a, const Value& b);

/// Lane-level less for SORT / SORTBY. Always sinks blank keys to the end
/// regardless of `descending`; otherwise applies `sort_cell_less_asc` (or
/// its mirror for descending). The caller's `std::stable_sort` relies on
/// this being a strict weak ordering.
bool sort_lane_less(const Value& key_a, const Value& key_b, bool descending);

/// ASCII case-insensitive lex comparison. Returns `true` iff `a < b`.
/// Walks both strings byte-by-byte with `tolower` applied, matching the
/// Excel-canonical text ordering used by SWITCH / UNIQUE for equality.
/// Avoids materialising a full lowercase copy.
bool ascii_ci_less(std::string_view a, std::string_view b);

/// Resolve a SORT / SORTBY order argument: `1` -> ascending,
/// `-1` -> descending, anything else -> `#VALUE!`. Returns `true` and
/// writes `descending` on success; `false` plus caller-visible error
/// otherwise.
bool resolve_sort_order_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, bool& descending, Value& error_out);

/// Resolve a 1-based / negative index for CHOOSECOLS / CHOOSEROWS into a
/// 0-based axis index. Returns `true` and writes `out` on success;
/// `false` and writes the caller-visible error on coercion failure, zero
/// index, or out-of-range.
bool resolve_choose_index(const parser::AstNode& node, std::uint32_t axis_size, Arena& arena,
                          const FunctionRegistry& registry, const EvalContext& ctx, std::uint32_t& out,
                          Value& error_out);

/// Decode a signed count argument for TAKE / DROP into a half-open
/// `[lo, hi)` slice of the source axis. `take == true` selects cells to
/// retain; `take == false` selects cells to drop (and the retained range
/// is the complement). On success populates `lo` / `hi`; on failure
/// writes the caller-visible error and returns `false`. Omitted argument
/// is signalled by passing `nullptr` for `node`.
bool resolve_take_drop_range(const parser::AstNode* node, std::uint32_t axis_size, bool take, Arena& arena,
                             const FunctionRegistry& registry, const EvalContext& ctx, std::uint32_t& lo,
                             std::uint32_t& hi, Value& error_out);

}  // namespace dynamic_array
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DYNAMIC_ARRAY_COMMON_H_
