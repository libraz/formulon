//
// Shared infrastructure for the GROUPBY / PIVOTBY lazy impls. Both Excel 365
// dynamic-array functions consume a (row_fields, values [, col_fields]) shape
// and a user-supplied aggregator (Form A: inline `LAMBDA`, Form B: name-bound
// `LAMBDA`, Form C: bare registry function name), then produce a 2D spilled
// result whose body cells are per-group aggregates.
//
// The helpers in this header carry the parts of the algorithm that are
// genuinely identical between GROUPBY (`groupby.cpp`) and PIVOTBY
// (`pivotby.cpp`): aggregator resolution, argument parsing, group-key
// equality, slice construction, aggregator invocation, output assembly and
// sort tie-breaking. Each .cpp owns its public `eval_<fn>_lazy` entry point
// and any helper that is specific to that surface (e.g. PIVOTBY's
// `order_axis_groups`).
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// `eval_node` recursion entry point.

#ifndef FORMULON_EVAL_GROUPBY_PIVOTBY_COMMON_H_
#define FORMULON_EVAL_GROUPBY_PIVOTBY_COMMON_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "eval/function_registry.h"
#include "eval/lambda_value.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

class EvalContext;

struct HeaderLayout {
  bool inputs_have_header = false;
  bool output_emits_header = false;
  std::uint32_t data_start_row = 0;
  std::uint32_t data_row_count = 0;
};

/// Empty-text cell Excel fills result placeholders with (pivot corners, empty
/// intersections, key cells of total rows); it is text, not blank.
inline Value placeholder_cell() {
  return Value::text(std::string_view("", 0U));
}

/// Generated header labels for `field_headers == 2`, e.g. "行フィールド 1",
/// "列フィールド 1" and "値 1" in ja-JP; `n` is 1-based.
std::string row_field_label(const EvalContext& ctx, std::uint32_t n);
std::string column_field_label(const EvalContext& ctx, std::uint32_t n);
std::string value_label(const EvalContext& ctx, std::uint32_t n);

enum class HeaderField : std::uint8_t { kRow, kColumn, kValue };

/// The header cell of the `index`-th (0-based) field of `kind`: copied from
/// `input_header_cells` when the input carries its own header row
/// (`field_headers == 3`), else the generated label.
Value header_label(Arena& arena, const EvalContext& ctx, int field_headers, HeaderField kind,
                   const Value* input_header_cells, std::uint32_t index);

/// Returns the locale-appropriate "Grand Total" label.
std::string_view grand_total_label(const EvalContext& ctx);

/// Returns the label of the grand-total row for a layout that also emits
/// subtotal rows. ja-JP names the two levels apart ("総計" above the
/// subtotals' "合計"); locales that draw no such distinction reuse
/// `grand_total_label`.
std::string_view hierarchy_grand_total_label(const EvalContext& ctx);

/// The grouping of the flat composite-key groups by their first `prefix_cols`
/// key columns. Prefix groups are numbered in first-occurrence order of their
/// key.
struct OuterGrouping {
  /// Prefix-group ordinal of each flat group, indexed by flat-group index.
  std::vector<std::size_t> outer_of_group;
  /// Every data row belonging to each prefix group, for the subtotal
  /// aggregation.
  std::vector<std::vector<std::uint32_t>> rows_of_outer;
  /// A representative source row per prefix group, for reading its key cells.
  std::vector<std::uint32_t> repr_of_outer;
  /// The first flat group of each prefix group.
  std::vector<std::size_t> first_group_of_outer;
};

/// Groups the flat groups produced by the composite-key pass by their first
/// `prefix_cols` key columns. `group_repr` and `group_rows` are the
/// per-flat-group representative row and member rows; `keys` is the key array
/// they index.
OuterGrouping build_prefix_grouping(const ArrayValue& keys, const std::vector<std::uint32_t>& group_repr,
                                    const std::vector<std::vector<std::uint32_t>>& group_rows,
                                    std::uint32_t prefix_cols);

/// A node of an axis hierarchy: `level` is the key-prefix length it covers
/// (0 = the whole axis, the axis's key count = a flat group) and `index`
/// numbers it within its level.
struct AxisEntity {
  std::uint32_t level = 0;
  std::size_t index = 0;
};

/// The key-prefix hierarchy of one axis: `prefix[L - 1]` groups the flat
/// groups by their first L key columns, for L in 1..keys-1. Empty when the
/// axis needs neither subtotals nor parent totals.
struct AxisHierarchy {
  std::uint32_t keys = 0;
  std::vector<OuterGrouping> prefix;
};

/// Builds the hierarchy of an axis with `keys.cols` key columns; the prefix
/// levels only when `with_prefixes`.
AxisHierarchy build_axis_hierarchy(const ArrayValue& keys, const std::vector<std::uint32_t>& group_repr,
                                   const std::vector<std::vector<std::uint32_t>>& group_rows, bool with_prefixes);

/// The ordinal of the `level`-prefix group that flat group `flat_group` is in.
std::size_t axis_group_at(const AxisHierarchy& hierarchy, std::uint32_t level, std::size_t flat_group);

/// The parent total of `e`: the next shorter key prefix containing it. A
/// prefix-1 group's parent, and the grand total's, is the grand total.
AxisEntity axis_parent(const AxisHierarchy& hierarchy, AxisEntity e);

/// A source row whose key cells read as `e`'s key.
std::uint32_t axis_entity_repr_row(const AxisHierarchy& hierarchy, const std::vector<std::uint32_t>& group_repr,
                                   AxisEntity e);

/// The data rows `e` covers; `all_rows` stands for the whole axis.
const std::vector<std::uint32_t>& axis_entity_rows(const AxisHierarchy& hierarchy,
                                                   const std::vector<std::vector<std::uint32_t>>& group_rows,
                                                   const std::vector<std::uint32_t>& all_rows, AxisEntity e);

/// The nodes of an axis in output order, the grand total excluded. `order`
/// lists the flat groups sorted. For |depth| >= 2 each prefix of length
/// 1..|depth|-1 adds a subtotal node, nested block by block (blocks in
/// first-appearance order within `order`): after its block for depth > 0, before
/// it for depth < 0. Otherwise the nodes are the flat groups in `order`.
std::vector<AxisEntity> plan_axis_slots(const AxisHierarchy& hierarchy, const std::vector<std::size_t>& order,
                                        int depth);

/// True when a `total_depth` request of magnitude `depth` fits an axis of
/// `levels` key columns: |depth| <= 1 always, deeper only up to `levels`.
bool total_depth_fits(int depth, std::uint32_t levels);

/// Label of the grand-total row / column; the hierarchy form when the layout
/// also emits subtotals.
std::string_view total_label(const EvalContext& ctx, bool with_subtotals);

/// A resolved aggregator. A one-parameter aggregator receives each group's
/// value slice; a two-parameter one (PERCENTOF's `(subset, all)` shape) also
/// receives `all_by_column[value_col]`, every filtered row of that column.
struct Aggregator {
  const LambdaValue* fn = nullptr;
  std::uint32_t arity = 1U;
  std::vector<const ArrayValue*> all_by_column;
  /// Array-literal AST of each `all_by_column` slice, bound alongside it; an
  /// entry is null when the AST could not be allocated.
  std::vector<const parser::AstNode*> all_ast_by_column;
};

/// Resolves the aggregator argument (3rd for GROUPBY, 4th for PIVOTBY)
/// through the shared lambda-helper resolver (`resolve_callable`): an inline
/// or name-bound `LAMBDA`, or a bare built-in name. It takes two arguments
/// when the `LAMBDA` or built-in requires exactly two; otherwise one. Returns
/// false and writes the scalar error to `*out_err` on failure.
bool resolve_aggregator(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, Aggregator* out, Value* out_err);

/// Returns `agg` with `all_by_column` bound for a two-parameter aggregator:
/// one slice per value column over `all_rows`, the rows of the result column
/// the cells sit in. A one-parameter aggregator comes back unchanged. False
/// on allocation failure.
bool bind_aggregator_totals(const ArrayValue& values, const std::vector<std::uint32_t>& all_rows, Arena& arena,
                            const Aggregator& agg, Aggregator* out);

/// Reads an array argument via the standard `eval_node_as_array` seam so a
/// Ref / RangeOp / ArrayLiteral / OFFSET-call argument keeps its 2D shape.
/// Returns nullptr and writes the appropriate scalar error to `*out_err` on
/// failure paths.
const ArrayValue* read_array_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx, Value* out_err);

/// Reads a scalar integer argument truncated toward zero, then validates it
/// is a member of `allowed`. Returns true on success; on failure writes the
/// appropriate scalar error to `*out_err` and returns false.
bool read_int_in_set(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, const int* allowed, std::size_t count, int* out, Value* out_err);

/// Reads a scalar integer argument truncated toward zero (no set membership).
bool read_int(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
              int* out, Value* out_err);

/// Optional integer-in-set with a default; the optional position is at
/// `arg_index` and the caller's `arity` is used to short-circuit missing args.
bool read_optional_int_in_set(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity,
                              int default_value, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                              const int* allowed, std::size_t count, int* out, Value* out_err);

/// Optional `field_headers` slot, a member of {0,1,2,3}.
bool read_field_headers(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity, int default_value,
                        Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx, int* out,
                        Value* out_err);

/// Optional `total_depth` slot: an integer (a fractional value is #VALUE!).
/// `total_depth_fits` bounds it by the axis's key count.
bool read_total_depth(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity, int default_value,
                      Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx, int* out, Value* out_err);

/// Reads an optional sort-order slot (GROUPBY `sort_order`, PIVOTBY
/// `row_sort_order` / `col_sort_order`): one signed 1-based column index, or
/// a vector of them. Absent or syntactically omitted leaves `*out` empty (the
/// default key-ascending sort); a supplied 0, a blank or a non-finite entry
/// is #VALUE!. `plan_group_sort` checks the indices against the columns.
bool read_optional_sort_spec(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity, Arena& arena,
                             const FunctionRegistry& registry, const EvalContext& ctx, std::vector<int>* out,
                             Value* out_err);

/// How one axis orders its groups. A sort spec numbers `key_cols` key
/// columns first, then `val_cols` value columns.
struct GroupSort {
  /// Per key column; keys order the groups column by column in `key_order`.
  std::vector<bool> key_descending;
  /// Key columns in comparison order: left to right for a hierarchy, the
  /// spec's key columns first for a flat table.
  std::vector<std::uint32_t> key_order;
  /// 1-based value column to sort by, ties broken by the keys ascending; 0
  /// sorts by the keys.
  std::uint32_t value_col = 0U;
  bool value_descending = false;
};

/// Plans `spec` over `key_cols` keys and `val_cols` values: key indices set
/// the direction of their key column, a lone value index sorts by that value.
/// `flat_table` compares the named key columns first, in spec order, then the
/// rest left to right. False (#VALUE!) for an index past the last column or a
/// vector naming a value column.
bool plan_group_sort(const std::vector<int>& spec, std::uint32_t key_cols, std::uint32_t val_cols, bool flat_table,
                     GroupSort* out);

/// One outer level of a hierarchical value sort: the prefix group each group
/// belongs to and that prefix group's aggregate.
struct OuterSortLevel {
  std::vector<std::size_t> prefix_of_group;
  std::vector<Value> value_of_group;
};

/// Builds one `OuterSortLevel` per key prefix shorter than `keys.cols`, each
/// prefix group aggregated over value column `value_col` with the aggregator
/// of its first member group (`agg_of_group`). Empty for a single key column.
std::vector<OuterSortLevel> build_outer_sort_levels(const ArrayValue& keys, const std::vector<std::uint32_t>& repr,
                                                    const std::vector<std::vector<std::uint32_t>>& members,
                                                    const ArrayValue& values, std::uint32_t value_col,
                                                    const std::vector<const Aggregator*>& agg_of_group, Arena& arena,
                                                    const FunctionRegistry& registry, const EvalContext& ctx);

/// Fills `*order` with the group indices ordered by `plan`; error-keyed
/// groups sink to the bottom. `value_keys[g]` is group g's sort value, read
/// only for a value sort. A value sort orders `outer_levels` first, outermost
/// first, so each prefix group stays contiguous and ranks by its aggregate.
void order_groups(const GroupSort& plan, const ArrayValue& keys, const std::vector<std::uint32_t>& repr,
                  const std::vector<bool>& is_error, const std::vector<Value>& value_keys,
                  const std::vector<OuterSortLevel>& outer_levels, std::vector<std::size_t>* order);

/// `field_headers` value that requests header detection from the values
/// argument; `resolve_auto_field_headers` turns it into a concrete mode.
inline constexpr int kFieldHeadersAuto = -1;

/// Returns `field_headers` unchanged unless it is `kFieldHeadersAuto`, which
/// becomes 1 (input header row, not shown) when some column of `values` has
/// text in its first row and a non-text cell below it, and 0 otherwise.
int resolve_auto_field_headers(int field_headers, const ArrayValue& values);

/// Computes input/output header layout from `field_headers ∈ {0,1,2,3}` and
/// the input row count.
Expected<HeaderLayout, ErrorCode> resolve_header_layout(int field_headers, std::uint32_t input_rows);

/// Reads the optional filter_array argument: a 1D column-shaped boolean mask
/// over the data rows.
bool read_filter_mask(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, std::uint32_t data_row_count, std::vector<bool>* include_row,
                      Value* out_err);

/// Resolves the header layout for `field_headers` and, when the call supplies
/// the filter argument (present and not omitted), reads its mask.
/// `*include_row` is sized to the layout's data row count.
bool read_layout_and_mask(const parser::AstNode& call, std::uint32_t filter_arg_index, std::uint32_t arity,
                          int field_headers, std::uint32_t input_rows, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx, HeaderLayout* layout, std::vector<bool>* include_row, Value* out_err);

/// Returns the list of absolute row indices included by `include_row`,
/// offset by `data_start_row`.
std::vector<std::uint32_t> collect_included_rows(const std::vector<bool>& include_row, std::uint32_t data_start_row);

/// Builds a hashable canonical key for one row. Text cells become their
/// dynamic-strength collation key; all scalar kinds retain their type tag.
/// Values which never compare equal (NaN, Array / Ref / Lambda) receive a
/// row-unique key.
std::string normalized_group_key(const ArrayValue& keys, std::uint32_t row);

/// True iff any column of the row's key is an Error value. Error-keyed
/// groups sort to the bottom (after all valid groups).
bool row_key_is_error(const ArrayValue& keys, std::uint32_t row);

/// Maps `row`'s group key to its group index, appending a new group (its
/// representative row, member list and error flag) on first occurrence so
/// groups keep first-occurrence order.
std::size_t find_or_add_group(const ArrayValue& keys, std::uint32_t row,
                              std::vector<std::uint32_t>* representative_rows,
                              std::vector<std::vector<std::uint32_t>>* member_rows, std::vector<bool>* is_error_group,
                              std::unordered_map<std::string, std::size_t>* index);

/// Builds a 1-column array containing one cell per row in `row_indices` from
/// `values`'s `value_col`-th column. Used to construct the per-group slice
/// passed to the aggregator.
const ArrayValue* build_group_slice(const ArrayValue& values, std::uint32_t value_col,
                                    const std::vector<std::uint32_t>& row_indices, Arena& arena);

/// Invokes the resolved aggregator for one group's slice of value column
/// `value_col` and returns whatever it produced. Errors are returned verbatim
/// (the per-group error isolation seam). A multi-cell array return surfaces
/// as `#CALC!`.
Value invoke_aggregator_for_group(const Aggregator& agg, std::uint32_t value_col, const ArrayValue* slice, Arena& arena,
                                  const FunctionRegistry& registry, const EvalContext& ctx);

/// Aggregates each value column over the given row indices.
std::vector<Value> aggregate_value_columns(const ArrayValue& values, std::uint32_t val_cols,
                                           const std::vector<std::uint32_t>& row_indices, const Aggregator& agg,
                                           Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                                           ErrorCode empty_error);

/// Renders one row of cells into the buffer.
void emit_row(std::vector<std::vector<Value>>* rows, std::vector<Value> row);

/// Materialises a rectangular output buffer into a spilled `Value::array`.
Value rows_to_array_value(const std::vector<std::vector<Value>>& rows, std::uint32_t out_cols, Arena& arena);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_GROUPBY_PIVOTBY_COMMON_H_
