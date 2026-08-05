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
// `find_or_add_group`).
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// `eval_node` recursion entry point.

#ifndef FORMULON_EVAL_GROUPBY_PIVOTBY_COMMON_H_
#define FORMULON_EVAL_GROUPBY_PIVOTBY_COMMON_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
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

// Discriminated reference to the aggregator selected for the call. Form C
// (bare function name) cannot be wrapped in a synthetic LambdaValue because a
// LambdaValue requires an AST body; per-group dispatch therefore branches on
// the kind tag and either invokes the lambda body or calls the registry impl.
struct AggregatorRef {
  enum class Kind { Lambda, Function };
  Kind kind = Kind::Lambda;
  const LambdaValue* lambda = nullptr;        // valid when kind == Lambda
  const FunctionDef* function_def = nullptr;  // valid when kind == Function
};

struct HeaderLayout {
  bool inputs_have_header = false;
  bool output_emits_header = false;
  std::uint32_t data_start_row = 0;
  std::uint32_t data_row_count = 0;
};

/// Returns the locale-appropriate "Grand Total" label.
std::string_view grand_total_label(const EvalContext& ctx);

/// Resolves the aggregator argument (3rd for GROUPBY, 4th for PIVOTBY) into
/// an `AggregatorRef`. Returns true on success and writes the resolved
/// aggregator to `*out`. Returns false on failure and writes the appropriate
/// scalar error to `*out_err`. See implementation for the Form A / B / C
/// resolution order.
bool resolve_aggregator(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, AggregatorRef* out, Value* out_err);

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

/// Optional plain integer with a default.
bool read_optional_int(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity, int default_value,
                       Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx, int* out,
                       Value* out_err);

/// Computes input/output header layout from `field_headers ∈ {0,1,2,3}` and
/// the input row count.
Expected<HeaderLayout, ErrorCode> resolve_header_layout(int field_headers, std::uint32_t input_rows);

/// Reads the optional filter_array argument: a 1D column-shaped boolean mask
/// over the data rows.
bool read_filter_mask(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, std::uint32_t data_row_count, std::vector<bool>* include_row,
                      Value* out_err);

/// Returns the list of absolute row indices included by `include_row`,
/// offset by `data_start_row`.
std::vector<std::uint32_t> collect_included_rows(const std::vector<bool>& include_row, std::uint32_t data_start_row);

/// Excel-canonical cell equality for GROUPBY / PIVOTBY group keys. Mirrors
/// UNIQUE's rules with one difference: Text comparison runs through
/// `fold_jp_text` first so half-width katakana folds to full-width.
bool group_cell_equal(const Value& a, const Value& b);

/// Multi-column key equality: walks each column of the keys and compares
/// cellwise via `group_cell_equal`.
bool group_key_equal(const ArrayValue& keys, std::uint32_t row_a, std::uint32_t row_b);

/// Builds a hashable canonical key for one row. Text cells are Japanese-folded
/// once here; all scalar kinds retain their type tag. Values which cannot be
/// equal under `group_cell_equal` receive a row-unique key.
std::string normalized_group_key(const ArrayValue& keys, std::uint32_t row);

/// True iff every column of the row's key is an Error value. Error-keyed
/// groups sort to the bottom (after all valid groups).
bool row_key_is_error(const ArrayValue& keys, std::uint32_t row);

/// Builds a 1-column array containing one cell per row in `row_indices` from
/// `values`'s `value_col`-th column. Used to construct the per-group slice
/// passed to the aggregator.
const ArrayValue* build_group_slice(const ArrayValue& values, std::uint32_t value_col,
                                    const std::vector<std::uint32_t>& row_indices, Arena& arena);

/// Invokes the resolved aggregator for one group's column slice and returns
/// whatever it produced. Errors are returned verbatim (the per-group error
/// isolation seam). A multi-cell array return surfaces as `#CALC!`.
Value invoke_aggregator_for_group(const AggregatorRef& agg, const ArrayValue* slice, Arena& arena,
                                  const FunctionRegistry& registry, const EvalContext& ctx);

/// Aggregates each value column over the given row indices.
std::vector<Value> aggregate_value_columns(const ArrayValue& values, std::uint32_t val_cols,
                                           const std::vector<std::uint32_t>& row_indices, const AggregatorRef& agg,
                                           Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                                           ErrorCode empty_error);

/// Renders one row of cells into the buffer.
void emit_row(std::vector<std::vector<Value>>* rows, const std::vector<Value>& row);

/// Materialises a rectangular output buffer into a spilled `Value::array`.
Value rows_to_array_value(const std::vector<std::vector<Value>>& rows, std::uint32_t out_cols, Arena& arena);

/// Comparator helper for sort tie-breaking: ascending compare on a single
/// scalar Value. Numbers compare by value; text compares by Mac-folded
/// UTF-8 bytes; cross-kind pairs use the kind() ordinal so ordering is
/// stable but unspecified in detail. Errors and Blanks are pushed to the
/// end (Excel's "blanks last" rule).
int cmp_value_asc(const Value& a, const Value& b);

/// Compares two group keys lexicographically across every column for the
/// stable-sort tie-break. Returns -1 / 0 / 1.
int cmp_keys_asc(const ArrayValue& keys, std::uint32_t a_row, std::uint32_t b_row);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_GROUPBY_PIVOTBY_COMMON_H_
