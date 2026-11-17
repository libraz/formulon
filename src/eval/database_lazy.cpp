//
// Implementation of the database-aggregation lazy family. See
// `database_lazy.h` for the dispatch-table contract and per-function
// semantics.
//
// The twelve D-functions all share the same argument shape and record-
// matching logic. A record is a data row of `database` (rows `[1, rows)`
// — row 0 is the header row) that satisfies the compound criterion
// defined by the `criteria` range: ANY criteria row matches, and within
// a criteria row every non-blank header-matched predicate must match.
// Once a record is selected, the twelve impls differ only in how they
// reduce the field-column values.

#include "eval/database_lazy.h"

#include <cassert>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/builtins/numeric_helpers.h"
#include "eval/coerce.h"
#include "eval/criteria.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/jp_fold.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "eval/scalar_args.h"
#include "eval/tree_walker/dispatch.h"
#include "excel_locale.h"
#include "numeric_aggregate_kernels.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::to_finite_value;

// ---------------------------------------------------------------------------
// Shared resolution helpers
// ---------------------------------------------------------------------------

// Resolves a range argument (database or criteria) to a flat row-major
// cell vector plus the (rows, cols) shape. Surfaces the same error codes
// as `resolve_range_arg`; additionally rejects empty rectangles with
// `#VALUE!` because neither a header-less database nor a header-less
// criteria table has defined semantics.
//
// Returns `true` on success. On failure writes the propagating error
// Value to `*out_err` and returns `false`.
bool resolve_table_arg(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx, std::vector<Value>* out_cells, std::uint32_t* out_rows,
                       std::uint32_t* out_cols, Value* out_err) {
  auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
  if (!resolved) {
    *out_err = Value::error(resolved.error());
    return false;
  }
  auto& rr = resolved.value();
  *out_rows = rr.rows;
  *out_cols = rr.cols;
  *out_cells = std::move(rr.cells);
  if (*out_rows == 0U || *out_cols == 0U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  return true;
}

// The database header row, text-coerced and folded; a header that cannot be
// coerced to text (Array / Ref / Lambda) never matches.
struct FoldedHeaders {
  std::vector<std::string> text;
  std::vector<bool> valid;
};

std::string fold_header_text(std::string_view text, bool jp_fold, bool fold_halfwidth_kana) {
  return jp_fold ? fold_jp_text(text, /*fold_fullwidth_digits=*/true, fold_halfwidth_kana) : std::string(text);
}

FoldedHeaders fold_headers(const std::vector<Value>& db_cells, std::uint32_t db_cols, bool jp_fold,
                           bool fold_halfwidth_kana) {
  FoldedHeaders headers;
  headers.text.resize(db_cols);
  headers.valid.assign(db_cols, false);
  for (std::uint32_t c = 0; c < db_cols; ++c) {
    // Error / Blank / Bool / Number headers use the text coercion for an
    // Excel-visible rendering.
    auto coerced = coerce_to_text(db_cells[c]);
    if (coerced) {
      headers.text[c] = fold_header_text(coerced.value(), jp_fold, fold_halfwidth_kana);
      headers.valid[c] = true;
    }
  }
  return headers;
}

// The first column whose folded header equals `folded_needle` under
// case-insensitive ASCII equality; the header count when none does.
std::uint32_t find_header_column(const std::string& folded_needle, const FoldedHeaders& headers) {
  for (std::uint32_t c = 0; c < headers.text.size(); ++c) {
    if (headers.valid[c] &&
        strings::case_insensitive_eq(std::string_view(headers.text[c]), std::string_view(folded_needle))) {
      return c;
    }
  }
  return static_cast<std::uint32_t>(headers.text.size());
}

// Maps a `field` argument to a 0-indexed column index into `database`.
//
//  * Number: truncated toward zero -> 1-based column index. Must lie in
//    `[1, db_cols]`.
//  * Bool  : TRUE coerces to 1 (column A); FALSE coerces to 0 and is
//    rejected with `#VALUE!` (no zero-th column).
//  * Text  : case-insensitive ASCII equality against each header cell's
//    `coerce_to_text` rendering; first match wins.
//  * Blank / Error: rejected with `#VALUE!`, whatever the error.
//
// Returns `true` on success with the 0-indexed column written to
// `*out_col_index`. On failure writes the propagating error Value to
// `*out_err` and returns `false`.
bool resolve_field_column(const Value& field_value, const std::vector<Value>& db_cells, std::uint32_t db_cols,
                          ExcelProfile profile, std::uint32_t* out_col_index, Value* out_err) {
  if (field_value.is_number() || field_value.is_boolean()) {
    // Bool coerces to 0/1 just like Excel's other integer-index arguments
    // (CHOOSE, INDEX). FALSE -> 0 falls through the `< 1` check below and
    // is rejected with `#VALUE!`.
    const double raw = field_value.is_boolean() ? (field_value.as_boolean() ? 1.0 : 0.0) : field_value.as_number();
    // Excel snaps a near integer up, then truncates toward zero.
    const double snapped = snap_near_integer(raw);
    if (!std::isfinite(snapped)) {
      *out_err = Value::error(ErrorCode::Value);
      return false;
    }
    const double idx = std::trunc(snapped);
    if (idx < 1.0 || idx > static_cast<double>(db_cols)) {
      *out_err = Value::error(ErrorCode::Value);
      return false;
    }
    *out_col_index = static_cast<std::uint32_t>(idx - 1.0);
    return true;
  }
  if (field_value.is_text()) {
    // Mac Excel folds the field-argument text and database header text
    // through the full kana/width fold. Windows keeps the existing
    // comparison, so half-width and full-width kana remain distinct there.
    const bool jp_fold = width_folding(profile) == WidthFolding::kMac;
    const FoldedHeaders headers = fold_headers(db_cells, db_cols, jp_fold, /*fold_halfwidth_kana=*/true);
    const std::uint32_t c = find_header_column(fold_header_text(field_value.as_text(), jp_fold, true), headers);
    if (c != db_cols) {
      *out_col_index = c;
      return true;
    }
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  // Blank, Array, Ref, Lambda, etc.: Excel rejects as #VALUE!.
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

// Finds the 0-indexed column whose header matches `header_needle` (a single
// criteria-header cell). Returns `db_cols` as a sentinel for "no match";
// callers treat that as a disjunct that cannot be satisfied.
std::uint32_t find_db_column(const Value& header_needle, const FoldedHeaders& headers, bool jp_fold,
                             bool fold_halfwidth_kana) {
  const std::uint32_t db_cols = static_cast<std::uint32_t>(headers.text.size());
  if (header_needle.is_blank()) {
    return db_cols;
  }
  auto needle_coerced = coerce_to_text(header_needle);
  if (!needle_coerced) {
    return db_cols;
  }
  return find_header_column(fold_header_text(needle_coerced.value(), jp_fold, fold_halfwidth_kana), headers);
}

struct CompiledCriterionPredicate {
  std::uint32_t db_col = 0;
  ParsedCriterion criterion;
};

struct CompiledCriterionRow {
  bool impossible = false;
  std::vector<CompiledCriterionPredicate> predicates;
};

// Resolves every used criteria header and parses each non-blank criterion
// once for this D-function evaluation. An unknown header makes only rows
// that use it impossible; blank cells remain omitted predicates, so an
// all-blank row is the empty AND and matches every record.
std::vector<CompiledCriterionRow> compile_criteria(const std::vector<Value>& db_cells, std::uint32_t db_cols,
                                                   const std::vector<Value>& crit_cells, std::uint32_t crit_rows,
                                                   std::uint32_t crit_cols, ExcelProfile profile) {
  std::vector<bool> used_columns(crit_cols, false);
  for (std::uint32_t r = 1; r < crit_rows; ++r) {
    for (std::uint32_t c = 0; c < crit_cols; ++c) {
      if (!crit_cells[(static_cast<std::size_t>(r) * crit_cols) + c].is_blank()) {
        used_columns[c] = true;
      }
    }
  }

  // Mac criteria-header matching follows the locale fact table: ja-JP keeps
  // half-width kana distinct, while en-US folds it with the other width and
  // kana variants. Windows keeps the existing comparison for both locales.
  const bool jp_fold = width_folding(profile) == WidthFolding::kMac;
  const bool fold_halfwidth_kana = jp_fold && !locale_facts(profile).criteria_header_keeps_halfwidth_kana;
  const FoldedHeaders headers = fold_headers(db_cells, db_cols, jp_fold, fold_halfwidth_kana);
  std::vector<std::uint32_t> db_column_for_criterion(crit_cols, db_cols);
  for (std::uint32_t c = 0; c < crit_cols; ++c) {
    if (used_columns[c]) {
      db_column_for_criterion[c] = find_db_column(crit_cells[c], headers, jp_fold, fold_halfwidth_kana);
    }
  }

  // `resolve_common` guarantees a header plus at least one criterion row.
  assert(crit_rows >= 2U);
  std::vector<CompiledCriterionRow> rows;
  rows.reserve(crit_rows - 1U);
  for (std::uint32_t r = 1; r < crit_rows; ++r) {
    CompiledCriterionRow row;
    for (std::uint32_t c = 0; c < crit_cols; ++c) {
      const Value& cell = crit_cells[(static_cast<std::size_t>(r) * crit_cols) + c];
      if (cell.is_blank()) {
        continue;
      }
      const std::uint32_t db_col = db_column_for_criterion[c];
      if (db_col == db_cols) {
        // A non-blank criterion references a missing database column, so
        // this OR row cannot match any record.
        row.impossible = true;
        break;
      }
      row.predicates.push_back({db_col, parse_criterion_dfunc(cell)});
    }
    rows.push_back(std::move(row));
  }
  return rows;
}

// Returns `true` iff database row `r` satisfies at least one precompiled
// OR row. Within a row every stored predicate must match; an empty
// predicate list is an all-blank criterion row and matches every record.
bool record_matches(std::uint32_t r, const std::vector<Value>& db_cells, std::uint32_t db_cols,
                    const std::vector<CompiledCriterionRow>& criteria, ExcelProfile profile) {
  for (const CompiledCriterionRow& row : criteria) {
    if (row.impossible) {
      continue;
    }
    bool matches_row = true;
    for (const CompiledCriterionPredicate& predicate : row.predicates) {
      const Value& db_cell = db_cells[(static_cast<std::size_t>(r) * db_cols) + predicate.db_col];
      if (!matches_criterion(db_cell, predicate.criterion, profile)) {
        matches_row = false;
        break;
      }
    }
    if (matches_row) {
      return true;
    }
  }
  return false;
}

// The resolved `database` and `criteria` tables of a D-function call.
struct DTables {
  std::vector<Value> db;
  std::uint32_t db_rows = 0;
  std::uint32_t db_cols = 0;
  std::vector<Value> crit;
  std::uint32_t crit_rows = 0;
  std::uint32_t crit_cols = 0;
};

// Resolves `database` (at least one data row beneath the header row) into
// `*out`. Returns `false` with the error in `*out_err`.
bool resolve_database(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, DTables* out, Value* out_err) {
  if (call.as_call_arity() != 3U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  if (!resolve_table_arg(call.as_call_arg(0), arena, registry, ctx, &out->db, &out->db_rows, &out->db_cols, out_err)) {
    return false;
  }
  if (out->db_rows < 2U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  return true;
}

// Resolves `criteria` (at least one criterion row beneath its headers) into
// `*out`. Returns `false` with the error in `*out_err`.
bool resolve_criteria(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, DTables* out, Value* out_err) {
  if (!resolve_table_arg(call.as_call_arg(2), arena, registry, ctx, &out->crit, &out->crit_rows, &out->crit_cols,
                         out_err)) {
    return false;
  }
  if (out->crit_rows < 2U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  return true;
}

// The `field` argument's value. A multi-cell reference is `#VALUE!`; an array
// value lifts the call.
Value field_arg_value(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const parser::AstNode& field_node = call.as_call_arg(1);
  if (field_node.kind() == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs = field_node.as_range_operand(0);
    const parser::AstNode& rhs = field_node.as_range_operand(1);
    if (field_node.as_range_arity() == 2U && lhs.kind() == parser::NodeKind::Ref &&
        rhs.kind() == parser::NodeKind::Ref) {
      const auto& lhs_ref = lhs.as_ref();
      const auto& rhs_ref = rhs.as_ref();
      if (lhs_ref.row != rhs_ref.row || lhs_ref.col != rhs_ref.col) {
        return Value::error(ErrorCode::Value);
      }
    }
  }
  const Value v = eval_node(field_node, arena, registry, ctx);
  if (v.is_array() && is_reference_shape(field_node)) {
    return Value::error(ErrorCode::Value);
  }
  return v;
}

// The field-column cells of every matching record, in record order.
std::vector<Value> collect_matching_field_values(const DTables& t, std::uint32_t field_col, ExcelProfile profile) {
  std::vector<Value> out;
  const std::vector<CompiledCriterionRow> compiled_criteria =
      compile_criteria(t.db, t.db_cols, t.crit, t.crit_rows, t.crit_cols, profile);
  out.reserve(t.db_rows - 1U);
  for (std::uint32_t r = 1; r < t.db_rows; ++r) {
    if (record_matches(r, t.db, t.db_cols, compiled_criteria, profile)) {
      out.push_back(t.db[(static_cast<std::size_t>(r) * t.db_cols) + field_col]);
    }
  }
  return out;
}

// The reduction a D-function applies to the field values of its matching
// records.
using FieldReduce = Value (*)(const std::vector<Value>& vals);

// Runs the argument shape every D-function shares — `database`, `field`,
// `criteria` — and reduces the field values of the matching records. An
// array `field` lifts the call element by element.
Value eval_dfunction(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, FieldReduce reduce) {
  DTables t;
  Value err = Value::blank();
  if (!resolve_database(call, arena, registry, ctx, &t, &err)) {
    return err;
  }
  const Value field = field_arg_value(call, arena, registry, ctx);
  const ExcelProfile profile = ctx.excel_profile();
  std::uint32_t field_col = 0;
  if (!field.is_array() && !resolve_field_column(field, t.db, t.db_cols, profile, &field_col, &err)) {
    return err;
  }
  if (!resolve_criteria(call, arena, registry, ctx, &t, &err)) {
    return err;
  }
  if (!field.is_array()) {
    return reduce(collect_matching_field_values(t, field_col, profile));
  }
  return lift_over_array_args(&field, 1U, arena, [&](const Value* args, std::uint32_t) {
    std::uint32_t col = 0;
    Value col_err = Value::blank();
    if (!resolve_field_column(args[0], t.db, t.db_cols, profile, &col, &col_err)) {
      return col_err;
    }
    return reduce(collect_matching_field_values(t, col, profile));
  });
}

// The numbers among the matched field values; text, Booleans and blanks are
// skipped as a range aggregate skips them. Returns `false` with the first
// error value in `*out_err`.
bool matching_numbers(const std::vector<Value>& vals, std::vector<double>* out, Value* out_err) {
  out->reserve(vals.size());
  for (const Value& v : vals) {
    if (v.is_error()) {
      *out_err = v;
      return false;
    }
    if (v.kind() != ValueKind::Number) {
      continue;
    }
    const double d = v.as_number();
    if (std::isfinite(d)) {
      out->push_back(d);
    }
  }
  return true;
}

// Sums the matched numbers left to right, which is the order Excel
// accumulates a range in.
double sum_of(const std::vector<double>& nums) {
  double sum = 0.0;
  for (double x : nums) {
    sum += x;
  }
  return sum;
}

Value reduce_sum(const std::vector<Value>& vals) {
  std::vector<double> nums;
  Value err = Value::blank();
  if (!matching_numbers(vals, &nums, &err)) {
    return err;
  }
  return to_finite_value(sum_of(nums));
}

// DCOUNT counts the numbers and skips everything else, errors included.
Value reduce_count(const std::vector<Value>& vals) {
  double count = 0.0;
  for (const Value& v : vals) {
    if (v.kind() == ValueKind::Number && std::isfinite(v.as_number())) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

// DCOUNTA counts every non-blank value, errors included.
Value reduce_counta(const std::vector<Value>& vals) {
  double count = 0.0;
  for (const Value& v : vals) {
    if (!v.is_blank()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

Value reduce_average(const std::vector<Value>& vals) {
  std::vector<double> nums;
  Value err = Value::blank();
  if (!matching_numbers(vals, &nums, &err)) {
    return err;
  }
  if (nums.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  return to_finite_value(sum_of(nums) / static_cast<double>(nums.size()));
}

// DMAX and DMIN differ only in which way the comparison runs; an empty
// match set is 0 for both, as it is for the range-aware MAX / MIN.
template <bool kWantMax>
Value reduce_extremum(const std::vector<Value>& vals) {
  std::vector<double> nums;
  Value err = Value::blank();
  if (!matching_numbers(vals, &nums, &err)) {
    return err;
  }
  if (nums.empty()) {
    return Value::number(0.0);
  }
  double best = nums[0];
  for (std::size_t i = 1; i < nums.size(); ++i) {
    if (kWantMax ? (nums[i] > best) : (nums[i] < best)) {
      best = nums[i];
    }
  }
  return to_finite_value(best);
}

Value reduce_product(const std::vector<Value>& vals) {
  std::vector<double> nums;
  Value err = Value::blank();
  if (!matching_numbers(vals, &nums, &err)) {
    return err;
  }
  if (nums.empty()) {
    return Value::number(0.0);
  }
  double prod = 1.0;
  for (double x : nums) {
    prod *= x;
  }
  return to_finite_value(prod);
}

// Variance driver shared by DVAR / DVARP / DSTDEV / DSTDEVP. `kSample`
// selects the unbiased estimator (divisor `n - 1`, needs `n >= 2`),
// otherwise the population one (divisor `n`, needs `n >= 1`).
template <bool kSample, bool kStddev>
Value reduce_variance(const std::vector<Value>& vals) {
  std::vector<double> nums;
  Value err = Value::blank();
  if (!matching_numbers(vals, &nums, &err)) {
    return err;
  }
  if (nums.size() < (kSample ? 2U : 1U)) {
    return Value::error(ErrorCode::Div0);
  }
  const double var = numeric_aggregate_kernels::dispersion_variance(nums, kSample);
  if (!kStddev) {
    return to_finite_value(var);
  }
  if (var < 0.0) {
    // Numerical noise guard; mathematically variance is non-negative.
    return Value::error(ErrorCode::Num);
  }
  return to_finite_value(std::sqrt(var));
}

Value reduce_get(const std::vector<Value>& vals) {
  if (vals.empty()) {
    return Value::error(ErrorCode::Value);
  }
  if (vals.size() > 1U) {
    return Value::error(ErrorCode::Num);
  }
  return vals[0];
}

}  // namespace

// ---------------------------------------------------------------------------
// Public impls
// ---------------------------------------------------------------------------

Value eval_dsum_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_sum);
}

Value eval_dcount_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_count);
}

Value eval_dcounta_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_counta);
}

Value eval_daverage_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_average);
}

Value eval_dmax_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_extremum<true>);
}

Value eval_dmin_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_extremum<false>);
}

Value eval_dproduct_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_product);
}

Value eval_dstdev_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_variance<true, true>);
}

Value eval_dstdevp_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_variance<false, true>);
}

Value eval_dvar_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_variance<true, false>);
}

Value eval_dvarp_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_variance<false, false>);
}

Value eval_dget_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  return eval_dfunction(call, arena, registry, ctx, reduce_get);
}

}  // namespace eval
}  // namespace formulon
