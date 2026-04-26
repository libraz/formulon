// Copyright 2026 libraz. Licensed under the MIT License.
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

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/coerce.h"
#include "eval/criteria.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/jp_fold.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

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
  ErrorCode err_code = ErrorCode::Value;
  if (!resolve_range_arg(arg_node, arena, registry, ctx, out_cells, &err_code, out_rows, out_cols)) {
    *out_err = Value::error(err_code);
    return false;
  }
  if (*out_rows == 0U || *out_cols == 0U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  return true;
}

// Maps a `field` argument to a 0-indexed column index into `database`.
//
//  * Number: truncated toward zero -> 1-based column index. Must lie in
//    `[1, db_cols]`.
//  * Text  : case-insensitive ASCII equality against each header cell's
//    `coerce_to_text` rendering; first match wins.
//  * Bool / Blank / Error: rejected with `#VALUE!`. (Bool in particular
//    is NOT treated as 0 / 1 — Excel rejects it.)
//
// Returns `true` on success with the 0-indexed column written to
// `*out_col_index`. On failure writes the propagating error Value to
// `*out_err` and returns `false`.
bool resolve_field_column(const Value& field_value, const std::vector<Value>& db_cells, std::uint32_t db_cols,
                          std::uint32_t* out_col_index, Value* out_err) {
  if (field_value.is_error()) {
    *out_err = field_value;
    return false;
  }
  if (field_value.is_number()) {
    const double raw = field_value.as_number();
    // Excel truncates (not rounds) toward zero, matching every other
    // integer-index argument (e.g. CHOOSE, INDEX).
    const long long idx = static_cast<long long>(raw);
    if (idx < 1 || static_cast<std::uint64_t>(idx) > db_cols) {
      *out_err = Value::error(ErrorCode::Value);
      return false;
    }
    *out_col_index = static_cast<std::uint32_t>(idx - 1);
    return true;
  }
  if (field_value.is_text()) {
    // Mac Excel ja-JP folds the field-arg text and the database header
    // text through a partial kana fold before comparing: hira/kata,
    // full-width Latin -> ASCII, and full-width digits -> ASCII digits
    // are folded; half-width katakana (U+FF61..U+FF9D) and the related
    // standalone voicing marks are NOT folded. See
    // `tests/oracle/cases/dfunc_kana_folding_probes.yaml`, in particular
    // `dsum_field_arg_halfwidth_vs_fullwidth_header` (expects #VALUE!,
    // i.e. no match) versus the hira/full-width-Latin/full-width-digit
    // sibling cases (which all match).
    const std::string folded_needle =
        fold_jp_text(field_value.as_text(), /*fold_fullwidth_digits=*/true, /*fold_halfwidth_kana=*/false);
    for (std::uint32_t c = 0; c < db_cols; ++c) {
      const Value& hdr = db_cells[c];
      // Error / Blank / Bool / Number headers: use the text coercion for
      // an Excel-visible rendering. A header cell that fails to coerce
      // (Array / Ref / Lambda) cannot match; keep scanning.
      auto coerced = coerce_to_text(hdr);
      if (!coerced) {
        continue;
      }
      const std::string folded_hdr =
          fold_jp_text(coerced.value(), /*fold_fullwidth_digits=*/true, /*fold_halfwidth_kana=*/false);
      if (strings::case_insensitive_eq(std::string_view(folded_hdr), std::string_view(folded_needle))) {
        *out_col_index = c;
        return true;
      }
    }
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  // Bool, Blank, anything else: Excel rejects as #VALUE!.
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

// Scans the database header row for the 0-indexed column whose header
// text matches `header_needle` (a single criteria-header cell) under
// case-insensitive ASCII equality. Returns `db_cols` as a sentinel for
// "no match"; callers treat that as a disjunct that cannot be satisfied.
std::uint32_t find_db_column(const Value& header_needle, const std::vector<Value>& db_cells, std::uint32_t db_cols) {
  if (header_needle.is_blank()) {
    return db_cols;
  }
  auto needle_coerced = coerce_to_text(header_needle);
  if (!needle_coerced) {
    return db_cols;
  }
  // Mac Excel ja-JP folds the criteria-block header text and the database
  // header text through a partial kana fold before comparing: hira/kata,
  // full-width Latin -> ASCII, and full-width digits -> ASCII digits are
  // folded; half-width katakana (U+FF61..U+FF9D) and the related
  // standalone voicing marks are NOT folded. See
  // `tests/oracle/cases/dfunc_kana_folding_probes.yaml`, in particular
  // `dsum_criteria_header_halfwidth_vs_fullwidth_db_header` (expects 0,
  // i.e. no match) versus the hira and full-width-Latin sibling cases
  // (which both match).
  const std::string folded_needle =
      fold_jp_text(needle_coerced.value(), /*fold_fullwidth_digits=*/true, /*fold_halfwidth_kana=*/false);
  for (std::uint32_t c = 0; c < db_cols; ++c) {
    auto hdr_coerced = coerce_to_text(db_cells[c]);
    if (!hdr_coerced) {
      continue;
    }
    const std::string folded_hdr =
        fold_jp_text(hdr_coerced.value(), /*fold_fullwidth_digits=*/true, /*fold_halfwidth_kana=*/false);
    if (strings::case_insensitive_eq(std::string_view(folded_hdr), std::string_view(folded_needle))) {
      return c;
    }
  }
  return db_cols;
}

// Tests whether database row `r` (0-indexed into `db_cells`) matches the
// criterion row `cr` of `criteria` (0-indexed into `crit_cells`). A
// criterion row matches when, for every criteria column `j`, either:
//   - `criteria[cr][j]` is blank (this criteria column is a no-op for
//     this row), OR
//   - `criteria[cr][j]` is non-blank AND the database column whose
//     header matches `criteria[0][j]` satisfies
//     `matches_criterion(db_cell, parse_criterion(crit_cell))`.
//
// If the criteria column's header has no corresponding database header
// AND the row's cell is non-blank, the criterion cannot be satisfied and
// the row fails. An entirely-blank criterion row matches every record
// (empty AND is true).
bool record_matches_criterion_row(std::uint32_t r, std::uint32_t cr, const std::vector<Value>& db_cells,
                                  std::uint32_t db_cols, const std::vector<Value>& crit_cells,
                                  std::uint32_t crit_cols) {
  for (std::uint32_t j = 0; j < crit_cols; ++j) {
    const Value& cell = crit_cells[(static_cast<std::size_t>(cr) * crit_cols) + j];
    if (cell.is_blank()) {
      continue;
    }
    const Value& header = crit_cells[j];
    const std::uint32_t db_col = find_db_column(header, db_cells, db_cols);
    if (db_col == db_cols) {
      // Criterion references a column the database does not have, and
      // the cell is non-blank. This disjunct cannot be satisfied.
      return false;
    }
    const Value& db_cell = db_cells[(static_cast<std::size_t>(r) * db_cols) + db_col];
    // Per Excel docs, D-function plain-text criteria are prefix-matched
    // (case-insensitive): "Sm" finds rows that begin with "Sm" such as
    // "Smith" and "Smithfield". `parse_criterion_dfunc` preserves every
    // other criterion shape (numeric, comparator-prefixed, wildcard,
    // error, blank).
    const ParsedCriterion parsed = parse_criterion_dfunc(cell);
    if (!matches_criterion(db_cell, parsed)) {
      return false;
    }
  }
  return true;
}

// Returns `true` iff database row `r` satisfies AT LEAST ONE criterion
// row of `criteria`. Callers must already have rejected criteria tables
// with fewer than two rows (header + at least one criterion row); this
// function therefore treats a header-only block as a no-match. Mac Excel
// 365 surfaces `#VALUE!` for that shape — see the `crit_rows < 2U` guard
// in `resolve_common()`.
bool record_matches(std::uint32_t r, const std::vector<Value>& db_cells, std::uint32_t db_cols,
                    const std::vector<Value>& crit_cells, std::uint32_t crit_rows, std::uint32_t crit_cols) {
  if (crit_rows < 2U) {
    return false;
  }
  for (std::uint32_t cr = 1; cr < crit_rows; ++cr) {
    if (record_matches_criterion_row(r, cr, db_cells, db_cols, crit_cells, crit_cols)) {
      return true;
    }
  }
  return false;
}

// Full argument resolution shared by every D-function. Writes the two
// resolved tables, the resolved 0-indexed field column, and (on error)
// the propagating error Value. Returns `true` on success.
bool resolve_common(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                    std::vector<Value>* out_db, std::uint32_t* out_db_rows, std::uint32_t* out_db_cols,
                    std::vector<Value>* out_crit, std::uint32_t* out_crit_rows, std::uint32_t* out_crit_cols,
                    std::uint32_t* out_field_col, Value* out_err) {
  if (call.as_call_arity() != 3U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  if (!resolve_table_arg(call.as_call_arg(0), arena, registry, ctx, out_db, out_db_rows, out_db_cols, out_err)) {
    return false;
  }
  // A D-function requires a database with at least one data row beneath
  // the header row; a single-row range (just the labels) surfaces
  // `#VALUE!` per Mac Excel 365. This also short-circuits the subsequent
  // field / criteria resolution, which would otherwise do useful work
  // before discovering there is nothing to aggregate.
  if (*out_db_rows < 2U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  // D-functions reject a multi-cell range as the field arg per Mac Excel
  // 365: =DCOUNT(_, $E$1:$F$1, _) yields #VALUE! rather than projecting to
  // a spill anchor. Without this guard, the post-spill RangeOp evaluator
  // would silently swallow the multi-cell shape into the top-left value.
  // Verified against `tests/oracle/cases/dfunc_field_range_arg.yaml`.
  {
    const parser::AstNode& field_node = call.as_call_arg(1);
    if (field_node.kind() == parser::NodeKind::RangeOp) {
      const parser::AstNode& lhs = field_node.as_range_lhs();
      const parser::AstNode& rhs = field_node.as_range_rhs();
      if (lhs.kind() == parser::NodeKind::Ref && rhs.kind() == parser::NodeKind::Ref) {
        const auto& lhs_ref = lhs.as_ref();
        const auto& rhs_ref = rhs.as_ref();
        if (lhs_ref.row != rhs_ref.row || lhs_ref.col != rhs_ref.col) {
          *out_err = Value::error(ErrorCode::Value);
          return false;
        }
      }
    }
  }
  // The `field` argument is a scalar; evaluate it eagerly so any error
  // in the subtree propagates with its real code.
  const Value field_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (!resolve_field_column(field_val, *out_db, *out_db_cols, out_field_col, out_err)) {
    return false;
  }
  if (!resolve_table_arg(call.as_call_arg(2), arena, registry, ctx, out_crit, out_crit_rows, out_crit_cols, out_err)) {
    return false;
  }
  // Mac Excel 365 (16.108.1, ja-JP) rejects a criteria block that has
  // only its header row with `#VALUE!`: the contract requires at least
  // one criterion row beneath the headers. Mirror that here rather than
  // silently treating "no criterion rows" as "match every record" — the
  // latter would over-count records and silently drift from Excel.
  if (*out_crit_rows < 2U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  return true;
}

// Collects the field-column cells of every matching record. Preserves
// original Value shape (Blank / Bool / Text / Number / Error) — the
// caller decides which shapes survive its aggregator.
std::vector<Value> collect_matching_field_values(const std::vector<Value>& db_cells, std::uint32_t db_rows,
                                                 std::uint32_t db_cols, const std::vector<Value>& crit_cells,
                                                 std::uint32_t crit_rows, std::uint32_t crit_cols,
                                                 std::uint32_t field_col) {
  std::vector<Value> out;
  if (db_rows <= 1U) {
    return out;
  }
  out.reserve(db_rows - 1U);
  for (std::uint32_t r = 1; r < db_rows; ++r) {
    if (record_matches(r, db_cells, db_cols, crit_cells, crit_rows, crit_cols)) {
      out.push_back(db_cells[(static_cast<std::size_t>(r) * db_cols) + field_col]);
    }
  }
  return out;
}

// Distils the numeric subset of `field_values` via `coerce_to_number`.
// Cells that fail to coerce (e.g. non-numeric text) or that carry an
// error are silently dropped, matching the SUM-over-range behaviour.
std::vector<double> collect_matching_numbers(const std::vector<Value>& field_values) {
  // Excel's range-aware numeric aggregators (SUM, AVERAGE, COUNT, MIN,
  // MAX, …) skip non-Number cells outright: Bool, Text (including numeric
  // text), Blank, and Error are all ignored. The D-family mirrors that
  // rule, so we reject anything that isn't a pure Number here instead of
  // piping through `coerce_to_number`, which would fold Bool -> 1/0.
  std::vector<double> out;
  out.reserve(field_values.size());
  for (const Value& v : field_values) {
    if (v.kind() != ValueKind::Number) {
      continue;
    }
    const double d = v.as_number();
    if (std::isnan(d) || std::isinf(d)) {
      continue;
    }
    out.push_back(d);
  }
  return out;
}

// Guards a final numeric result against NaN / infinity. Mirrors the
// helper used in `hypothesis_lazy.cpp`.
Value finite_number(double r) {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

}  // namespace

// ---------------------------------------------------------------------------
// Public impls
// ---------------------------------------------------------------------------

Value eval_dsum_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  std::vector<Value> db;
  std::vector<Value> crit;
  std::uint32_t db_rows = 0;
  std::uint32_t db_cols = 0;
  std::uint32_t crit_rows = 0;
  std::uint32_t crit_cols = 0;
  std::uint32_t field_col = 0;
  Value err = Value::blank();
  if (!resolve_common(call, arena, registry, ctx, &db, &db_rows, &db_cols, &crit, &crit_rows, &crit_cols, &field_col,
                      &err)) {
    return err;
  }
  const std::vector<Value> vals =
      collect_matching_field_values(db, db_rows, db_cols, crit, crit_rows, crit_cols, field_col);
  const std::vector<double> nums = collect_matching_numbers(vals);
  double sum = 0.0;
  for (double x : nums) {
    sum += x;
  }
  return finite_number(sum);
}

Value eval_dcount_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  std::vector<Value> db;
  std::vector<Value> crit;
  std::uint32_t db_rows = 0;
  std::uint32_t db_cols = 0;
  std::uint32_t crit_rows = 0;
  std::uint32_t crit_cols = 0;
  std::uint32_t field_col = 0;
  Value err = Value::blank();
  if (!resolve_common(call, arena, registry, ctx, &db, &db_rows, &db_cols, &crit, &crit_rows, &crit_cols, &field_col,
                      &err)) {
    return err;
  }
  const std::vector<Value> vals =
      collect_matching_field_values(db, db_rows, db_cols, crit, crit_rows, crit_cols, field_col);
  const std::vector<double> nums = collect_matching_numbers(vals);
  return Value::number(static_cast<double>(nums.size()));
}

Value eval_dcounta_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  std::vector<Value> db;
  std::vector<Value> crit;
  std::uint32_t db_rows = 0;
  std::uint32_t db_cols = 0;
  std::uint32_t crit_rows = 0;
  std::uint32_t crit_cols = 0;
  std::uint32_t field_col = 0;
  Value err = Value::blank();
  if (!resolve_common(call, arena, registry, ctx, &db, &db_rows, &db_cols, &crit, &crit_rows, &crit_cols, &field_col,
                      &err)) {
    return err;
  }
  const std::vector<Value> vals =
      collect_matching_field_values(db, db_rows, db_cols, crit, crit_rows, crit_cols, field_col);
  double count = 0.0;
  for (const Value& v : vals) {
    if (!v.is_blank()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

Value eval_daverage_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  std::vector<Value> db;
  std::vector<Value> crit;
  std::uint32_t db_rows = 0;
  std::uint32_t db_cols = 0;
  std::uint32_t crit_rows = 0;
  std::uint32_t crit_cols = 0;
  std::uint32_t field_col = 0;
  Value err = Value::blank();
  if (!resolve_common(call, arena, registry, ctx, &db, &db_rows, &db_cols, &crit, &crit_rows, &crit_cols, &field_col,
                      &err)) {
    return err;
  }
  const std::vector<Value> vals =
      collect_matching_field_values(db, db_rows, db_cols, crit, crit_rows, crit_cols, field_col);
  const std::vector<double> nums = collect_matching_numbers(vals);
  if (nums.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  double sum = 0.0;
  for (double x : nums) {
    sum += x;
  }
  return finite_number(sum / static_cast<double>(nums.size()));
}

Value eval_dmax_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  std::vector<Value> db;
  std::vector<Value> crit;
  std::uint32_t db_rows = 0;
  std::uint32_t db_cols = 0;
  std::uint32_t crit_rows = 0;
  std::uint32_t crit_cols = 0;
  std::uint32_t field_col = 0;
  Value err = Value::blank();
  if (!resolve_common(call, arena, registry, ctx, &db, &db_rows, &db_cols, &crit, &crit_rows, &crit_cols, &field_col,
                      &err)) {
    return err;
  }
  const std::vector<Value> vals =
      collect_matching_field_values(db, db_rows, db_cols, crit, crit_rows, crit_cols, field_col);
  const std::vector<double> nums = collect_matching_numbers(vals);
  if (nums.empty()) {
    return Value::number(0.0);
  }
  double best = nums[0];
  for (std::size_t i = 1; i < nums.size(); ++i) {
    if (nums[i] > best) {
      best = nums[i];
    }
  }
  return finite_number(best);
}

Value eval_dmin_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  std::vector<Value> db;
  std::vector<Value> crit;
  std::uint32_t db_rows = 0;
  std::uint32_t db_cols = 0;
  std::uint32_t crit_rows = 0;
  std::uint32_t crit_cols = 0;
  std::uint32_t field_col = 0;
  Value err = Value::blank();
  if (!resolve_common(call, arena, registry, ctx, &db, &db_rows, &db_cols, &crit, &crit_rows, &crit_cols, &field_col,
                      &err)) {
    return err;
  }
  const std::vector<Value> vals =
      collect_matching_field_values(db, db_rows, db_cols, crit, crit_rows, crit_cols, field_col);
  const std::vector<double> nums = collect_matching_numbers(vals);
  if (nums.empty()) {
    return Value::number(0.0);
  }
  double best = nums[0];
  for (std::size_t i = 1; i < nums.size(); ++i) {
    if (nums[i] < best) {
      best = nums[i];
    }
  }
  return finite_number(best);
}

Value eval_dproduct_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  std::vector<Value> db;
  std::vector<Value> crit;
  std::uint32_t db_rows = 0;
  std::uint32_t db_cols = 0;
  std::uint32_t crit_rows = 0;
  std::uint32_t crit_cols = 0;
  std::uint32_t field_col = 0;
  Value err = Value::blank();
  if (!resolve_common(call, arena, registry, ctx, &db, &db_rows, &db_cols, &crit, &crit_rows, &crit_cols, &field_col,
                      &err)) {
    return err;
  }
  const std::vector<Value> vals =
      collect_matching_field_values(db, db_rows, db_cols, crit, crit_rows, crit_cols, field_col);
  const std::vector<double> nums = collect_matching_numbers(vals);
  if (nums.empty()) {
    return Value::number(0.0);
  }
  double prod = 1.0;
  for (double x : nums) {
    prod *= x;
  }
  return finite_number(prod);
}

namespace {

// Variance driver shared by DVAR / DVARP / DSTDEV / DSTDEVP. `sample`
// selects the unbiased sample-variance estimator (divisor `n - 1`);
// otherwise the biased population estimator (divisor `n`). Caller
// handles the `sqrt` for standard-deviation functions.
//
// Returns `false` with `*out_err` written when the match set is too
// small for the chosen estimator; on success writes the variance to
// `*out_var`. Sample estimator requires `n >= 2`; population requires
// `n >= 1`.
bool compute_variance(const std::vector<double>& nums, bool sample, double* out_var, Value* out_err) {
  const std::size_t n = nums.size();
  if (sample) {
    if (n < 2U) {
      *out_err = Value::error(ErrorCode::Div0);
      return false;
    }
  } else {
    if (n == 0U) {
      *out_err = Value::error(ErrorCode::Div0);
      return false;
    }
  }
  double sum = 0.0;
  for (double x : nums) {
    sum += x;
  }
  const double dn = static_cast<double>(n);
  const double mean = sum / dn;
  double ss = 0.0;
  for (double x : nums) {
    const double d = x - mean;
    ss += d * d;
  }
  const double divisor = sample ? (dn - 1.0) : dn;
  *out_var = ss / divisor;
  return true;
}

Value aggregate_variance(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, bool sample, bool stddev) {
  std::vector<Value> db;
  std::vector<Value> crit;
  std::uint32_t db_rows = 0;
  std::uint32_t db_cols = 0;
  std::uint32_t crit_rows = 0;
  std::uint32_t crit_cols = 0;
  std::uint32_t field_col = 0;
  Value err = Value::blank();
  if (!resolve_common(call, arena, registry, ctx, &db, &db_rows, &db_cols, &crit, &crit_rows, &crit_cols, &field_col,
                      &err)) {
    return err;
  }
  const std::vector<Value> vals =
      collect_matching_field_values(db, db_rows, db_cols, crit, crit_rows, crit_cols, field_col);
  const std::vector<double> nums = collect_matching_numbers(vals);
  double var = 0.0;
  if (!compute_variance(nums, sample, &var, &err)) {
    return err;
  }
  if (stddev) {
    if (var < 0.0) {
      // Numerical noise guard; mathematically variance is non-negative.
      return Value::error(ErrorCode::Num);
    }
    return finite_number(std::sqrt(var));
  }
  return finite_number(var);
}

}  // namespace

Value eval_dstdev_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  return aggregate_variance(call, arena, registry, ctx, /*sample=*/true, /*stddev=*/true);
}

Value eval_dstdevp_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  return aggregate_variance(call, arena, registry, ctx, /*sample=*/false, /*stddev=*/true);
}

Value eval_dvar_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  return aggregate_variance(call, arena, registry, ctx, /*sample=*/true, /*stddev=*/false);
}

Value eval_dvarp_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return aggregate_variance(call, arena, registry, ctx, /*sample=*/false, /*stddev=*/false);
}

Value eval_dget_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  std::vector<Value> db;
  std::vector<Value> crit;
  std::uint32_t db_rows = 0;
  std::uint32_t db_cols = 0;
  std::uint32_t crit_rows = 0;
  std::uint32_t crit_cols = 0;
  std::uint32_t field_col = 0;
  Value err = Value::blank();
  if (!resolve_common(call, arena, registry, ctx, &db, &db_rows, &db_cols, &crit, &crit_rows, &crit_cols, &field_col,
                      &err)) {
    return err;
  }
  const std::vector<Value> vals =
      collect_matching_field_values(db, db_rows, db_cols, crit, crit_rows, crit_cols, field_col);
  if (vals.empty()) {
    return Value::error(ErrorCode::Value);
  }
  if (vals.size() > 1U) {
    return Value::error(ErrorCode::Num);
  }
  return vals[0];
}

}  // namespace eval
}  // namespace formulon
