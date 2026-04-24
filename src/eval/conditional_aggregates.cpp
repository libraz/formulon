// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the `*IF` / `*IFS` conditional aggregator lazy
// impls. See `conditional_aggregates.h` for the dispatch-table contract
// and `eval/lazy_impls.h` for the shared vocabulary.

#include "eval/conditional_aggregates.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/criteria.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// COUNTIFS / SUMIFS / AVERAGEIFS / MAXIFS / MINIFS support
// ---------------------------------------------------------------------------
//
// The five multi-criteria aggregators share the shape "N parallel (range,
// criterion) pairs, each range the same size". Size mismatches are
// `#VALUE!` in Excel 365 — stricter than SUMIF's clamp-to-min behaviour,
// which we also match here.

/// Resolves `pair_count` consecutive (range, criterion) pairs starting at
/// argument index `first_pair_index` in `call`. Each criteria range is
/// resolved through `resolve_range_arg` and must have exactly
/// `expected_size` cells; otherwise the helper fails with `#VALUE!`.
/// Each criterion sub-expression is evaluated once; an error Value
/// propagates.
///
/// On success, appends the resolved cell vectors (in pair order) to
/// `*out_cell_arrays` and appends the parsed criteria to `*out_parsed`,
/// then returns `true`.
///
/// On failure, writes the error Value to propagate into `*out_err_value`
/// and returns `false`. `*out_cell_arrays` / `*out_parsed` may have
/// partially-accumulated state on failure; callers must not read them in
/// that case.
///
/// `*out_parsed` uses `unique_ptr` indirection so the heap-resident
/// `ParsedCriterion::rhs_storage` string is never relocated after parse
/// and its `rhs_text` `string_view` stays stable across vector growth.
bool resolve_criteria_pairs(const parser::AstNode& call, std::uint32_t first_pair_index, std::uint32_t pair_count,
                            std::size_t expected_size, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::vector<std::vector<Value>>* out_cell_arrays,
                            std::vector<std::unique_ptr<ParsedCriterion>>* out_parsed, Value* out_err_value) {
  out_cell_arrays->reserve(out_cell_arrays->size() + pair_count);
  out_parsed->reserve(out_parsed->size() + pair_count);
  for (std::uint32_t k = 0; k < pair_count; ++k) {
    const std::uint32_t range_idx = first_pair_index + (k * 2);
    const std::uint32_t crit_idx = range_idx + 1;
    std::vector<Value> cells;
    ErrorCode range_err = ErrorCode::Value;
    if (!resolve_range_arg(call.as_call_arg(range_idx), arena, registry, ctx, &cells, &range_err)) {
      *out_err_value = Value::error(range_err);
      return false;
    }
    if (cells.size() != expected_size) {
      *out_err_value = Value::error(ErrorCode::Value);
      return false;
    }
    // An error-valued criterion is NOT propagated: Excel's *IFS functions
    // accept an error criterion as a filter over error cells with the
    // matching code (see `parse_criterion` ValueKind::Error branch).
    const Value crit_val = eval_node(call.as_call_arg(crit_idx), arena, registry, ctx);
    auto parsed = std::make_unique<ParsedCriterion>(parse_criterion(crit_val));
    out_cell_arrays->push_back(std::move(cells));
    out_parsed->push_back(std::move(parsed));
  }
  return true;
}

/// Tests whether position `i` in the parallel criteria-range arrays
/// satisfies every parsed criterion. Short-circuits on the first failure.
bool all_criteria_match(const std::vector<std::vector<Value>>& criteria_cells,
                        const std::vector<std::unique_ptr<ParsedCriterion>>& parsed, std::size_t i) {
  const std::size_t n = parsed.size();
  for (std::size_t k = 0; k < n; ++k) {
    if (!matches_criterion(criteria_cells[k][i], *parsed[k])) {
      return false;
    }
  }
  return true;
}

}  // namespace

// ---------------------------------------------------------------------------
// COUNTIF / SUMIF / AVERAGEIF support
// ---------------------------------------------------------------------------
//
// The three conditional aggregators share a lot of shape: arg 0 is a
// criteria range, arg 1 is a scalar criterion, and `SUMIF` / `AVERAGEIF`
// additionally take an optional parallel range. The helpers below split
// that shape into reusable building blocks so the three impls themselves
// read as straight-line code.

// COUNTIF(range, criterion)
//
// Counts cells in `range` that match `criterion`. Errors encountered in
// the range itself are silently skipped (per-cell #DIV/0! does NOT
// propagate); an error CRITERION, however, propagates. An error in
// resolving the range itself (e.g. `#REF!` from a missing sheet) is
// surfaced directly.
Value eval_countif_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  if (call.as_call_arity() != 2) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> cells;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(0), arena, registry, ctx, &cells, &range_err)) {
    return Value::error(range_err);
  }
  // An error-valued criterion (e.g. `COUNTIF(range, #N/A)`) is NOT
  // propagated: `parse_criterion` turns it into an error-match filter.
  const Value criterion_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
  const ParsedCriterion parsed = parse_criterion(criterion_val);
  double count = 0.0;
  for (const Value& cell : cells) {
    if (matches_criterion(cell, parsed)) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

// SUMIF(range, criterion [, sum_range])
//
// When `sum_range` is omitted, the sum is taken over `range` itself. When
// provided, `sum_range` must also be a range/Ref; the two are iterated in
// parallel.
//
// Accepted divergence: if `sum_range` has a different cell count than
// `range`, we iterate `min(range.size(), sum_range.size())` rather than
// reshaping from the anchor as Excel does. Aligns with the
// range-vs-direct divergence already documented in `eval_context.cpp`.
//
// Matching cells whose sum-side value is not numeric (text, bool, blank)
// are excluded from the sum. Matches Excel — SUMIF sums only numbers even
// when the criterion itself passed on a non-numeric cell.
Value eval_sumif_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 2 && arity != 3) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> criteria_cells;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(0), arena, registry, ctx, &criteria_cells, &range_err)) {
    return Value::error(range_err);
  }
  // Error criterion is NOT propagated; `parse_criterion` converts it to a
  // filter over error cells with the same code.
  const Value criterion_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
  const ParsedCriterion parsed = parse_criterion(criterion_val);

  // Choose the effective sum range: either the explicit third arg, or a
  // copy of the criteria range when sum_range is omitted.
  const std::vector<Value>* sum_cells_ptr = nullptr;
  std::vector<Value> explicit_sum_cells;
  if (arity == 3) {
    ErrorCode sum_err = ErrorCode::Value;
    if (!resolve_range_arg(call.as_call_arg(2), arena, registry, ctx, &explicit_sum_cells, &sum_err)) {
      return Value::error(sum_err);
    }
    sum_cells_ptr = &explicit_sum_cells;
  } else {
    sum_cells_ptr = &criteria_cells;
  }
  const std::vector<Value>& sum_cells = *sum_cells_ptr;

  const std::size_t n = criteria_cells.size() < sum_cells.size() ? criteria_cells.size() : sum_cells.size();
  double sum = 0.0;
  for (std::size_t i = 0; i < n; ++i) {
    if (!matches_criterion(criteria_cells[i], parsed)) {
      continue;
    }
    const Value& sv = sum_cells[i];
    // Excel propagates errors that live at a matching position in the
    // sum range (e.g. a `#N/A` cell whose criterion-side row matches).
    // Non-matching errors are ignored, and non-numeric matches (text /
    // bool / blank) are silently skipped, same as Excel's SUMIF.
    if (sv.is_error()) {
      return sv;
    }
    if (!sv.is_number()) {
      continue;
    }
    sum += sv.as_number();
  }
  return Value::number(sum);
}

// AVERAGEIF(range, criterion [, average_range])
//
// Returns the arithmetic mean of the matching numeric cells on the
// averaging side, or `#DIV/0!` when no matches qualify. Non-numeric
// matches (text, bool, blank) are excluded from both the numerator and
// the denominator, matching Excel.
Value eval_averageif_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 2 && arity != 3) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> criteria_cells;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(0), arena, registry, ctx, &criteria_cells, &range_err)) {
    return Value::error(range_err);
  }
  // Error criterion is NOT propagated; `parse_criterion` converts it to a
  // filter over error cells with the same code.
  const Value criterion_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
  const ParsedCriterion parsed = parse_criterion(criterion_val);

  const std::vector<Value>* avg_cells_ptr = nullptr;
  std::vector<Value> explicit_avg_cells;
  if (arity == 3) {
    ErrorCode avg_err = ErrorCode::Value;
    if (!resolve_range_arg(call.as_call_arg(2), arena, registry, ctx, &explicit_avg_cells, &avg_err)) {
      return Value::error(avg_err);
    }
    avg_cells_ptr = &explicit_avg_cells;
  } else {
    avg_cells_ptr = &criteria_cells;
  }
  const std::vector<Value>& avg_cells = *avg_cells_ptr;

  const std::size_t n = criteria_cells.size() < avg_cells.size() ? criteria_cells.size() : avg_cells.size();
  double sum = 0.0;
  double count = 0.0;
  for (std::size_t i = 0; i < n; ++i) {
    if (!matches_criterion(criteria_cells[i], parsed)) {
      continue;
    }
    const Value& av = avg_cells[i];
    if (av.is_error()) {
      // Errors at matching positions propagate, matching Excel AVERAGEIF.
      return av;
    }
    if (!av.is_number()) {
      continue;
    }
    sum += av.as_number();
    count += 1.0;
  }
  if (count == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  return Value::number(sum / count);
}

// COUNTIFS(range1, crit1 [, range2, crit2, ...])
//
// Counts cells where every `(range_k[i], crit_k)` pair matches in lockstep.
// Arity must be even and >= 2. All criteria ranges must have the same cell
// count; Excel 365 raises `#VALUE!` on shape mismatch (stricter than
// SUMIF's clamp-to-min behaviour) and we match that.
Value eval_countifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2 || (arity % 2) != 0) {
    return Value::error(ErrorCode::Value);
  }
  // Resolve the first criteria range to fix the expected size.
  std::vector<std::vector<Value>> criteria_cells;
  std::vector<std::unique_ptr<ParsedCriterion>> parsed;
  std::vector<Value> first_cells;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(0), arena, registry, ctx, &first_cells, &range_err)) {
    return Value::error(range_err);
  }
  const std::size_t expected_size = first_cells.size();
  // Error criterion is NOT propagated; it filters error cells (see
  // `parse_criterion` ValueKind::Error).
  const Value first_crit = eval_node(call.as_call_arg(1), arena, registry, ctx);
  criteria_cells.push_back(std::move(first_cells));
  parsed.push_back(std::make_unique<ParsedCriterion>(parse_criterion(first_crit)));

  const std::uint32_t remaining_pairs = (arity - 2) / 2;
  if (remaining_pairs > 0) {
    Value err = Value::number(0.0);
    if (!resolve_criteria_pairs(call, /*first_pair_index=*/2, remaining_pairs, expected_size, arena, registry, ctx,
                                &criteria_cells, &parsed, &err)) {
      return err;
    }
  }

  double count = 0.0;
  for (std::size_t i = 0; i < expected_size; ++i) {
    if (all_criteria_match(criteria_cells, parsed, i)) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

// SUMIFS(sum_range, range1, crit1 [, range2, crit2, ...])
//
// Sums cells of `sum_range` where every parallel `(range_k[i], crit_k)`
// pair matches. Arity must be odd and >= 3. `sum_range` and every criteria
// range must share the same cell count; mismatch -> `#VALUE!`. Non-numeric
// cells in `sum_range` at matching positions are skipped. Empty match
// pool -> `0`.
Value eval_sumifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3 || (arity % 2) != 1) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> sum_cells;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(0), arena, registry, ctx, &sum_cells, &range_err)) {
    return Value::error(range_err);
  }
  const std::size_t expected_size = sum_cells.size();

  std::vector<std::vector<Value>> criteria_cells;
  std::vector<std::unique_ptr<ParsedCriterion>> parsed;
  const std::uint32_t pair_count = (arity - 1) / 2;
  Value err = Value::number(0.0);
  if (!resolve_criteria_pairs(call, /*first_pair_index=*/1, pair_count, expected_size, arena, registry, ctx,
                              &criteria_cells, &parsed, &err)) {
    return err;
  }

  double sum = 0.0;
  for (std::size_t i = 0; i < expected_size; ++i) {
    if (!all_criteria_match(criteria_cells, parsed, i)) {
      continue;
    }
    const Value& sv = sum_cells[i];
    if (sv.is_error()) {
      return sv;
    }
    if (!sv.is_number()) {
      continue;
    }
    sum += sv.as_number();
  }
  return Value::number(sum);
}

// AVERAGEIFS(avg_range, range1, crit1 [, range2, crit2, ...])
//
// Returns the mean of the numeric cells in `avg_range` at positions where
// every parallel criterion matches. Empty match pool -> `#DIV/0!`. Non-
// numeric cells in `avg_range` at matching positions are excluded from
// both the numerator and the denominator, matching Excel.
Value eval_averageifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3 || (arity % 2) != 1) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> avg_cells;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(0), arena, registry, ctx, &avg_cells, &range_err)) {
    return Value::error(range_err);
  }
  const std::size_t expected_size = avg_cells.size();

  std::vector<std::vector<Value>> criteria_cells;
  std::vector<std::unique_ptr<ParsedCriterion>> parsed;
  const std::uint32_t pair_count = (arity - 1) / 2;
  Value err = Value::number(0.0);
  if (!resolve_criteria_pairs(call, /*first_pair_index=*/1, pair_count, expected_size, arena, registry, ctx,
                              &criteria_cells, &parsed, &err)) {
    return err;
  }

  double sum = 0.0;
  double count = 0.0;
  for (std::size_t i = 0; i < expected_size; ++i) {
    if (!all_criteria_match(criteria_cells, parsed, i)) {
      continue;
    }
    const Value& av = avg_cells[i];
    if (av.is_error()) {
      return av;
    }
    if (!av.is_number()) {
      continue;
    }
    sum += av.as_number();
    count += 1.0;
  }
  if (count == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  return Value::number(sum / count);
}

// MAXIFS(max_range, range1, crit1 [, range2, crit2, ...])
//
// Returns the maximum numeric value in `max_range` at positions where
// every parallel criterion matches. Non-numeric cells at matching
// positions are skipped. Empty numeric pool -> `0` (Excel's quirk — no
// `#NUM!` even though "max of empty" is mathematically undefined).
Value eval_maxifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3 || (arity % 2) != 1) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> max_cells;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(0), arena, registry, ctx, &max_cells, &range_err)) {
    return Value::error(range_err);
  }
  const std::size_t expected_size = max_cells.size();

  std::vector<std::vector<Value>> criteria_cells;
  std::vector<std::unique_ptr<ParsedCriterion>> parsed;
  const std::uint32_t pair_count = (arity - 1) / 2;
  Value err = Value::number(0.0);
  if (!resolve_criteria_pairs(call, /*first_pair_index=*/1, pair_count, expected_size, arena, registry, ctx,
                              &criteria_cells, &parsed, &err)) {
    return err;
  }

  bool any = false;
  double best = 0.0;
  for (std::size_t i = 0; i < expected_size; ++i) {
    if (!all_criteria_match(criteria_cells, parsed, i)) {
      continue;
    }
    const Value& mv = max_cells[i];
    if (mv.is_error()) {
      return mv;
    }
    if (!mv.is_number()) {
      continue;
    }
    const double x = mv.as_number();
    if (!any || x > best) {
      best = x;
      any = true;
    }
  }
  if (!any) {
    return Value::number(0.0);
  }
  return Value::number(best);
}

// MINIFS(min_range, range1, crit1 [, range2, crit2, ...])
//
// Symmetric to MAXIFS: returns the minimum numeric value. Empty numeric
// pool -> `0`.
Value eval_minifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3 || (arity % 2) != 1) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> min_cells;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(0), arena, registry, ctx, &min_cells, &range_err)) {
    return Value::error(range_err);
  }
  const std::size_t expected_size = min_cells.size();

  std::vector<std::vector<Value>> criteria_cells;
  std::vector<std::unique_ptr<ParsedCriterion>> parsed;
  const std::uint32_t pair_count = (arity - 1) / 2;
  Value err = Value::number(0.0);
  if (!resolve_criteria_pairs(call, /*first_pair_index=*/1, pair_count, expected_size, arena, registry, ctx,
                              &criteria_cells, &parsed, &err)) {
    return err;
  }

  bool any = false;
  double best = 0.0;
  for (std::size_t i = 0; i < expected_size; ++i) {
    if (!all_criteria_match(criteria_cells, parsed, i)) {
      continue;
    }
    const Value& mv = min_cells[i];
    if (mv.is_error()) {
      return mv;
    }
    if (!mv.is_number()) {
      continue;
    }
    const double x = mv.as_number();
    if (!any || x < best) {
      best = x;
      any = true;
    }
  }
  if (!any) {
    return Value::number(0.0);
  }
  return Value::number(best);
}

}  // namespace eval
}  // namespace formulon
