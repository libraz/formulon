// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "eval/aggregate_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <vector>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// 1..13 are the SUBTOTAL-aligned modes; 14..19 are the AGGREGATE-only "k-
// arg" modes. Storing the integer code rather than an enum keeps the
// k-arg validation switch easy to read against the Excel docs.
constexpr int kCodeAverage = 1;
constexpr int kCodeCount = 2;
constexpr int kCodeCountA = 3;
constexpr int kCodeMax = 4;
constexpr int kCodeMin = 5;
constexpr int kCodeProduct = 6;
constexpr int kCodeStdevS = 7;
constexpr int kCodeStdevP = 8;
constexpr int kCodeSum = 9;
constexpr int kCodeVarS = 10;
constexpr int kCodeVarP = 11;
constexpr int kCodeMedian = 12;
constexpr int kCodeModeSngl = 13;
constexpr int kCodeLarge = 14;
constexpr int kCodeSmall = 15;
constexpr int kCodePercentileInc = 16;
constexpr int kCodeQuartileInc = 17;
constexpr int kCodePercentileExc = 18;
constexpr int kCodeQuartileExc = 19;

constexpr int kFnMin = 1;
constexpr int kFnMax = 19;
constexpr int kFnKArgFirst = 14;  // 14..19 take a trailing k arg.

// Reads a required scalar metadata argument (function_num / options / k).
// Errors propagate verbatim; non-coercible values surface `#VALUE!`. The
// error path writes through `*out_err` and returns NaN; callers must check
// `out_err->is_error()` before consuming the numeric result.
double read_scalar(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   Value* out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return 0.0;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    *out_err = Value::error(coerced.error());
    return 0.0;
  }
  const double x = coerced.value();
  if (!std::isfinite(x)) {
    *out_err = Value::error(ErrorCode::Num);
    return 0.0;
  }
  return x;
}

// Appends every scalar Value produced by `arg_node` to `out_cells`, mirroring
// PERCENTOF's `sum_arg_for_percentof` provenance walk. LET-bound NameRefs
// resolve to their bound AST when range-shaped. On any expansion failure
// (e.g. `#REF!` from a missing sheet) returns false with the propagating
// error in `*out_err`. Returns true on a clean walk.
//
// Unlike PERCENTOF this helper does NOT filter by Value kind: AGGREGATE's
// per-mode rules (numeric branches drop non-numerics; COUNTA counts them;
// the error-ignore bit decides whether errors short-circuit) are applied
// later by `apply_filters`.
bool collect_arg(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                 const EvalContext& ctx, std::vector<Value>* out_cells, Value* out_err) {
  const parser::AstNode* effective = &arg_node;
  if (arg_node.kind() == parser::NodeKind::NameRef) {
    const parser::AstNode& resolved = resolve_name_ast(arg_node, ctx.name_env());
    if (&resolved != &arg_node && is_range_shaped_ast(resolved)) {
      effective = &resolved;
    }
  }
  const parser::AstNode& node = *effective;
  const parser::NodeKind k = node.kind();

  // Range / Ref / SpillRef / RangeOp -> use the canonical resolver.
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp || k == parser::NodeKind::SpillRef) {
    std::vector<Value> cells;
    ErrorCode range_err = ErrorCode::Value;
    if (!resolve_range_arg(node, arena, registry, ctx, &cells, &range_err)) {
      *out_err = Value::error(range_err);
      return false;
    }
    out_cells->insert(out_cells->end(), cells.begin(), cells.end());
    return true;
  }

  // Inline array literal `{a;b;c}` walked in row-major order.
  if (k == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t rows = node.as_array_rows();
    const std::uint32_t cols = node.as_array_cols();
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        const Value v = eval_node(node.as_array_element(r, c), arena, registry, ctx);
        out_cells->push_back(v);
      }
    }
    return true;
  }

  // Anything else (literal scalar, arithmetic expression, function call) is
  // treated as a single direct scalar. The caller's per-mode filter decides
  // whether non-numerics contribute.
  const Value v = eval_node(node, arena, registry, ctx);
  out_cells->push_back(v);
  return true;
}

// Filters `cells` in place according to the options bit and the function
// code's expected provenance:
//
//   * Errors: dropped silently when ignore_errors == true; the first error
//     short-circuits the call and is written to `*out_err` otherwise.
//   * COUNTA (code 3): keep every non-blank, drop blanks.
//   * Numeric modes (everything else): keep only Numbers; drop Bool / Text /
//     Blank.
//
// Returns true on a clean filter, false (with `*out_err` populated) when an
// un-ignored error short-circuits.
bool apply_filters(std::vector<Value>* cells, int code, bool ignore_errors, Value* out_err) {
  std::vector<Value> kept;
  kept.reserve(cells->size());
  for (const Value& v : *cells) {
    if (v.is_error()) {
      if (ignore_errors) {
        continue;
      }
      *out_err = v;
      return false;
    }
    if (code == kCodeCountA) {
      if (!v.is_blank()) {
        kept.push_back(v);
      }
      continue;
    }
    // Numeric branches: drop everything that is not a Number.
    if (v.is_number()) {
      kept.push_back(v);
    }
  }
  *cells = std::move(kept);
  return true;
}

// Helper: extract the numeric slice once we know every kept cell is a
// Number (true for codes 1, 2, 4..19).
std::vector<double> to_numbers(const std::vector<Value>& cells) {
  std::vector<double> out;
  out.reserve(cells.size());
  for (const Value& v : cells) {
    if (v.is_number()) {
      out.push_back(v.as_number());
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Mode runners (codes 1..13). These mirror SUBTOTAL's behaviour — empty
// ranges follow the same Excel conventions: AVERAGE/VAR/STDEV/MEDIAN ->
// #DIV/0!, MIN/MAX/PRODUCT/SUM -> 0, COUNT/COUNTA -> 0, MODE.SNGL -> #N/A.

Value run_sum(const std::vector<double>& xs) {
  double total = 0.0;
  for (double x : xs) {
    total += x;
  }
  if (!std::isfinite(total)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(total);
}

Value run_product(const std::vector<double>& xs) {
  if (xs.empty()) {
    return Value::number(0.0);
  }
  double total = 1.0;
  for (double x : xs) {
    total *= x;
  }
  if (!std::isfinite(total)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(total);
}

Value run_min_max(const std::vector<double>& xs, bool want_max) {
  if (xs.empty()) {
    return Value::number(0.0);
  }
  double best = xs[0];
  for (std::size_t i = 1; i < xs.size(); ++i) {
    if (want_max ? (xs[i] > best) : (xs[i] < best)) {
      best = xs[i];
    }
  }
  if (!std::isfinite(best)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(best);
}

Value run_average(const std::vector<double>& xs) {
  if (xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  double total = 0.0;
  for (double x : xs) {
    total += x;
  }
  const double avg = total / static_cast<double>(xs.size());
  if (!std::isfinite(avg)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(avg);
}

Value run_count(const std::vector<Value>& cells) {
  // After `apply_filters`, numeric branches retain only Numbers; this gives
  // the same answer as iterating the post-filter `cells` directly.
  std::uint32_t n = 0;
  for (const Value& v : cells) {
    if (v.is_number()) {
      ++n;
    }
  }
  return Value::number(static_cast<double>(n));
}

Value run_counta(const std::vector<Value>& cells) {
  // The COUNTA branch of `apply_filters` already dropped Blanks; everything
  // remaining contributes 1.
  return Value::number(static_cast<double>(cells.size()));
}

// Two-pass variance (matches SUBTOTAL's run_variance). Sample uses n-1;
// population uses n. n < required denominator size -> #DIV/0!.
Value run_variance(const std::vector<double>& xs, bool population) {
  const std::size_t n = xs.size();
  const std::size_t need = population ? 1U : 2U;
  if (n < need) {
    return Value::error(ErrorCode::Div0);
  }
  double sum = 0.0;
  for (double x : xs) {
    sum += x;
  }
  const double mean = sum / static_cast<double>(n);
  double sq = 0.0;
  for (double x : xs) {
    const double d = x - mean;
    sq += d * d;
  }
  const double denom = population ? static_cast<double>(n) : static_cast<double>(n - 1);
  const double var = sq / denom;
  if (!std::isfinite(var)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(var);
}

Value run_stdev(const std::vector<double>& xs, bool population) {
  const Value var = run_variance(xs, population);
  if (!var.is_number()) {
    return var;
  }
  const double v = var.as_number();
  if (v < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(std::sqrt(v));
}

Value run_median(std::vector<double> xs) {
  if (xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  std::sort(xs.begin(), xs.end());
  const std::size_t n = xs.size();
  const double m = (n % 2 == 1U) ? xs[n / 2] : 0.5 * (xs[n / 2 - 1U] + xs[n / 2]);
  if (!std::isfinite(m)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(m);
}

// MODE.SNGL: smallest value tied for the highest frequency. Excel uses
// exact double equality for tie-breaking, so we mirror that. Returns #N/A
// when no value repeats. On ties the *smallest* value wins (matches Excel
// 2013+; the legacy MODE matched the *first-encountered* mode, which is
// not what MODE.SNGL specifies).
Value run_mode_sngl(std::vector<double> xs) {
  if (xs.empty()) {
    return Value::error(ErrorCode::NA);
  }
  std::sort(xs.begin(), xs.end());
  std::size_t best_run = 1;
  double best_value = 0.0;
  bool found = false;
  std::size_t i = 0;
  while (i < xs.size()) {
    std::size_t j = i + 1;
    while (j < xs.size() && xs[j] == xs[i]) {
      ++j;
    }
    const std::size_t run = j - i;
    if (run >= 2 && (!found || run > best_run || (run == best_run && xs[i] < best_value))) {
      best_run = run;
      best_value = xs[i];
      found = true;
    }
    i = j;
  }
  if (!found) {
    return Value::error(ErrorCode::NA);
  }
  return Value::number(best_value);
}

// LARGE / SMALL — k must be a positive integer in [1, n]. k is truncated.
Value run_large_small(std::vector<double> xs, double k_raw, bool want_large) {
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  const double k_trunc = std::trunc(k_raw);
  if (!std::isfinite(k_trunc) || k_trunc < 1.0 || k_trunc > static_cast<double>(xs.size())) {
    return Value::error(ErrorCode::Num);
  }
  const auto k = static_cast<std::size_t>(k_trunc);
  std::sort(xs.begin(), xs.end());
  // LARGE: k-th largest = xs[n - k]. SMALL: k-th smallest = xs[k - 1].
  const double picked = want_large ? xs[xs.size() - k] : xs[k - 1];
  if (!std::isfinite(picked)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(picked);
}

// PERCENTILE.INC — Excel: position = 1 + p*(n-1) (1-based), linear interp.
// Domain: 0 <= p <= 1; otherwise #NUM!. Empty data -> #NUM!.
Value run_percentile_inc(std::vector<double> xs, double p) {
  if (xs.empty() || !std::isfinite(p) || p < 0.0 || p > 1.0) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const std::size_t n = xs.size();
  if (n == 1) {
    return Value::number(xs[0]);
  }
  const double pos = 1.0 + p * static_cast<double>(n - 1);  // 1-based
  const double floor_pos = std::floor(pos);
  const auto lo_index = static_cast<std::size_t>(floor_pos) - 1U;  // 0-based
  const double frac = pos - floor_pos;
  if (frac == 0.0 || lo_index + 1U >= n) {
    return Value::number(xs[lo_index]);
  }
  const double interpolated = xs[lo_index] + frac * (xs[lo_index + 1] - xs[lo_index]);
  if (!std::isfinite(interpolated)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(interpolated);
}

// PERCENTILE.EXC — Excel: position = p*(n+1) (1-based), linear interp.
// Domain: 0 < p < 1 AND p*(n+1) ∈ [1, n]; otherwise #NUM!. Empty -> #NUM!.
Value run_percentile_exc(std::vector<double> xs, double p) {
  if (xs.empty() || !std::isfinite(p) || p <= 0.0 || p >= 1.0) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const std::size_t n = xs.size();
  const double pos = p * static_cast<double>(n + 1);  // 1-based
  if (pos < 1.0 || pos > static_cast<double>(n)) {
    return Value::error(ErrorCode::Num);
  }
  const double floor_pos = std::floor(pos);
  const auto lo_index = static_cast<std::size_t>(floor_pos) - 1U;  // 0-based
  const double frac = pos - floor_pos;
  if (frac == 0.0 || lo_index + 1U >= n) {
    return Value::number(xs[lo_index]);
  }
  const double interpolated = xs[lo_index] + frac * (xs[lo_index + 1] - xs[lo_index]);
  if (!std::isfinite(interpolated)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(interpolated);
}

// QUARTILE.INC delegates to PERCENTILE.INC at p ∈ {0, 0.25, 0.5, 0.75, 1.0}.
// `quart` must be an integer in [0, 4]; truncated like the rest.
Value run_quartile_inc(std::vector<double> xs, double quart_raw) {
  const double q_trunc = std::trunc(quart_raw);
  if (!std::isfinite(q_trunc) || q_trunc < 0.0 || q_trunc > 4.0) {
    return Value::error(ErrorCode::Num);
  }
  static constexpr double kProb[] = {0.0, 0.25, 0.5, 0.75, 1.0};
  const auto idx = static_cast<std::size_t>(q_trunc);
  return run_percentile_inc(std::move(xs), kProb[idx]);
}

// QUARTILE.EXC delegates to PERCENTILE.EXC at p ∈ {0.25, 0.5, 0.75}.
// `quart` must be an integer in {1, 2, 3}; 0 and 4 are rejected.
Value run_quartile_exc(std::vector<double> xs, double quart_raw) {
  const double q_trunc = std::trunc(quart_raw);
  if (!std::isfinite(q_trunc) || q_trunc < 1.0 || q_trunc > 3.0) {
    return Value::error(ErrorCode::Num);
  }
  static constexpr double kProb[] = {0.25, 0.5, 0.75};
  const auto idx = static_cast<std::size_t>(q_trunc) - 1U;
  return run_percentile_exc(std::move(xs), kProb[idx]);
}

}  // namespace

Value eval_aggregate_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // function_num + options + at least one data arg.
  if (arity < 3U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();
  const double fn_raw = read_scalar(call.as_call_arg(0), arena, registry, ctx, &err);
  if (err.is_error()) {
    return err;
  }
  const int code = static_cast<int>(std::trunc(fn_raw));
  if (code < kFnMin || code > kFnMax) {
    return Value::error(ErrorCode::Value);
  }

  const double opts_raw = read_scalar(call.as_call_arg(1), arena, registry, ctx, &err);
  if (err.is_error()) {
    return err;
  }
  const int options = static_cast<int>(std::trunc(opts_raw));
  if (options < 0 || options > 7) {
    return Value::error(ErrorCode::Value);
  }
  // Bit 1 (mask 2) of the options byte is the error-ignore bit. The other
  // bits (hidden rows, nested SUBTOTAL/AGGREGATE) are not yet observable;
  // see the file header.
  const bool ignore_errors = (options & 2) != 0;

  std::vector<Value> cells;

  if (code >= kFnKArgFirst) {
    // 14..19 — Excel requires exactly one data range plus a trailing k.
    // Anything other than `(fn, options, data, k)` -> #VALUE!.
    if (arity != 4U) {
      return Value::error(ErrorCode::Value);
    }
    if (!collect_arg(call.as_call_arg(2), arena, registry, ctx, &cells, &err)) {
      return err;
    }
    if (!apply_filters(&cells, code, ignore_errors, &err)) {
      return err;
    }
    // k is a scalar metadata arg: errors propagate regardless of the
    // options bit (matches the function_num / options contract).
    const double k_raw = read_scalar(call.as_call_arg(3), arena, registry, ctx, &err);
    if (err.is_error()) {
      return err;
    }
    std::vector<double> xs = to_numbers(cells);
    switch (code) {
      case kCodeLarge:
        return run_large_small(std::move(xs), k_raw, /*want_large=*/true);
      case kCodeSmall:
        return run_large_small(std::move(xs), k_raw, /*want_large=*/false);
      case kCodePercentileInc:
        return run_percentile_inc(std::move(xs), k_raw);
      case kCodeQuartileInc:
        return run_quartile_inc(std::move(xs), k_raw);
      case kCodePercentileExc:
        return run_percentile_exc(std::move(xs), k_raw);
      case kCodeQuartileExc:
        return run_quartile_exc(std::move(xs), k_raw);
      default:
        // Unreachable: code is constrained to [14, 19] in this branch.
        return Value::error(ErrorCode::Value);
    }
  }

  // Codes 1..13 — every remaining positional arg is data.
  for (std::uint32_t i = 2; i < arity; ++i) {
    if (!collect_arg(call.as_call_arg(i), arena, registry, ctx, &cells, &err)) {
      return err;
    }
  }
  if (!apply_filters(&cells, code, ignore_errors, &err)) {
    return err;
  }

  switch (code) {
    case kCodeAverage:
      return run_average(to_numbers(cells));
    case kCodeCount:
      return run_count(cells);
    case kCodeCountA:
      return run_counta(cells);
    case kCodeMax:
      return run_min_max(to_numbers(cells), /*want_max=*/true);
    case kCodeMin:
      return run_min_max(to_numbers(cells), /*want_max=*/false);
    case kCodeProduct:
      return run_product(to_numbers(cells));
    case kCodeStdevS:
      return run_stdev(to_numbers(cells), /*population=*/false);
    case kCodeStdevP:
      return run_stdev(to_numbers(cells), /*population=*/true);
    case kCodeSum:
      return run_sum(to_numbers(cells));
    case kCodeVarS:
      return run_variance(to_numbers(cells), /*population=*/false);
    case kCodeVarP:
      return run_variance(to_numbers(cells), /*population=*/true);
    case kCodeMedian:
      return run_median(to_numbers(cells));
    case kCodeModeSngl:
      return run_mode_sngl(to_numbers(cells));
    default:
      // Unreachable: code is constrained to [1, 13] in this branch.
      return Value::error(ErrorCode::Value);
  }
}

}  // namespace eval
}  // namespace formulon
