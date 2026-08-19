//
// Implementation of the pairwise linear-regression lazy impls:
// CORREL, COVARIANCE.P, COVARIANCE.S, SLOPE, INTERCEPT, RSQ,
// FORECAST.LINEAR (aliased as FORECAST), STEYX, and the paired sum-of-
// products family SUMX2PY2 / SUMX2MY2 / SUMXMY2.
//
// Every function shares the same front-end work: walk two parallel AST
// arguments — each of which may be a `Ref`, a `RangeOp`, or an inline
// `ArrayLiteral` — produce a matching pair of flat `(cells, rows,
// cols)` tuples, reject a shape mismatch with `#N/A`, propagate any
// error cell in scan order, and otherwise distil the surviving
// numeric pairs into two `std::vector<double>`. The mathematical
// back-end is a handful of one-liners on the mean, sum-of-squared
// deviations, and sum-of-cross-products.
//
// See `eval/shape_ops_lazy.cpp` for the sibling SUMPRODUCT family this
// file is modelled on. The shape-resolution helper `resolve_array_arg_na`
// is shared with `eval/hypothesis_lazy.cpp` via `eval/range_args.{h,cpp}`
// — both families need the same `#N/A`-vocabulary remapping that
// SUMPRODUCT does not.

#include "eval/regression_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <variant>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/numeric_pairs.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Paired collection for the regression family. Excel accepts a row-vs-column
// pairing here (e.g. A1:A3 against `{1,2,3}`) as long as the total cell counts
// match, so the transpose is permitted; the hypothesis family is stricter.
//
// The arguments stay in source order — `pairs.first` therefore carries the
// leading argument, which for SLOPE / INTERCEPT / RSQ / STEYX is the known-y
// series and for the LINEST-style drivers is known-x.
std::variant<Value, NumericPairs> collect_regression_pairs(const parser::AstNode& lead_arg,
                                                           const parser::AstNode& trail_arg, Arena& arena,
                                                           const FunctionRegistry& registry, const EvalContext& ctx) {
  return collect_numeric_pairs(lead_arg, trail_arg, arena, registry, ctx, /*allow_transpose=*/true);
}

// Mean and the three sums-of-deviations the regression functions need:
//   sum_xx = Σ (x_i - mean_x)^2
//   sum_yy = Σ (y_i - mean_y)^2
//   sum_xy = Σ (x_i - mean_x)(y_i - mean_y)
struct RegressionStats {
  double mean_x;
  double mean_y;
  double sum_xx;
  double sum_yy;
  double sum_xy;
};

RegressionStats compute_regression_stats(const NumericPairs& p) noexcept {
  const std::size_t n = p.second.size();
  RegressionStats s{};
  if (n == 0) {
    return s;
  }
  double sum_x = 0.0;
  double sum_y = 0.0;
  for (std::size_t i = 0; i < n; ++i) {
    sum_x += p.second[i];
    sum_y += p.first[i];
  }
  const double dn = static_cast<double>(n);
  s.mean_x = sum_x / dn;
  s.mean_y = sum_y / dn;
  for (std::size_t i = 0; i < n; ++i) {
    const double dx = p.second[i] - s.mean_x;
    const double dy = p.first[i] - s.mean_y;
    s.sum_xx += dx * dx;
    s.sum_yy += dy * dy;
    s.sum_xy += dx * dy;
  }
  return s;
}

// Guards the final numeric result. Any NaN / infinity becomes `#NUM!`
// so the caller never surfaces a non-finite value to the user.
Value finite_number(double r) {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Shared front-end for every 2-arity regression lazy impl: arity check
// + pair collection. Returns either the error `Value` to surface (on
// the left of the variant) or the distilled pairs (on the right).
std::variant<Value, NumericPairs> prepare_pairs(const parser::AstNode& call, Arena& arena,
                                                const FunctionRegistry& registry, const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value{Value::error(ErrorCode::Value)};
  }
  return collect_regression_pairs(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx);
}

// Computes slope / intercept together since INTERCEPT is just
// `mean_y - slope * mean_x`. Returns `false` on a degenerate data set
// (n < 2 or sum_xx == 0) with `#DIV/0!` written to `*out_err`;
// otherwise writes the slope / intercept and returns `true`.
bool compute_slope_intercept(const NumericPairs& pairs, double* out_slope, double* out_intercept, Value* out_err) {
  if (pairs.second.size() < 2U) {
    *out_err = Value::error(ErrorCode::Div0);
    return false;
  }
  const RegressionStats s = compute_regression_stats(pairs);
  if (s.sum_xx == 0.0) {
    *out_err = Value::error(ErrorCode::Div0);
    return false;
  }
  const double slope = s.sum_xy / s.sum_xx;
  *out_slope = slope;
  *out_intercept = s.mean_y - slope * s.mean_x;
  return true;
}

}  // namespace

Value eval_correl_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  auto prepared = prepare_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  // Pearson correlation is undefined for fewer than two points (the
  // sample variances collapse to zero) and when either marginal
  // variance is exactly zero (the denominator would be zero).
  if (pairs.second.size() < 2U) {
    return Value::error(ErrorCode::Div0);
  }
  const RegressionStats s = compute_regression_stats(pairs);
  if (s.sum_xx == 0.0 || s.sum_yy == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  return finite_number(s.sum_xy / std::sqrt(s.sum_xx * s.sum_yy));
}

Value eval_covariance_p_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  auto prepared = prepare_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  // Population covariance is defined for any n >= 1 (variance of a
  // single point is zero); only n == 0 is degenerate.
  if (pairs.second.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  const RegressionStats s = compute_regression_stats(pairs);
  return finite_number(s.sum_xy / static_cast<double>(pairs.second.size()));
}

Value eval_covariance_s_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  auto prepared = prepare_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  // Sample covariance uses divisor (n - 1); a single point yields 0/0.
  if (pairs.second.size() < 2U) {
    return Value::error(ErrorCode::Div0);
  }
  const RegressionStats s = compute_regression_stats(pairs);
  return finite_number(s.sum_xy / static_cast<double>(pairs.second.size() - 1U));
}

Value eval_slope_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  auto prepared = prepare_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  double slope = 0.0;
  double intercept = 0.0;
  Value err = Value::blank();
  if (!compute_slope_intercept(pairs, &slope, &intercept, &err)) {
    return err;
  }
  return finite_number(slope);
}

Value eval_intercept_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  auto prepared = prepare_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  double slope = 0.0;
  double intercept = 0.0;
  Value err = Value::blank();
  if (!compute_slope_intercept(pairs, &slope, &intercept, &err)) {
    return err;
  }
  return finite_number(intercept);
}

Value eval_rsq_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  auto prepared = prepare_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  if (pairs.second.size() < 2U) {
    return Value::error(ErrorCode::Div0);
  }
  const RegressionStats s = compute_regression_stats(pairs);
  if (s.sum_xx == 0.0 || s.sum_yy == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  // RSQ = CORREL^2 = sum_xy^2 / (sum_xx * sum_yy). Computing the ratio
  // directly avoids the intermediate sqrt in CORREL and stays closer
  // to the double-precision limit.
  return finite_number((s.sum_xy * s.sum_xy) / (s.sum_xx * s.sum_yy));
}

Value eval_steyx_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  auto prepared = prepare_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  // Residual standard error has (n - 2) degrees of freedom, so we need at
  // least 3 pairs. A collinear x-vector (sum_xx == 0) also makes the
  // regression undefined.
  if (pairs.second.size() < 3U) {
    return Value::error(ErrorCode::Div0);
  }
  const RegressionStats s = compute_regression_stats(pairs);
  if (s.sum_xx == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double residual_ss = s.sum_yy - (s.sum_xy * s.sum_xy) / s.sum_xx;
  // Floating-point subtraction can produce a tiny negative when the fit
  // is essentially exact; clamp to zero before taking the root.
  const double clamped = residual_ss < 0.0 ? 0.0 : residual_ss;
  return finite_number(std::sqrt(clamped / static_cast<double>(pairs.second.size() - 2U)));
}

namespace {

// Shared front-end for SUMX2PY2 / SUMX2MY2 / SUMXMY2. These take a pair
// of arrays in `(array_x, array_y)` order — the opposite of the rest of
// this file's `(known_y, known_x)` convention — so the impls cannot
// reuse `prepare_pairs` directly. Error propagation must still run in
// Excel's left-to-right order (array_x first), so we pass the arguments
// to `collect_numeric_pairs` in their declared order; that leaves
// array_x's cells in `pairs.first` and array_y's cells in `pairs.second`. The
// caller unpacks both fields with explicit local names to keep the
// subsequent arithmetic readable.
std::variant<Value, NumericPairs> prepare_sumx_pairs(const parser::AstNode& call, Arena& arena,
                                                     const FunctionRegistry& registry, const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value{Value::error(ErrorCode::Value)};
  }
  return collect_regression_pairs(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx);
}

}  // namespace

Value eval_sumx2py2_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  auto prepared = prepare_sumx_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  // Unpack with Excel-facing names: the first argument (array_x) lives in
  // `pairs.first`, and the second (array_y) lives in `pairs.second`. See
  // `prepare_sumx_pairs` for the reason.
  const std::vector<double>& x = pairs.first;
  const std::vector<double>& y = pairs.second;
  if (x.empty()) {
    return Value::error(ErrorCode::NA);
  }
  double total = 0.0;
  for (std::size_t i = 0; i < x.size(); ++i) {
    total += x[i] * x[i] + y[i] * y[i];
  }
  return finite_number(total);
}

Value eval_sumx2my2_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  auto prepared = prepare_sumx_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  const std::vector<double>& x = pairs.first;
  const std::vector<double>& y = pairs.second;
  if (x.empty()) {
    return Value::error(ErrorCode::NA);
  }
  double total = 0.0;
  for (std::size_t i = 0; i < x.size(); ++i) {
    total += x[i] * x[i] - y[i] * y[i];
  }
  return finite_number(total);
}

Value eval_sumxmy2_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  auto prepared = prepare_sumx_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  const std::vector<double>& x = pairs.first;
  const std::vector<double>& y = pairs.second;
  if (x.empty()) {
    return Value::error(ErrorCode::NA);
  }
  double total = 0.0;
  for (std::size_t i = 0; i < x.size(); ++i) {
    const double d = x[i] - y[i];
    total += d * d;
  }
  return finite_number(total);
}

Value eval_forecast_linear_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx) {
  if (call.as_call_arity() != 3U) {
    return Value::error(ErrorCode::Value);
  }
  // The first argument is a scalar x-value. Evaluate eagerly and
  // propagate any error — this is the only argument where a bare
  // number literal is valid.
  const Value x_val = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (x_val.is_error()) {
    return x_val;
  }
  if (!x_val.is_number()) {
    // Excel's FORECAST rejects non-numeric scalars with #VALUE!. A
    // Bool scalar is also rejected here because the function's
    // signature is explicitly numeric (Excel matches this behaviour).
    return Value::error(ErrorCode::Value);
  }
  const double x = x_val.as_number();

  const auto prepared = collect_regression_pairs(call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  double slope = 0.0;
  double intercept = 0.0;
  Value err = Value::blank();
  if (!compute_slope_intercept(pairs, &slope, &intercept, &err)) {
    return err;
  }
  return finite_number(intercept + slope * x);
}

Value eval_frequency_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value::error(ErrorCode::Value);
  }

  auto data_resolved = resolve_array_arg_na(call.as_call_arg(0), arena, registry, ctx);
  if (!data_resolved) {
    return Value::error(data_resolved.error());
  }
  RangeResult data_arr = std::move(data_resolved.value());
  auto bins_resolved = resolve_array_arg_na(call.as_call_arg(1), arena, registry, ctx);
  if (!bins_resolved) {
    return Value::error(bins_resolved.error());
  }
  RangeResult bins_arr = std::move(bins_resolved.value());

  // Error propagation: data_array errors propagate verbatim (leftmost
  // wins, row-major scan). bins_array errors are silently skipped — Mac
  // Excel treats error cells in the bin list as non-numeric and ignores
  // them, just like Blank / Bool / Text cells (verified against
  // FREQUENCY({1;2;3}, {2;#N/A}) returning the {2;#N/A} bins reduced to
  // [2] with count<=2 in slot 0).
  for (const Value& v : data_arr.cells) {
    if (v.is_error()) {
      return v;
    }
  }

  // Distil bins_array into a flat numeric vector.
  // Non-numeric cells (including Bool) drop out; Excel does not coerce
  // here. Excel buckets against numeric bins in ascending order, even
  // when the source bins_array is unsorted.
  std::vector<double> bins;
  bins.reserve(bins_arr.cells.size());
  for (const Value& v : bins_arr.cells) {
    if (v.is_number()) {
      bins.push_back(v.as_number());
    }
  }
  std::sort(bins.begin(), bins.end());
  const std::size_t n_bins = bins.size();

  if (bins_arr.cells.empty()) {
    return Value::blank();
  }

  // Allocate the count buffer. Even with zero numeric bins the result is
  // a 1x1 array containing the total numeric data count (matches Mac
  // Excel's documented degenerate case for empty bins).
  std::vector<std::uint64_t> counts(n_bins + 1U, 0U);

  // Walk data_array. For each numeric cell, find the first bin index i
  // where value <= bins[i]; if none satisfies, drop into the trailing
  // extra slot.
  for (const Value& v : data_arr.cells) {
    if (!v.is_number()) {
      continue;
    }
    const double x = v.as_number();
    bool placed = false;
    for (std::size_t i = 0; i < n_bins; ++i) {
      if (x <= bins[i]) {
        counts[i] += 1U;
        placed = true;
        break;
      }
    }
    if (!placed) {
      counts[n_bins] += 1U;
    }
  }

  // Materialise as a column ArrayValue. With zero numeric bins, this is
  // a 1x1 array; otherwise (n_bins + 1) x 1.
  const std::uint32_t out_rows = static_cast<std::uint32_t>(n_bins + 1U);
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(out_rows, 1U, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t i = 0; i < out_rows; ++i) {
    buffer[i] = Value::number(static_cast<double>(counts[i]));
  }
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
