// Copyright 2026 libraz. Licensed under the MIT License.
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
// file is modelled on; the shape-resolution helpers are deliberately
// kept in-TU so neither side pays for the other's specialisation.

#include "eval/regression_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <variant>
#include <vector>

#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// One resolved array argument: flat row-major cells plus the rectangle
// shape used by the shape-match check.
struct ResolvedArray {
  std::uint32_t rows;
  std::uint32_t cols;
  std::vector<Value> cells;
};

// Resolves a single array argument. Accepts `Ref`, `RangeOp`, and
// `ArrayLiteral`; any other shape (a scalar literal, a function call,
// arithmetic) yields `#N/A` because Excel's regression family uses
// `#N/A` for shape errors rather than the `#VALUE!` used by
// SUMPRODUCT. Returns `true` on success; on failure writes the Excel
// error into `*out_err` and returns `false`.
bool resolve_array_arg(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx, ResolvedArray* out, Value* out_err) {
  const parser::NodeKind k = arg_node.kind();
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    ErrorCode err_code = ErrorCode::NA;
    if (!resolve_range_arg(arg_node, arena, registry, ctx, &out->cells, &err_code, &out->rows, &out->cols)) {
      // `resolve_range_arg` reports `#VALUE!` for non-Ref / non-RangeOp
      // shapes and `#REF!` for expansion failures. For the regression
      // family we remap the shape-rejection case to `#N/A` to match
      // Excel; `#REF!` passes through unchanged.
      *out_err = Value::error(err_code == ErrorCode::Value ? ErrorCode::NA : err_code);
      return false;
    }
    return true;
  }
  if (k == parser::NodeKind::ArrayLiteral) {
    out->rows = arg_node.as_array_rows();
    out->cols = arg_node.as_array_cols();
    const std::size_t total = static_cast<std::size_t>(out->rows) * out->cols;
    out->cells.clear();
    out->cells.reserve(total);
    for (std::uint32_t r = 0; r < out->rows; ++r) {
      for (std::uint32_t c = 0; c < out->cols; ++c) {
        Value v = eval_node(arg_node.as_array_element(r, c), arena, registry, ctx);
        out->cells.push_back(v);
      }
    }
    return true;
  }
  // A scalar / call / arithmetic subtree is not a valid array here.
  // Evaluate it first so an error in the subtree propagates with its
  // real code; otherwise reject with `#N/A`.
  const Value v = eval_node(arg_node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  *out_err = Value::error(ErrorCode::NA);
  return false;
}

// Paired (x, y) numeric samples distilled from the two resolved array
// arguments.
struct NumericPairs {
  std::vector<double> x;
  std::vector<double> y;
};

// Resolves both array arguments, enforces shape match, propagates
// errors in row-major scan order (y-array first, then x-array, matching
// Excel's left-to-right rule), and collects every pair whose *both*
// cells are numeric. If either cell in a pair is non-numeric the whole
// pair is dropped — the dropping rule must not misalign the two
// sequences.
//
// Returns the error `Value` to propagate on the left side of the
// variant; otherwise a populated `NumericPairs` on the right.
std::variant<Value, NumericPairs> collect_numeric_pairs(const parser::AstNode& y_arg, const parser::AstNode& x_arg,
                                                        Arena& arena, const FunctionRegistry& registry,
                                                        const EvalContext& ctx) {
  ResolvedArray y_arr{};
  Value err = Value::blank();
  if (!resolve_array_arg(y_arg, arena, registry, ctx, &y_arr, &err)) {
    return err;
  }
  ResolvedArray x_arr{};
  if (!resolve_array_arg(x_arg, arena, registry, ctx, &x_arr, &err)) {
    return err;
  }

  // Shape mismatch — the regression family uses #N/A (unlike
  // SUMPRODUCT, which reports #VALUE!).
  if (y_arr.rows != x_arr.rows || y_arr.cols != x_arr.cols) {
    return Value{Value::error(ErrorCode::NA)};
  }

  // Error propagation runs over every cell in both arrays, even cells
  // that would be dropped by the text-pair rule. Scan y first then x
  // so the leftmost-argument error wins, matching Excel's precedence.
  for (const Value& v : y_arr.cells) {
    if (v.is_error()) {
      return v;
    }
  }
  for (const Value& v : x_arr.cells) {
    if (v.is_error()) {
      return v;
    }
  }

  NumericPairs pairs;
  const std::size_t n = y_arr.cells.size();
  pairs.x.reserve(n);
  pairs.y.reserve(n);
  for (std::size_t i = 0; i < n; ++i) {
    const Value& y = y_arr.cells[i];
    const Value& x = x_arr.cells[i];
    // Drop the pair if either side is non-numeric. Blank / Bool / Text
    // are all treated as "not a numeric sample" — matches Excel's
    // CORREL / COVAR / SLOPE behaviour on mixed ranges.
    if (!y.is_number() || !x.is_number()) {
      continue;
    }
    pairs.y.push_back(y.as_number());
    pairs.x.push_back(x.as_number());
  }
  return pairs;
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
  const std::size_t n = p.x.size();
  RegressionStats s{};
  if (n == 0) {
    return s;
  }
  double sum_x = 0.0;
  double sum_y = 0.0;
  for (std::size_t i = 0; i < n; ++i) {
    sum_x += p.x[i];
    sum_y += p.y[i];
  }
  const double dn = static_cast<double>(n);
  s.mean_x = sum_x / dn;
  s.mean_y = sum_y / dn;
  for (std::size_t i = 0; i < n; ++i) {
    const double dx = p.x[i] - s.mean_x;
    const double dy = p.y[i] - s.mean_y;
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
  return collect_numeric_pairs(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx);
}

// Computes slope / intercept together since INTERCEPT is just
// `mean_y - slope * mean_x`. Returns `false` on a degenerate data set
// (n < 2 or sum_xx == 0) with `#DIV/0!` written to `*out_err`;
// otherwise writes the slope / intercept and returns `true`.
bool compute_slope_intercept(const NumericPairs& pairs, double* out_slope, double* out_intercept, Value* out_err) {
  if (pairs.x.size() < 2U) {
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
  if (pairs.x.size() < 2U) {
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
  if (pairs.x.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  const RegressionStats s = compute_regression_stats(pairs);
  return finite_number(s.sum_xy / static_cast<double>(pairs.x.size()));
}

Value eval_covariance_s_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  auto prepared = prepare_pairs(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  // Sample covariance uses divisor (n - 1); a single point yields 0/0.
  if (pairs.x.size() < 2U) {
    return Value::error(ErrorCode::Div0);
  }
  const RegressionStats s = compute_regression_stats(pairs);
  return finite_number(s.sum_xy / static_cast<double>(pairs.x.size() - 1U));
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
  if (pairs.x.size() < 2U) {
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
  if (pairs.x.size() < 3U) {
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
  return finite_number(std::sqrt(clamped / static_cast<double>(pairs.x.size() - 2U)));
}

namespace {

// Shared front-end for SUMX2PY2 / SUMX2MY2 / SUMXMY2. These take a pair
// of arrays in `(array_x, array_y)` order — the opposite of the rest of
// this file's `(known_y, known_x)` convention — so the impls cannot
// reuse `prepare_pairs` directly. Error propagation must still run in
// Excel's left-to-right order (array_x first), so we pass the arguments
// to `collect_numeric_pairs` in their declared order; that leaves
// array_x's cells in `pairs.y` and array_y's cells in `pairs.x`. The
// caller unpacks both fields with explicit local names to keep the
// subsequent arithmetic readable.
std::variant<Value, NumericPairs> prepare_sumx_pairs(const parser::AstNode& call, Arena& arena,
                                                     const FunctionRegistry& registry, const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value{Value::error(ErrorCode::Value)};
  }
  return collect_numeric_pairs(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx);
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
  // `pairs.y`, and the second (array_y) lives in `pairs.x`. See
  // `prepare_sumx_pairs` for the reason.
  const std::vector<double>& x = pairs.y;
  const std::vector<double>& y = pairs.x;
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
  const std::vector<double>& x = pairs.y;
  const std::vector<double>& y = pairs.x;
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
  const std::vector<double>& x = pairs.y;
  const std::vector<double>& y = pairs.x;
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

  const auto prepared = collect_numeric_pairs(call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx);
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

}  // namespace eval
}  // namespace formulon
