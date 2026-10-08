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
#include <optional>
#include <utility>
#include <variant>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/builtins/numeric_helpers.h"
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

using builtins_detail::to_finite_value;

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
// `mean_y - slope * mean_x`. Returns `false` with the collection error,
// or `#DIV/0!` on a degenerate data set (n < 2 or sum_xx == 0), written
// to `*out_err`; otherwise writes the slope / intercept and returns `true`.
bool compute_slope_intercept(const std::variant<Value, NumericPairs>& prepared, double* out_slope,
                             double* out_intercept, Value* out_err) {
  if (std::holds_alternative<Value>(prepared)) {
    *out_err = std::get<Value>(prepared);
    return false;
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
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

// COVARIANCE.P / COVARIANCE.S given two scalar expressions treat each as a
// one-element array. Returns the single pair (a non-numeric scalar leaves
// none), or an error Value when either argument evaluates to an error;
// nullopt when either argument is not a scalar expression.
std::optional<std::variant<Value, NumericPairs>> scalar_covariance_pairs(const parser::AstNode& call, Arena& arena,
                                                                         const FunctionRegistry& registry,
                                                                         const EvalContext& ctx) {
  RangeResult scalars[2];
  for (std::uint32_t i = 0; i < 2U; ++i) {
    auto resolved = resolve_range_arg(call.as_call_arg(i), arena, registry, ctx);
    if (!resolved || !resolved.value().from_scalar || resolved.value().cells.size() != 1U) {
      return std::nullopt;
    }
    scalars[i] = std::move(resolved.value());
  }
  for (const RangeResult& r : scalars) {
    if (r.cells.front().is_error()) {
      return std::variant<Value, NumericPairs>{r.cells.front()};
    }
  }
  NumericPairs pairs;
  if (scalars[0].cells.front().is_number() && scalars[1].cells.front().is_number()) {
    pairs.first.push_back(scalars[0].cells.front().as_number());
    pairs.second.push_back(scalars[1].cells.front().as_number());
  }
  return std::variant<Value, NumericPairs>{std::move(pairs)};
}

enum class PairStat {
  Correl,
  CovarianceP,
  CovarianceS,
  Slope,
  Intercept,
  Rsq,
  Steyx,
  SumX2PY2,
  SumX2MY2,
  SumXMY2,
};

// Shared driver for the 2-arity pairwise statistics: prepares the pairs once, then reduces per `kind`.
Value eval_pair_stat(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, PairStat kind) {
  auto prepared = prepare_pairs(call, arena, registry, ctx);
  if ((kind == PairStat::CovarianceP || kind == PairStat::CovarianceS) && std::holds_alternative<Value>(prepared) &&
      std::get<Value>(prepared).is_error() && std::get<Value>(prepared).as_error() == ErrorCode::NA) {
    if (auto scalar = scalar_covariance_pairs(call, arena, registry, ctx)) {
      prepared = std::move(*scalar);
    }
  }
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  const std::size_t n = pairs.second.size();
  switch (kind) {
    case PairStat::Slope:
    case PairStat::Intercept: {
      double slope = 0.0;
      double intercept = 0.0;
      Value err = Value::blank();
      if (!compute_slope_intercept(prepared, &slope, &intercept, &err)) {
        return err;
      }
      return to_finite_value(kind == PairStat::Slope ? slope : intercept);
    }
    case PairStat::SumX2PY2:
    case PairStat::SumX2MY2:
    case PairStat::SumXMY2: {
      // Collection runs left to right: array_x lands in `pairs.first`, array_y in `pairs.second`.
      const std::vector<double>& x = pairs.first;
      const std::vector<double>& y = pairs.second;
      if (x.empty()) {
        return Value::error(ErrorCode::NA);
      }
      double total = 0.0;
      for (std::size_t i = 0; i < x.size(); ++i) {
        if (kind == PairStat::SumX2PY2) {
          total += x[i] * x[i] + y[i] * y[i];
        } else if (kind == PairStat::SumX2MY2) {
          total += x[i] * x[i] - y[i] * y[i];
        } else {
          const double d = x[i] - y[i];
          total += d * d;
        }
      }
      return to_finite_value(total);
    }
    default:
      break;
  }
  // Pearson / RSQ need n >= 2, COVARIANCE.P needs n >= 1, STEYX needs n >= 3 (n - 2 degrees of freedom).
  const std::size_t min_n = kind == PairStat::Steyx ? 3U : (kind == PairStat::CovarianceP ? 1U : 2U);
  if (n < min_n) {
    return Value::error(ErrorCode::Div0);
  }
  const RegressionStats s = compute_regression_stats(pairs);
  switch (kind) {
    case PairStat::Correl:
      // A zero marginal variance makes the denominator zero.
      if (s.sum_xx == 0.0 || s.sum_yy == 0.0) {
        return Value::error(ErrorCode::Div0);
      }
      return to_finite_value(s.sum_xy / std::sqrt(s.sum_xx * s.sum_yy));
    case PairStat::CovarianceP:
      return to_finite_value(s.sum_xy / static_cast<double>(n));
    case PairStat::CovarianceS:
      return to_finite_value(s.sum_xy / static_cast<double>(n - 1U));
    case PairStat::Rsq:
      if (s.sum_xx == 0.0 || s.sum_yy == 0.0) {
        return Value::error(ErrorCode::Div0);
      }
      // CORREL^2 computed directly avoids the intermediate sqrt.
      return to_finite_value((s.sum_xy * s.sum_xy) / (s.sum_xx * s.sum_yy));
    default: {  // Steyx
      if (s.sum_xx == 0.0) {
        return Value::error(ErrorCode::Div0);
      }
      const double residual_ss = s.sum_yy - (s.sum_xy * s.sum_xy) / s.sum_xx;
      // Rounding can leave a tiny negative on an exact fit; clamp before the root.
      const double clamped = residual_ss < 0.0 ? 0.0 : residual_ss;
      return to_finite_value(std::sqrt(clamped / static_cast<double>(n - 2U)));
    }
  }
}

}  // namespace

Value eval_correl_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  return eval_pair_stat(call, arena, registry, ctx, PairStat::Correl);
}

Value eval_covariance_p_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  return eval_pair_stat(call, arena, registry, ctx, PairStat::CovarianceP);
}

Value eval_covariance_s_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  return eval_pair_stat(call, arena, registry, ctx, PairStat::CovarianceS);
}

Value eval_slope_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return eval_pair_stat(call, arena, registry, ctx, PairStat::Slope);
}

Value eval_intercept_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  return eval_pair_stat(call, arena, registry, ctx, PairStat::Intercept);
}

Value eval_rsq_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  return eval_pair_stat(call, arena, registry, ctx, PairStat::Rsq);
}

Value eval_steyx_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return eval_pair_stat(call, arena, registry, ctx, PairStat::Steyx);
}

Value eval_sumx2py2_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  return eval_pair_stat(call, arena, registry, ctx, PairStat::SumX2PY2);
}

Value eval_sumx2my2_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  return eval_pair_stat(call, arena, registry, ctx, PairStat::SumX2MY2);
}

Value eval_sumxmy2_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  return eval_pair_stat(call, arena, registry, ctx, PairStat::SumXMY2);
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
  double slope = 0.0;
  double intercept = 0.0;
  Value err = Value::blank();
  if (!compute_slope_intercept(prepared, &slope, &intercept, &err)) {
    return err;
  }
  return to_finite_value(intercept + slope * x);
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
