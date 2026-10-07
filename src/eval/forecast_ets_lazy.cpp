//
// Implementation of the FORECAST.ETS family lazy impls.
//
// All four functions share a Holt-Winters additive triple-exponential
// smoothing core (`eval/holt_winters.h`) over a series prepared by
// `eval/ets_series.h`. The pipeline per call is:
//
//   1. Materialise both array arguments (`values`, `timeline`) via
//      `eval_node_as_array` so RangeOp / Ref / ArrayLiteral / arithmetic
//      broadcast subtrees all collapse to a flat `ArrayValue`.
//   2. Coerce every timeline / values cell to a number, propagating the
//      leftmost error in scan order (timeline first, then values) and
//      rejecting non-numeric / non-blank cells with `#VALUE!`.
//   3. Pair (t_i, y_i), stable-sort by t_i ascending, aggregate runs of
//      identical t_i via the user-selected mode (default AVG).
//   4. Detect the median delta-t step, validate that every consecutive
//      delta lies within +/-30% of the median, and resample the series
//      onto an evenly spaced grid; gaps are interpolated linearly when
//      `data_completion = 1` or filled with zero when `= 0`.
//   5. Auto-detect or accept the user-provided seasonality length `m`.
//   6. Initialise (L_0, B_0, S_0..S_{m-1}) from the first one or two
//      seasons (the "two-season-means" method).
//   7. Optimise (alpha, beta[, gamma]) by bounded Nelder-Mead on the
//      in-sample one-step-ahead SSE.
//   8. Read off the requested output: forecast, confidence half-width,
//      detected seasonality, or one of eight diagnostic statistics.
//
// `coerce_to_number` follows the wider Mac Excel coercion contract; that
// is intentionally NOT used for timeline / values cells here because Mac
// Excel's FORECAST.ETS rejects text in either array with `#VALUE!`,
// matching LINEST's strict-matrix rule rather than SUM's permissive
// "skip non-numerics" rule. See `is_strictly_numeric` below.
//
// Oracle-pending calibration: every constant flagged with the comment
// "ORACLE-PENDING" should be re-validated against the macOS Excel 365
// golden corpus in a follow-up commit. The structural behaviour (error
// codes, monotonicity in `h` / `confidence`, stability of the optimiser)
// is locked in here; the exact numeric outputs are not.

#include "eval/forecast_ets_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/builtins/numeric_helpers.h"
#include "eval/builtins/stats/stats_helpers.h"
#include "eval/coerce.h"
#include "eval/ets_series.h"
#include "eval/eval_context.h"
#include "eval/holt_winters.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::to_finite_value;

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

// Maximum manually-specified seasonality. Above this the third argument
// is rejected with `#NUM!` (matches Excel's documented 8760 limit).
constexpr std::uint32_t kMaxSeasonalityArg = 8760U;

using ets::AggregationMode;
using ets::detect_seasonality;
using ets::fit_holt_winters;
using ets::HoltWintersFit;
using ets::median_delta;
using ets::resample_to_grid;
using ets::Series;
using ets::sort_and_aggregate;

// ---------------------------------------------------------------------------
// Array materialisation + strict numeric coercion
// ---------------------------------------------------------------------------

// Strict numeric reading for timeline / values cells.
// Numbers pass through; Booleans coerce to 1/0; Blank fails as `#VALUE!`
// (FORECAST.ETS treats blanks in the data series as missing -- but the
// data-completion fill happens AFTER pairing, so a blank cell that
// survives the array materialisation is genuinely missing data and we
// reject the call). Errors propagate verbatim. Text fails as `#VALUE!`
// matching Excel's strict-matrix rule.
//
// Returns `true` on success with the number written to `*out`, otherwise
// writes the propagating error into `*out_err`.
bool coerce_strict_numeric(const Value& v, double* out, Value* out_err) {
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  if (v.is_number()) {
    *out = v.as_number();
    return true;
  }
  if (v.is_boolean()) {
    *out = v.as_boolean() ? 1.0 : 0.0;
    return true;
  }
  // Blank / Text / Array / Lambda / Ref are all rejected.
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

// ---------------------------------------------------------------------------
// Front-end: shared input preprocessing
// ---------------------------------------------------------------------------

// Reads an optional scalar number argument. Blank yields `default_value`;
// errors propagate.
bool read_double_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, double default_value, double* out, Value* out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  if (v.is_blank()) {
    *out = default_value;
    return true;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    *out_err = Value::error(coerced.error());
    return false;
  }
  *out = coerced.value();
  return true;
}

// Reads an optional scalar argument as a non-negative integer (truncated
// toward zero). Same semantics as `read_double_arg`; a non-finite value is
// flagged via `*out_err` with the supplied error code.
bool read_int_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                  std::int64_t default_value, std::int64_t* out, Value* out_err, ErrorCode oo_domain) {
  double d = 0.0;
  if (!read_double_arg(node, arena, registry, ctx, static_cast<double>(default_value), &d, out_err)) {
    return false;
  }
  if (!std::isfinite(d)) {
    *out_err = Value::error(oo_domain);
    return false;
  }
  *out = static_cast<std::int64_t>(d);  // truncate toward zero
  return true;
}

bool read_optional_int_arg(const parser::AstNode* node, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, std::int64_t default_value, std::int64_t* out, Value* out_err,
                           ErrorCode domain_error) {
  if (node == nullptr) {
    *out = default_value;
    return true;
  }
  return read_int_arg(*node, arena, registry, ctx, default_value, out, out_err, domain_error);
}

bool read_required_finite_number(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx, double* out, Value* out_err, ErrorCode non_finite_error) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    *out_err = Value::error(coerced.error());
    return false;
  }
  const double n = coerced.value();
  if (!std::isfinite(n)) {
    *out_err = Value::error(non_finite_error);
    return false;
  }
  *out = n;
  return true;
}

bool read_required_truncated_int(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx, std::int64_t* out, Value* out_err,
                                 ErrorCode non_finite_error) {
  double n = 0.0;
  if (!read_required_finite_number(node, arena, registry, ctx, &n, out_err, non_finite_error)) {
    return false;
  }
  *out = static_cast<std::int64_t>(n);
  return true;
}

// Aggregated state from the preprocessing pipeline. Returned to each
// front-end after timeline / values have been paired, sorted, aggregated,
// and resampled onto an evenly spaced grid.
struct Preprocessed {
  Series resampled;      // evenly spaced (t, y) pairs.
  double step = 0.0;     // grid step (median delta-t of original timeline).
  double t0 = 0.0;       // first original timeline value (== resampled.t[0]).
  std::uint32_t m = 1U;  // resolved seasonality length.
};

struct ForecastOptions {
  std::int64_t seasonality = 1;
  int data_completion = 1;
  AggregationMode aggregation = AggregationMode::kAverage;
};

// Runs the timeline / values preprocessing pipeline. Inputs:
//   * `values_node`, `timeline_node` -- the two array AST args.
//   * `seasonality` -- 0 (force non-seasonal), 1 (auto-detect), or
//     >= 2 manual override (capped at kMaxSeasonalityArg).
//   * `data_completion` -- 0 (zero-fill) or 1 (linear interpolation).
//   * `aggregation` -- 1..7 enumeration.
// On success populates `*out` and returns `true`. On failure writes the
// propagating error into `*out_err`.
bool preprocess(const parser::AstNode& values_node, const parser::AstNode& timeline_node, Arena& arena,
                const FunctionRegistry& registry, const EvalContext& ctx, std::int64_t seasonality, int data_completion,
                AggregationMode aggregation, Preprocessed* out, Value* out_err) {
  // Materialise both arrays.
  const ArrayValue* timeline_arr = nullptr;
  const ArrayValue* values_arr = nullptr;
  if (!resolve_array_value(timeline_node, arena, registry, ctx, &timeline_arr, out_err)) {
    return false;
  }
  if (!resolve_array_value(values_node, arena, registry, ctx, &values_arr, out_err)) {
    return false;
  }
  const std::size_t n_t = static_cast<std::size_t>(timeline_arr->rows) * timeline_arr->cols;
  const std::size_t n_v = static_cast<std::size_t>(values_arr->rows) * values_arr->cols;
  if (n_t != n_v) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  if (n_t == 0U) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  if (n_t < 2U) {
    // Single-point series: Mac Excel 365 surfaces #DIV/0! here (verified
    // against the oracle for `=FORECAST.ETS(2, {10}, {1}, 0)`), distinct
    // from the #N/A reserved for length-mismatch.
    *out_err = Value::error(ErrorCode::Div0);
    return false;
  }

  // Error propagation: timeline first, then values.
  for (std::size_t i = 0; i < n_t; ++i) {
    if (timeline_arr->cells[i].is_error()) {
      *out_err = timeline_arr->cells[i];
      return false;
    }
  }
  for (std::size_t i = 0; i < n_v; ++i) {
    if (values_arr->cells[i].is_error()) {
      *out_err = values_arr->cells[i];
      return false;
    }
  }

  // Strict numeric coercion.
  Series raw;
  raw.t.resize(n_t);
  raw.y.resize(n_v);
  for (std::size_t i = 0; i < n_t; ++i) {
    if (!coerce_strict_numeric(timeline_arr->cells[i], &raw.t[i], out_err)) {
      return false;
    }
  }
  for (std::size_t i = 0; i < n_v; ++i) {
    if (!coerce_strict_numeric(values_arr->cells[i], &raw.y[i], out_err)) {
      return false;
    }
  }

  // Sort + aggregate.
  Series agg = sort_and_aggregate(raw, aggregation);
  if (agg.t.size() < 2U) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  // Reject any non-positive step in the sorted timeline (defensive --
  // duplicates have been collapsed by aggregation, so consecutive deltas
  // must be strictly positive in a well-formed timeline).
  for (std::size_t i = 1; i < agg.t.size(); ++i) {
    const double d = agg.t[i] - agg.t[i - 1U];
    if (d <= 0.0 || !std::isfinite(d)) {
      *out_err = Value::error(ErrorCode::Num);
      return false;
    }
  }

  const double step = median_delta(agg.t);
  if (step <= 0.0 || !std::isfinite(step)) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }

  Series resampled;
  if (!resample_to_grid(agg, step, data_completion, &resampled, out_err)) {
    return false;
  }
  if (resampled.y.size() < 2U) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }

  // Resolve seasonality.
  std::uint32_t m = 1U;
  if (seasonality == 0) {
    m = 1U;
  } else if (seasonality == 1) {
    m = detect_seasonality(resampled.y);
  } else if (seasonality >= 2 && seasonality <= static_cast<std::int64_t>(kMaxSeasonalityArg)) {
    m = static_cast<std::uint32_t>(seasonality);
  } else {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }

  out->resampled = std::move(resampled);
  out->step = step;
  out->t0 = agg.t.front();
  out->m = m;
  return true;
}

// Reads the optional [aggregation] argument and validates it. Returns
// `true` on success with the parsed mode; otherwise writes the error.
bool read_aggregation(const parser::AstNode* node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, AggregationMode* out, Value* out_err) {
  std::int64_t v = 1;
  if (!read_optional_int_arg(node, arena, registry, ctx, 1, &v, out_err, ErrorCode::Value)) {
    return false;
  }
  if (v < 1 || v > 7) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *out = static_cast<AggregationMode>(v);
  return true;
}

// Reads the optional [data_completion] argument and validates it. The
// allowed domain is exactly {0, 1}; anything else is `#VALUE!`.
bool read_data_completion(const parser::AstNode* node, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx, int* out, Value* out_err) {
  std::int64_t v = 1;
  if (!read_optional_int_arg(node, arena, registry, ctx, 1, &v, out_err, ErrorCode::Value)) {
    return false;
  }
  if (v != 0 && v != 1) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *out = static_cast<int>(v);
  return true;
}

// Reads the optional [seasonality] argument. Domain: 0 (force non-
// seasonal), 1 (auto-detect, default), or 2..kMaxSeasonalityArg.
bool read_seasonality(const parser::AstNode* node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, std::int64_t* out, Value* out_err) {
  std::int64_t v = 1;
  if (!read_optional_int_arg(node, arena, registry, ctx, 1, &v, out_err, ErrorCode::Num)) {
    return false;
  }
  if (v < 0 || v > static_cast<std::int64_t>(kMaxSeasonalityArg)) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  *out = v;
  return true;
}

bool read_forecast_options(const parser::AstNode& call, std::uint32_t arity, std::uint32_t first_optional, Arena& arena,
                           const FunctionRegistry& registry, const EvalContext& ctx, ForecastOptions* out,
                           Value* out_err) {
  ForecastOptions opts;
  if (!read_seasonality(arity > first_optional ? &call.as_call_arg(first_optional) : nullptr, arena, registry, ctx,
                        &opts.seasonality, out_err)) {
    return false;
  }
  if (!read_data_completion(arity > first_optional + 1U ? &call.as_call_arg(first_optional + 1U) : nullptr, arena,
                            registry, ctx, &opts.data_completion, out_err)) {
    return false;
  }
  if (!read_aggregation(arity > first_optional + 2U ? &call.as_call_arg(first_optional + 2U) : nullptr, arena, registry,
                        ctx, &opts.aggregation, out_err)) {
    return false;
  }
  *out = opts;
  return true;
}

bool read_seasonality_options(const parser::AstNode& call, std::uint32_t arity, std::uint32_t first_optional,
                              Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                              ForecastOptions* out, Value* out_err) {
  ForecastOptions opts;
  opts.seasonality = 1;
  if (!read_data_completion(arity > first_optional ? &call.as_call_arg(first_optional) : nullptr, arena, registry, ctx,
                            &opts.data_completion, out_err)) {
    return false;
  }
  if (!read_aggregation(arity > first_optional + 1U ? &call.as_call_arg(first_optional + 1U) : nullptr, arena, registry,
                        ctx, &opts.aggregation, out_err)) {
    return false;
  }
  *out = opts;
  return true;
}

// Computes the integer step count h between the last training timeline
// value and `target_date`. Caller has already validated `target_date >=
// t0`. Returns the rounded step count; the caller may compare against
// the in-grid index of the last training point to derive the forecast
// horizon.
std::int64_t target_step_index(double target_date, double t0, double step) noexcept {
  return static_cast<std::int64_t>(std::round((target_date - t0) / step));
}

// Computes the seasonal index correction for a forecast at grid offset
// `k` relative to the first training point, given a training-series of
// length `n` and seasonality `m`. The Holt-Winters formula uses
// `S_{n - m + 1 + (h - 1) mod m}` where `h = k - (n - 1)` is the
// forecast horizon counted from the last training point.
double seasonal_correction(const std::vector<double>& season, std::uint32_t m, std::int64_t h) noexcept {
  if (m <= 1U || season.empty()) {
    return 0.0;
  }
  // Excel's formula: S_{n - m + 1 + (h - 1) mod m}, but our `season`
  // vector is indexed by `t mod m` not by absolute position. Equivalent
  // mapping: pick season slot `(n_minus_one + h) mod m` since the loop
  // in `simulate` writes season[t % m] for each t; thus the next slot
  // for t = n is `n % m` and we add `h - 1` to walk forward.
  // To keep this self-consistent with `simulate`'s indexing we work
  // directly in modular arithmetic over season-vector indices.
  const std::int64_t mm = static_cast<std::int64_t>(m);
  const std::int64_t shifted = ((h - 1) % mm + mm) % mm;
  return season[static_cast<std::size_t>(shifted)];
}

// Shared tail of FORECAST.ETS and FORECAST.ETS.CONFINT: reads the options
// from `first_optional`, preprocesses (values, timeline) from args 1-2,
// rejects `target_date < t0` with #NUM!, fits Holt-Winters, and writes
// the horizon h counted from the last training point.
bool fit_to_target(const parser::AstNode& call, std::uint32_t arity, std::uint32_t first_optional, Arena& arena,
                   const FunctionRegistry& registry, const EvalContext& ctx, double target_date,
                   HoltWintersFit* out_fit, std::int64_t* out_h, Value* out_err) {
  ForecastOptions opts;
  if (!read_forecast_options(call, arity, first_optional, arena, registry, ctx, &opts, out_err)) {
    return false;
  }

  Preprocessed pre;
  if (!preprocess(call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx, opts.seasonality,
                  opts.data_completion, opts.aggregation, &pre, out_err)) {
    return false;
  }

  // target_date must lie at or after the first timeline value.
  if (target_date < pre.t0) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }

  if (!fit_holt_winters(pre.resampled.y, pre.m, out_fit, out_err)) {
    return false;
  }

  // The last training point lives at grid index `n - 1`; `target_date`
  // lives at grid index `target_idx`.
  const std::int64_t n = static_cast<std::int64_t>(pre.resampled.y.size());
  const std::int64_t target_idx = target_step_index(target_date, pre.t0, pre.step);
  *out_h = target_idx - (n - 1);
  return true;
}

// ---------------------------------------------------------------------------
// Front-end impls
// ---------------------------------------------------------------------------

}  // namespace

Value eval_forecast_ets_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 6U) {
    return Value::error(ErrorCode::Value);
  }

  // arg 0: target_date scalar.
  Value err = Value::blank();
  double target_date = 0.0;
  if (!read_required_finite_number(call.as_call_arg(0), arena, registry, ctx, &target_date, &err, ErrorCode::Num)) {
    return err;
  }

  HoltWintersFit fit;
  std::int64_t h = 0;
  if (!fit_to_target(call, arity, 3U, arena, registry, ctx, target_date, &fit, &h, &err)) {
    return err;
  }

  // Negative h (target_date inside the training window) returns an
  // interpolated in-sample value computed as L + h*B + S correction.
  const double level_term = fit.level + static_cast<double>(h) * fit.trend;
  const double seasonal_term = seasonal_correction(fit.season, fit.m, h);
  const double forecast = level_term + seasonal_term;
  return to_finite_value(forecast);
}

Value eval_forecast_ets_confint_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 7U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();
  double target_date = 0.0;
  if (!read_required_finite_number(call.as_call_arg(0), arena, registry, ctx, &target_date, &err, ErrorCode::Num)) {
    return err;
  }

  double confidence = 0.95;
  if (arity > 3U) {
    if (!read_double_arg(call.as_call_arg(3), arena, registry, ctx, 0.95, &confidence, &err)) {
      return err;
    }
  }
  // Mac Excel 365 accepts confidence == 0 (degenerate CI, returns 0 since
  // z = InverseStandardNormal(0.5) = 0). The lower-bound rejection is
  // strict (< 0); upper bound stays inclusive on 1 because z diverges to
  // +infinity there.
  if (!std::isfinite(confidence) || confidence < 0.0 || confidence >= 1.0) {
    return Value::error(ErrorCode::Num);
  }

  HoltWintersFit fit;
  std::int64_t h = 0;
  if (!fit_to_target(call, arity, 4U, arena, registry, ctx, target_date, &fit, &h, &err)) {
    return err;
  }
  // Mac Excel 365 rejects target_date inside the training window with
  // #NUM!. The half-width formula z * RMSE * sqrt(h) is only meaningful
  // for strictly positive horizons.
  if (h < 1) {
    return Value::error(ErrorCode::Num);
  }

  // Simplified normal-approximation half-width:
  //     hw = z * RMSE * sqrt(h)
  // where z is the inverse standard-normal CDF at (1 + confidence) / 2.
  // ORACLE-PENDING: the Hyndman recursion `sqrt(1 + sum_k)` may need to
  // be substituted if oracle parity demands it.
  const double tail_prob = (1.0 + confidence) * 0.5;
  const double z = stats_detail::InverseStandardNormal(tail_prob);
  const double hw = z * fit.rmse * std::sqrt(static_cast<double>(h));
  return to_finite_value(hw);
}

Value eval_forecast_ets_seasonality_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                         const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();
  ForecastOptions opts;
  if (!read_seasonality_options(call, arity, 2U, arena, registry, ctx, &opts, &err)) {
    return err;
  }

  // Force auto-detect by passing seasonality = 1 to preprocess; the
  // resampling path is unchanged.
  Preprocessed pre;
  if (!preprocess(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx, opts.seasonality,
                  opts.data_completion, opts.aggregation, &pre, &err)) {
    return err;
  }
  // Mac Excel 365 reports 0 (not 1) when no period is detected. Internally
  // m = 1 means non-seasonal for the Holt-Winters fit; map it to 0 at the
  // SEASONALITY API boundary.
  return Value::number(static_cast<double>(pre.m == 1U ? 0U : pre.m));
}

Value eval_forecast_ets_stat_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                  const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 6U) {
    return Value::error(ErrorCode::Value);
  }

  // arg 2: statistic_type scalar.
  Value err = Value::blank();
  std::int64_t stat_type = 0;
  if (!read_required_truncated_int(call.as_call_arg(2), arena, registry, ctx, &stat_type, &err, ErrorCode::Num)) {
    return err;
  }
  if (stat_type < 1 || stat_type > 8) {
    return Value::error(ErrorCode::Num);
  }

  ForecastOptions opts;
  if (!read_forecast_options(call, arity, 3U, arena, registry, ctx, &opts, &err)) {
    return err;
  }

  Preprocessed pre;
  if (!preprocess(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx, opts.seasonality,
                  opts.data_completion, opts.aggregation, &pre, &err)) {
    return err;
  }

  // statistic_type = 8 (step_size) does not require a fit. Return the
  // grid step directly. ORACLE-PENDING: the exact "step_size" definition
  // (median delta-t of the original timeline before vs after aggregation)
  // is calibrated below at the post-aggregation median; verify against
  // oracle.
  if (stat_type == 8) {
    return Value::number(pre.step);
  }

  HoltWintersFit fit;
  if (!fit_holt_winters(pre.resampled.y, pre.m, &fit, &err)) {
    return err;
  }

  double result = 0.0;
  switch (stat_type) {
    case 1:
      result = fit.alpha;
      break;
    case 2:
      result = fit.beta;
      break;
    case 3:
      result = fit.gamma;
      break;
    case 4:
      result = fit.mase;
      break;
    case 5:
      result = fit.smape;
      break;
    case 6:
      result = fit.mae;
      break;
    case 7:
      result = fit.rmse;
      break;
    default:
      return Value::error(ErrorCode::Num);
  }
  return to_finite_value(result);
}

}  // namespace eval
}  // namespace formulon
