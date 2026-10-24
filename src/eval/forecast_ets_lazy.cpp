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
//   2. Reduce each cell to a number / blank / text / boolean / error and
//      hand both arrays to `ets::prepare_series`, which pairs and orders
//      them, detects the timeline step and regularises the series.
//   3. Resolve the seasonality length `m` (none, given, or detected).
//   4. Fit the additive Holt-Winters model: the initial state comes from a
//      moving-average decomposition and (alpha, beta[, gamma]) from a fixed
//      candidate table scored by a holdout-aware SSE.
//   5. Read off the requested output: forecast, confidence half-width,
//      detected seasonality, or one of eight diagnostic statistics.
//

#include "eval/forecast_ets_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <utility>
#include <vector>

#include "eval/array_alloc.h"
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

// Largest accepted seasonality argument; above this it is `#NUM!`.
constexpr std::int64_t kMaxSeasonalityArg = 8784;

// A seasonality argument at or above this percentage of the series length
// is treated as no seasonality.
constexpr std::size_t kSeasonalityLengthPercent = 70U;

// Largest forecast horizon in steps past the last node; Mac Excel 365 returns
// `#NUM!` for a target beyond it (2147483659 answers, 2147483660 does not, on
// a 1..12 timeline).
constexpr double kMaxHorizonSteps = 2147483647.0;

using ets::AggregationMode;
using ets::detect_seasonality;
using ets::fit_holt_winters;
using ets::HoltWintersFit;
using ets::prepare_series;
using ets::PreparedSeries;
using ets::SeriesCell;
using ets::target_position;
using ets::Timeline;
using ets::whole_count;

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

// Reads an optional option argument that must be an integer in [lo, hi];
// anything else is `#NUM!`. A missing or blank argument yields `default_value`.
bool read_option(const parser::AstNode* node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                 std::int64_t default_value, std::int64_t lo, std::int64_t hi, std::int64_t* out, Value* out_err) {
  double d = static_cast<double>(default_value);
  if (node != nullptr && !read_double_arg(*node, arena, registry, ctx, d, &d, out_err)) {
    return false;
  }
  const std::optional<std::int64_t> v = whole_count(d);
  if (!v || d != std::floor(d) || *v < lo || *v > hi) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  *out = *v;
  return true;
}

// Coerces one scalar argument value to a number; a non-finite result is
// `non_finite_error`.
bool finite_number(const Value& v, double* out, Value* out_err, ErrorCode non_finite_error) {
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

// Reads a required scalar truncated to an integer; a value outside the
// `whole_count` range is `non_finite_error` like a non-finite one.
bool read_required_truncated_int(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx, std::int64_t* out, Value* out_err,
                                 ErrorCode non_finite_error) {
  double n = 0.0;
  if (!finite_number(eval_node(node, arena, registry, ctx), &n, out_err, non_finite_error)) {
    return false;
  }
  const std::optional<std::int64_t> v = whole_count(n);
  if (!v) {
    *out_err = Value::error(non_finite_error);
    return false;
  }
  *out = *v;
  return true;
}

struct ForecastOptions {
  std::int64_t seasonality = 1;
  int data_completion = 1;
  AggregationMode aggregation = AggregationMode::kAverage;
};

// A prepared series with its resolved seasonality length (1 = none).
struct Prepared {
  PreparedSeries series;
  std::uint32_t m = 1U;
};

// Reduces an array of values to the cell kinds the series reader tells apart.
std::vector<SeriesCell> to_series_cells(const ArrayValue& arr) {
  const std::size_t n = static_cast<std::size_t>(arr.rows) * arr.cols;
  std::vector<SeriesCell> cells(n);
  for (std::size_t i = 0; i < n; ++i) {
    const Value& v = arr.cells[i];
    SeriesCell& c = cells[i];
    if (v.is_number()) {
      c.kind = SeriesCell::Kind::kNumber;
      c.number = v.as_number();
    } else if (v.is_text()) {
      c.kind = SeriesCell::Kind::kText;
    } else if (v.is_boolean()) {
      c.kind = SeriesCell::Kind::kBool;
    } else if (v.is_error()) {
      c.kind = SeriesCell::Kind::kError;
      c.error = v.as_error();
    }
  }
  return cells;
}

// Seasonality length the fit runs with: detected for 1, none for 0, and none
// when the argument covers most of the series.
std::uint32_t resolve_seasonality(const std::vector<double>& y, std::int64_t seasonality) {
  if (seasonality == 1) {
    return detect_seasonality(y);
  }
  if (seasonality == 0 || static_cast<std::size_t>(seasonality) >= (kSeasonalityLengthPercent * y.size()) / 100U) {
    return 1U;
  }
  return static_cast<std::uint32_t>(seasonality);
}

// Materialises and regularises the (values, timeline) arguments, then resolves
// the seasonality length. On failure writes the error into `*out_err`.
bool prepare(const parser::AstNode& values_node, const parser::AstNode& timeline_node, Arena& arena,
             const FunctionRegistry& registry, const EvalContext& ctx, const ForecastOptions& opts, Prepared* out,
             Value* out_err) {
  const ArrayValue* timeline_arr = nullptr;
  const ArrayValue* values_arr = nullptr;
  if (!resolve_array_value(timeline_node, arena, registry, ctx, &timeline_arr, out_err)) {
    return false;
  }
  if (!resolve_array_value(values_node, arena, registry, ctx, &values_arr, out_err)) {
    return false;
  }
  ErrorCode err = ErrorCode::NA;
  if (!prepare_series(to_series_cells(*values_arr), to_series_cells(*timeline_arr), opts.data_completion,
                      opts.aggregation, &out->series, &err)) {
    *out_err = Value::error(err);
    return false;
  }
  out->m = resolve_seasonality(out->series.y, opts.seasonality);
  return true;
}

// Reads an integer option that must lie in [lo, hi]; anything else is `#NUM!`.
bool read_aggregation(const parser::AstNode* node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, AggregationMode* out, Value* out_err) {
  std::int64_t v = 1;
  if (!read_option(node, arena, registry, ctx, 1, 1, 7, &v, out_err)) {
    return false;
  }
  *out = static_cast<AggregationMode>(v);
  return true;
}

bool read_data_completion(const parser::AstNode* node, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx, int* out, Value* out_err) {
  std::int64_t v = 1;
  if (!read_option(node, arena, registry, ctx, 1, 0, 1, &v, out_err)) {
    return false;
  }
  *out = static_cast<int>(v);
  return true;
}

// Reads the optional [seasonality] argument: 0 (none), 1 (detect, default) or
// 2..kMaxSeasonalityArg.
bool read_seasonality(const parser::AstNode* node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, std::int64_t* out, Value* out_err) {
  return read_option(node, arena, registry, ctx, 1, 0, kMaxSeasonalityArg, out, out_err);
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

// Linear interpolation of the value at `pos` between steps `s` and `s + 1`.
double interpolate(std::int64_t s, double at_s, double at_next, double pos) {
  const double span = static_cast<double>(s + 1) - static_cast<double>(s);
  const double diff = at_next - at_s;
  double v = pos - static_cast<double>(s);
  v = v * diff;
  v = v / span;
  return v + at_s;
}

// A fitted model with the grid that places target dates on it.
struct FittedSeries {
  HoltWintersFit fit;
  Timeline timeline;
  std::size_t n = 0;
};

// Shared head of FORECAST.ETS and FORECAST.ETS.CONFINT: reads the options
// from `first_optional`, prepares the series from args 1-2 and fits it.
bool fit_series(const parser::AstNode& call, std::uint32_t arity, std::uint32_t first_optional, Arena& arena,
                const FunctionRegistry& registry, const EvalContext& ctx, FittedSeries* out, Value* out_err) {
  ForecastOptions opts;
  if (!read_forecast_options(call, arity, first_optional, arena, registry, ctx, &opts, out_err)) {
    return false;
  }
  Prepared pre;
  if (!prepare(call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx, opts, &pre, out_err)) {
    return false;
  }
  if (!fit_holt_winters(pre.series.y, pre.m, &out->fit, out_err)) {
    return false;
  }
  out->timeline = pre.series.timeline;
  out->n = pre.series.y.size();
  return true;
}

// Locates `target_date` on the grid as the fractional `*pos` and its whole
// step `*s`; a horizon past `kMaxHorizonSteps` is `#NUM!`.
bool locate(const FittedSeries& fs, double target_date, double* pos, std::int64_t* s, bool* interpolated,
            Value* out_err) {
  ErrorCode err = ErrorCode::NA;
  if (!target_position(fs.timeline, target_date, pos, interpolated, &err)) {
    *out_err = Value::error(err);
    return false;
  }
  const std::optional<std::int64_t> whole = whole_count(*pos);
  if (!whole || *pos > kMaxHorizonSteps) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  *s = *whole;
  return true;
}

// Point forecast at one target date.
Value forecast_at(const FittedSeries& fs, double target_date) {
  double pos = 0.0;
  std::int64_t s = 0;
  bool interpolated = false;
  Value err = Value::blank();
  if (!locate(fs, target_date, &pos, &s, &interpolated, &err)) {
    return err;
  }
  // A target at or before the first grid node has no forecast.
  if (s <= -static_cast<std::int64_t>(fs.n)) {
    return Value::error(ErrorCode::Num);
  }
  double f = fs.fit.forecast(s);
  if (interpolated) {
    f = interpolate(s, f, fs.fit.forecast(s + 1), pos);
  }
  return to_finite_value(fs.fit.zero_if_negligible(f));
}

// Confidence-interval half-width at one target date.
Value confint_at(const FittedSeries& fs, double target_date, double confidence) {
  double pos = 0.0;
  std::int64_t s = 0;
  bool interpolated = false;
  Value err = Value::blank();
  if (!locate(fs, target_date, &pos, &s, &interpolated, &err)) {
    return err;
  }
  // The interval is undefined at and before the last observation.
  if (s < 0 || (s == 0 && !interpolated)) {
    return Value::error(ErrorCode::Num);
  }
  double variance = 0.0;
  if (s != 0) {
    variance = fs.fit.variance(s);
    if (interpolated) {
      variance = interpolate(s, variance, fs.fit.variance(s + 1), pos);
    }
  } else {
    variance = interpolate(0, 0.0, fs.fit.variance(1), pos);
  }
  // half-width = z * sqrt(variance), z the normal quantile at (1 + confidence) / 2.
  double p = 1.0 - confidence;
  p = p / 2.0;
  p = -p;
  p = p + 1.0;
  const double z = stats_detail::InverseStandardNormal(p);
  return to_finite_value(z * (variance > 0.0 ? std::sqrt(variance) : 0.0));
}

// Applies `at` to target_date: once for a scalar, per element of an array
// keeping its shape. As in the scalar call, which reads target_date first, a
// target error wins over `head_err`, the failure of the rest of the call.
template <typename At>
Value at_each_target(const Value& target, bool head_ok, const Value& head_err, const At& at, Arena& arena) {
  const auto one = [&](const Value& v) -> Value {
    double t = 0.0;
    Value err = Value::blank();
    if (!finite_number(v, &t, &err, ErrorCode::Num)) {
      return err;
    }
    return head_ok ? at(t) : head_err;
  };
  if (!target.is_array()) {
    return one(target);
  }
  const ArrayValue& in = *target.as_array();
  Value* cells = nullptr;
  ArrayValue* out = allocate_array_value(in.rows, in.cols, arena, cells, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t count = static_cast<std::size_t>(in.rows) * in.cols;
  for (std::size_t i = 0; i < count; ++i) {
    cells[i] = one(in.cells[i]);
  }
  return count == 1U ? cells[0] : Value::array(out);
}

}  // namespace

Value eval_forecast_ets_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 6U) {
    return Value::error(ErrorCode::Value);
  }

  // arg 0: target_date, a scalar or an array lifted per element.
  Value err = Value::blank();
  const Value target = eval_node(call.as_call_arg(0), arena, registry, ctx);
  double target_date = 0.0;
  if (!target.is_array() && !finite_number(target, &target_date, &err, ErrorCode::Num)) {
    return err;
  }

  FittedSeries fs;
  const bool fitted = fit_series(call, arity, 3U, arena, registry, ctx, &fs, &err);
  return at_each_target(target, fitted, err, [&fs](double t) { return forecast_at(fs, t); }, arena);
}

Value eval_forecast_ets_confint_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 7U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();
  const Value target = eval_node(call.as_call_arg(0), arena, registry, ctx);
  double target_date = 0.0;
  if (!target.is_array() && !finite_number(target, &target_date, &err, ErrorCode::Num)) {
    return err;
  }

  double confidence = 0.95;
  bool head_ok = arity <= 3U || read_double_arg(call.as_call_arg(3), arena, registry, ctx, 0.95, &confidence, &err);
  // Mac Excel 365 accepts confidence == 0 (degenerate CI, returns 0 since
  // z = InverseStandardNormal(0.5) = 0). The lower-bound rejection is
  // strict (< 0); upper bound stays inclusive on 1 because z diverges to
  // +infinity there.
  if (head_ok && (!std::isfinite(confidence) || confidence < 0.0 || confidence >= 1.0)) {
    err = Value::error(ErrorCode::Num);
    head_ok = false;
  }

  FittedSeries fs;
  head_ok = head_ok && fit_series(call, arity, 4U, arena, registry, ctx, &fs, &err);
  return at_each_target(
      target, head_ok, err, [&fs, confidence](double t) { return confint_at(fs, t, confidence); }, arena);
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

  Prepared pre;
  if (!prepare(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx, opts, &pre, &err)) {
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

  Prepared pre;
  if (!prepare(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx, opts, &pre, &err)) {
    return err;
  }

  // statistic_type = 8 is the timeline step and needs no fit.
  if (stat_type == 8) {
    return Value::number(pre.series.timeline.step);
  }

  HoltWintersFit fit;
  if (!fit_holt_winters(pre.series.y, pre.m, &fit, &err)) {
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
