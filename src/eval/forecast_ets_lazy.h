// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lazy impls for the FORECAST.ETS family: FORECAST.ETS,
// FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY, and FORECAST.ETS.STAT.
//
// All four functions share a Holt-Winters additive triple-exponential
// smoothing core. The core fits three smoothing parameters
// `(alpha, beta, gamma) in (0, 1)^3` to in-sample one-step-ahead SSE via
// bounded Nelder-Mead, and produces:
//   * a point forecast `F_{t+h} = L_t + h*B_t + S_{t-m+1+(h-1) mod m}`,
//   * the in-sample fit residuals (RMSE / MAE / MASE / SMAPE),
//   * an auto-detected seasonality length (ACF on lags 2..min(n/2, 52)),
//   * a normal-approximation confidence half-width
//     `hw = z * RMSE * sqrt(h)`.
//
// They ride the lazy-dispatch seam rather than the eager path because
// the `values` and `timeline` arguments may be `Ref` / `RangeOp` /
// `ArrayLiteral` shapes that must reach the impl with their `(rows,
// cols)` rectangle intact: timeline preprocessing pairs the two arrays
// element-wise and rejects shape mismatches with `#N/A`. See
// `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// `eval_node` recursion entry point.
//
// The implementation lives in a single TU (`forecast_ets_lazy.cpp`)
// because the four impls share a substantial helper surface
// (preprocessing, fit, ACF) that we deliberately keep in-TU.

#ifndef FORMULON_EVAL_FORECAST_ETS_LAZY_H_
#define FORMULON_EVAL_FORECAST_ETS_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `FORECAST.ETS(target_date, values, timeline, [seasonality=1],
///   [data_completion=1], [aggregation=1])` -- point forecast.
///
/// Returns a Number on success. `target_date` must lie at or after the
/// first timeline entry; `target_date < t_0` -> `#NUM!`. `values` and
/// `timeline` must have equal length (`#N/A` otherwise). A single-point
/// series surfaces `#DIV/0!` (matches Mac Excel 365). With a seasonal
/// model, fewer than `2 * m` post-preprocessing points surfaces `#N/A`.
/// Errors in either array propagate left-to-right (timeline cells
/// scanned first).
Value eval_forecast_ets_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx);

/// `FORECAST.ETS.CONFINT(target_date, values, timeline,
///   [confidence=0.95], [seasonality=1], [data_completion=1],
///   [aggregation=1])` -- one-sided confidence-interval half-width.
///
/// Returns `z * RMSE * sqrt(h)` where `z =
/// stats_detail::InverseStandardNormal((1 + confidence) / 2)`, RMSE is
/// the in-sample residual standard error from the fit, and `h` is the
/// integer step count between the last training timeline entry and
/// `target_date`. `confidence` outside `[0, 1)` -> `#NUM!` (0 is
/// degenerate but accepted, returning 0). `target_date` inside or before
/// the training window (`h < 1`) -> `#NUM!`. All other errors match
/// `FORECAST.ETS`.
Value eval_forecast_ets_confint_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                     const EvalContext& ctx);

/// `FORECAST.ETS.SEASONALITY(values, timeline, [data_completion=1],
///   [aggregation=1])` -- detected seasonality length.
///
/// Auto-detects the dominant period using the autocorrelation function
/// over lags `2..min(n/2, 52)`. Returns the lag with the largest
/// `|ACF|` if its magnitude is at least 0.3; otherwise returns 0
/// (Mac Excel 365 convention for "no period detected"). Series too
/// short to detect a period (`n < 4`) returns 0. Other errors match
/// `FORECAST.ETS`.
Value eval_forecast_ets_seasonality_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                         const EvalContext& ctx);

/// `FORECAST.ETS.STAT(values, timeline, statistic_type, [seasonality=1],
///   [data_completion=1], [aggregation=1])` -- diagnostic statistic.
///
/// `statistic_type` selects one of:
///   1 -> alpha (level smoothing parameter)
///   2 -> beta  (trend smoothing parameter)
///   3 -> gamma (seasonal smoothing parameter; 0 if non-seasonal)
///   4 -> MASE  (mean absolute scaled error)
///   5 -> SMAPE (symmetric MAPE)
///   6 -> MAE   (mean absolute error of in-sample fit)
///   7 -> RMSE  (root mean squared error of in-sample fit)
///   8 -> step_size (median delta-t of the original timeline before
///                   aggregation)
/// Out-of-range / non-integer (after truncation outside 1..8) -> `#NUM!`.
/// Other errors match `FORECAST.ETS`.
Value eval_forecast_ets_stat_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                  const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FORECAST_ETS_LAZY_H_
