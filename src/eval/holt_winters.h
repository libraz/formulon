//
// Additive Holt-Winters fit used by the FORECAST.ETS family: smoothing
// parameters are chosen from a fixed candidate table by a holdout-aware SSE.

#ifndef FORMULON_EVAL_HOLT_WINTERS_H_
#define FORMULON_EVAL_HOLT_WINTERS_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "value.h"

namespace formulon {
namespace eval {
namespace ets {

/// Fitted smoothing parameters, final state and fit diagnostics.
struct HoltWintersFit {
  double alpha = 0.0;
  double beta = 0.0;
  /// Equals DBL_EPSILON when the model is non-seasonal.
  double gamma = 0.0;
  double mae = 0.0;
  double rmse = 0.0;
  double mase = 0.0;
  double smape = 0.0;
  /// Mean squared one-step error over the whole series.
  double sigma2 = 0.0;
  /// Trend of the initial state.
  double initial_trend = 0.0;
  /// Season length used by the recurrences (at least 2).
  std::size_t period = 2U;
  /// [level, trend, seasonal slots...]; the seasonal slots are a ring.
  std::vector<double> state;

  /// Point forecast `h` steps after the last observation (`h` may be <= 0).
  /// Callers keep `|h|` at most 2^53 (`ets::kMaxWholeCount`).
  double forecast(std::int64_t h) const;

  /// Returns 0 for a forecast that is negligible next to the initial trend.
  double zero_if_negligible(double r) const;

  /// Forecast error variance at horizon `1 <= h <= 2^53`; the bound keeps
  /// `2 * h - 1` inside the int64 range.
  double variance(std::int64_t h) const;
};

/// Fits to `y` with season length `m` (1 means non-seasonal). On success fills
/// `*out` and returns true; writes `*out_err` and returns false when `y` has
/// fewer than two points.
bool fit_holt_winters(const std::vector<double>& y, std::uint32_t m, HoltWintersFit* out, Value* out_err);

}  // namespace ets
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_HOLT_WINTERS_H_
