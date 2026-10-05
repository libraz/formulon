//
// Holt-Winters additive triple-exponential smoothing fit used by the
// FORECAST.ETS family.

#ifndef FORMULON_EVAL_HOLT_WINTERS_H_
#define FORMULON_EVAL_HOLT_WINTERS_H_

#include <cstdint>
#include <vector>

#include "value.h"

namespace formulon {
namespace eval {
namespace ets {

/// Fitted smoothing parameters, final state and in-sample diagnostics.

struct HoltWintersFit {
  // Smoothing parameters fitted by Nelder-Mead. `gamma` is exactly 0.0
  // when the model is non-seasonal (m == 1).
  double alpha = 0.0;
  double beta = 0.0;
  double gamma = 0.0;
  // Final state at the end of the in-sample series.
  double level = 0.0;
  double trend = 0.0;
  std::vector<double> season;  // length m (or empty when m == 1)
  std::uint32_t m = 1U;
  // In-sample fit diagnostics. `residuals` has the same length as the
  // resampled series; the first m entries (or the first 1 when m == 1)
  // are zero because no one-step-ahead forecast is defined there.
  std::vector<double> residuals;
  double mae = 0.0;
  double rmse = 0.0;
  double mase = 0.0;
  double smape = 0.0;
};

/// Fits a Holt-Winters model to `y` with seasonality `m`. On success fills
/// `*out` and returns true; writes `*out_err` and returns false on degenerate
/// inputs (n < 2, or seasonal m with n < 2*m).
bool fit_holt_winters(const std::vector<double>& y, std::uint32_t m, HoltWintersFit* out, Value* out_err);

}  // namespace ets
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_HOLT_WINTERS_H_
