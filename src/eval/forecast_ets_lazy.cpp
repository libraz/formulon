//
// Implementation of the FORECAST.ETS family lazy impls.
//
// All four functions share a single Holt-Winters additive triple-exponential
// smoothing core implemented in `HoltWintersFit`. The pipeline per call is:
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

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <utility>
#include <vector>

#include "eval/builtins/stats/stats_helpers.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
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

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

// Step-size deviation tolerance for the resample grid. A consecutive delta-t
// is required to fall within +/-`kStepDeviation` of the median step
// (Excel's documented +/-30% tolerance). ORACLE-PENDING.
constexpr double kStepDeviation = 0.3;

// Cap on autocorrelation lags inspected by the seasonality detector. Excel
// allows a maximum seasonality of 8760 (hours in a year), but for the ACF
// scan the detection ceiling is 52 (weeks per year) which matches the
// typical "weekly seasonality on weekly data" calibration. ORACLE-PENDING.
constexpr std::uint32_t kMaxSeasonalityDetectLag = 52U;

// Maximum manually-specified seasonality. Above this the third argument
// is rejected with `#NUM!` (matches Excel's documented 8760 limit).
constexpr std::uint32_t kMaxSeasonalityArg = 8760U;

// Minimum |ACF| magnitude required to declare a detected seasonality.
// Below this the auto-detector returns m = 1. ORACLE-PENDING.
constexpr double kAcfDetectionThreshold = 0.3;

// Relative floor on the linear-trend residual below which the detrended
// series is treated as pure rounding noise and no period is reported. The
// comparison is against the raw mean-centred variance, so it scales with
// the data instead of assuming a magnitude.
constexpr double kDetrendedResidualEpsilon = 1e-20;

// Nelder-Mead bounds clamp -- coordinates are squeezed into
// (kBoundEps, 1 - kBoundEps) so the gradient never escapes the unit
// cube and the smoothing recurrences stay well-defined.
constexpr double kBoundEps = 1e-6;

// Nelder-Mead convergence: stop when (fmax - fmin) over the simplex falls
// below this, or when iteration count exceeds kMaxNelderMeadIters.
constexpr double kNelderMeadFTol = 1e-9;
constexpr int kMaxNelderMeadIters = 200;

// ---------------------------------------------------------------------------
// Aggregation modes (Excel's documented enumeration)
// ---------------------------------------------------------------------------

enum class AggregationMode : std::uint32_t {
  kAverage = 1U,
  kCount = 2U,
  kCountA = 3U,
  kMax = 4U,
  kMedian = 5U,
  kMin = 6U,
  kSum = 7U,
};

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
// (t, y) pair container shared across the helpers
// ---------------------------------------------------------------------------

struct Series {
  std::vector<double> t;
  std::vector<double> y;
};

// ---------------------------------------------------------------------------
// Aggregation
// ---------------------------------------------------------------------------

// Aggregates a run `[begin, end)` of identical-timeline points using the
// requested mode. The COUNT / COUNTA modes return the run length (COUNT
// excludes blanks but blanks have already been rejected at the strict-
// numeric step, so the two modes coincide here -- both return the count).
double aggregate_run(const Series& src, std::size_t begin, std::size_t end, AggregationMode mode) {
  const std::size_t n = end - begin;
  if (n == 0U) {
    return 0.0;
  }
  if (n == 1U) {
    // Optimisation: a singleton run with any mode collapses to its
    // single sample (COUNT/COUNTA collapse to 1, but a singleton is
    // already its own answer for 5 of the 7 modes; we special-case
    // COUNT/COUNTA below).
    if (mode == AggregationMode::kCount || mode == AggregationMode::kCountA) {
      return 1.0;
    }
    return src.y[begin];
  }
  switch (mode) {
    case AggregationMode::kCount:
    case AggregationMode::kCountA:
      return static_cast<double>(n);
    case AggregationMode::kSum: {
      double s = 0.0;
      for (std::size_t i = begin; i < end; ++i)
        s += src.y[i];
      return s;
    }
    case AggregationMode::kMax: {
      double m = src.y[begin];
      for (std::size_t i = begin + 1U; i < end; ++i) {
        if (src.y[i] > m)
          m = src.y[i];
      }
      return m;
    }
    case AggregationMode::kMin: {
      double m = src.y[begin];
      for (std::size_t i = begin + 1U; i < end; ++i) {
        if (src.y[i] < m)
          m = src.y[i];
      }
      return m;
    }
    case AggregationMode::kMedian: {
      std::vector<double> tmp(src.y.begin() + static_cast<std::ptrdiff_t>(begin),
                              src.y.begin() + static_cast<std::ptrdiff_t>(end));
      std::sort(tmp.begin(), tmp.end());
      const std::size_t mid = tmp.size() / 2U;
      if ((tmp.size() & 1U) == 1U) {
        return tmp[mid];
      }
      return 0.5 * (tmp[mid - 1U] + tmp[mid]);
    }
    case AggregationMode::kAverage:
    default: {
      double s = 0.0;
      for (std::size_t i = begin; i < end; ++i)
        s += src.y[i];
      return s / static_cast<double>(n);
    }
  }
}

// Sorts `src` by t (stable) and collapses runs of identical t via `mode`.
// Returns the new (deduped, sorted) series.
Series sort_and_aggregate(const Series& src, AggregationMode mode) {
  const std::size_t n = src.t.size();
  std::vector<std::size_t> idx(n);
  for (std::size_t i = 0; i < n; ++i)
    idx[i] = i;
  std::stable_sort(idx.begin(), idx.end(),
                   [&src](std::size_t a, std::size_t b) noexcept { return src.t[a] < src.t[b]; });
  Series sorted;
  sorted.t.reserve(n);
  sorted.y.reserve(n);
  for (std::size_t i = 0; i < n; ++i) {
    sorted.t.push_back(src.t[idx[i]]);
    sorted.y.push_back(src.y[idx[i]]);
  }
  // Aggregate consecutive runs with identical t.
  Series out;
  out.t.reserve(sorted.t.size());
  out.y.reserve(sorted.y.size());
  std::size_t i = 0;
  while (i < sorted.t.size()) {
    std::size_t j = i + 1U;
    while (j < sorted.t.size() && sorted.t[j] == sorted.t[i]) {
      ++j;
    }
    out.t.push_back(sorted.t[i]);
    out.y.push_back(aggregate_run(sorted, i, j, mode));
    i = j;
  }
  return out;
}

// ---------------------------------------------------------------------------
// Step detection and resampling
// ---------------------------------------------------------------------------

// Computes the median of consecutive deltas in a sorted timeline. Caller
// guarantees `series.t.size() >= 2`.
double median_delta(const std::vector<double>& t) {
  std::vector<double> deltas;
  deltas.reserve(t.size() - 1U);
  for (std::size_t i = 1; i < t.size(); ++i) {
    deltas.push_back(t[i] - t[i - 1U]);
  }
  std::sort(deltas.begin(), deltas.end());
  const std::size_t mid = deltas.size() / 2U;
  if ((deltas.size() & 1U) == 1U) {
    return deltas[mid];
  }
  return 0.5 * (deltas[mid - 1U] + deltas[mid]);
}

// Resamples a sorted, deduplicated `series` onto an evenly spaced grid
// with step `step`. The first point is taken as the grid origin; for
// each later grid index `k = round((t_i - t_0) / step)` the value is
// the corresponding y. Missing slots are then filled per
// `data_completion`: 0 -> zero-fill, 1 -> linear interpolation between
// the nearest known neighbours.
//
// On step / regularity failures writes `#NUM!` into `*out_err` and
// returns `false`; otherwise writes the resampled series into `*out`.
bool resample_to_grid(const Series& sorted, double step, int data_completion, Series* out, Value* out_err) {
  if (step <= 0.0 || !std::isfinite(step)) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  const double t0 = sorted.t.front();
  const double t_last = sorted.t.back();
  // Compute the integer index of the last point. A negative or non-finite
  // span is impossible after sorting, but the round-then-cast guard below
  // copes with floating-point jitter.
  const double span = t_last - t0;
  const double last_idx_d = std::round(span / step);
  if (!std::isfinite(last_idx_d) || last_idx_d < 0.0) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  const std::size_t grid_size = static_cast<std::size_t>(last_idx_d) + 1U;
  // Defensive cap: a grid size > 2^20 would mean the timeline spans >1M
  // steps, which is well beyond Excel's practical limits and almost
  // certainly indicates a step-detection failure.
  if (grid_size > (1U << 20U)) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  std::vector<double> grid_y(grid_size, 0.0);
  std::vector<bool> filled(grid_size, false);
  for (std::size_t i = 0; i < sorted.t.size(); ++i) {
    const double rel = (sorted.t[i] - t0) / step;
    const double rounded = std::round(rel);
    // Regularity check: every input must land within +/-30% of an
    // integer grid index when measured in step units.
    const double resid = std::fabs(rel - rounded);
    if (resid > kStepDeviation) {
      *out_err = Value::error(ErrorCode::Num);
      return false;
    }
    const std::size_t k = static_cast<std::size_t>(rounded);
    if (k >= grid_size) {
      *out_err = Value::error(ErrorCode::Num);
      return false;
    }
    if (filled[k]) {
      // Two distinct timeline values mapping to the same grid index
      // means the +/-30% tolerance was insufficient; treat as a
      // step-detection failure.
      *out_err = Value::error(ErrorCode::Num);
      return false;
    }
    grid_y[k] = sorted.y[i];
    filled[k] = true;
  }

  if (data_completion == 0) {
    // Zero-fill: missing slots already initialised to 0.0; nothing
    // further to do.
  } else {
    // Linear interpolation between known neighbours. The endpoints are
    // always filled (the first sample lives at index 0; the last lives
    // at index grid_size - 1) by construction.
    std::size_t i = 0;
    while (i < grid_size) {
      if (filled[i]) {
        ++i;
        continue;
      }
      // Find the previous filled index (must exist; first cell is
      // filled by construction).
      std::size_t prev = i;
      while (prev > 0U && !filled[prev])
        --prev;
      // Find the next filled index (must exist; last cell is filled).
      std::size_t next = i;
      while (next < grid_size && !filled[next])
        ++next;
      if (next >= grid_size) {
        // No right anchor -- defensive; shouldn't happen since the last
        // cell is always filled. Treat as a regularity failure.
        *out_err = Value::error(ErrorCode::Num);
        return false;
      }
      const double y_prev = grid_y[prev];
      const double y_next = grid_y[next];
      for (std::size_t k = prev + 1U; k < next; ++k) {
        const double frac = static_cast<double>(k - prev) / static_cast<double>(next - prev);
        grid_y[k] = y_prev + frac * (y_next - y_prev);
        filled[k] = true;
      }
      i = next;
    }
  }

  out->t.resize(grid_size);
  out->y.resize(grid_size);
  for (std::size_t i = 0; i < grid_size; ++i) {
    out->t[i] = t0 + static_cast<double>(i) * step;
    out->y[i] = grid_y[i];
  }
  return true;
}

// ---------------------------------------------------------------------------
// Seasonality detection (autocorrelation)
// ---------------------------------------------------------------------------

// Returns the auto-detected seasonality length (>= 1). 1 means "no
// seasonality" -- both for series too short to scan and for series
// where no lag clears the threshold.
std::uint32_t detect_seasonality(const std::vector<double>& y) {
  const std::size_t n = y.size();
  if (n < 4U) {
    return 1U;
  }
  // The autocorrelation scan runs on the linear-trend residual, not on the
  // raw series. A trending series is non-stationary: every lag of a plain
  // ramp autocorrelates near 1, so an un-detrended scan reports a period
  // for data that has none. Ordinary least squares against the sample
  // index is enough here because the resample grid is equally spaced by
  // construction, which lets the regression sums close in one pass.
  const double count = static_cast<double>(n);
  double sum_y = 0.0;
  double sum_iy = 0.0;
  for (std::size_t i = 0; i < n; ++i) {
    sum_y += y[i];
    sum_iy += static_cast<double>(i) * y[i];
  }
  const double mean_y = sum_y / count;
  const double mean_i = (count - 1.0) / 2.0;
  // sum (i - mean_i)^2 over 0..n-1 in closed form.
  const double sxx = count * (count * count - 1.0) / 12.0;
  const double sxy = sum_iy - mean_i * sum_y;
  const double slope = (sxx > 0.0) ? (sxy / sxx) : 0.0;

  std::vector<double> resid(n);
  double var = 0.0;
  double raw_var = 0.0;
  for (std::size_t i = 0; i < n; ++i) {
    const double fitted = mean_y + slope * (static_cast<double>(i) - mean_i);
    resid[i] = y[i] - fitted;
    var += resid[i] * resid[i];
    const double centered = y[i] - mean_y;
    raw_var += centered * centered;
  }
  // A residual that is pure rounding noise carries no seasonal signal, so
  // compare it against the scale of the original series rather than to an
  // exact zero: a perfect ramp leaves residuals around 1e-15, whose
  // autocorrelations are otherwise arbitrary and clear the threshold.
  if (raw_var == 0.0 || var <= raw_var * kDetrendedResidualEpsilon) {
    return 1U;
  }
  const std::uint32_t max_lag_u = std::min<std::uint32_t>(static_cast<std::uint32_t>(n / 2U), kMaxSeasonalityDetectLag);
  if (max_lag_u < 2U) {
    return 1U;
  }
  // The residual is mean-zero by construction, so the centring term drops
  // out of the autocorrelation numerator.
  //
  // Only positive autocorrelation counts. A seasonality is a repetition,
  // and a negative correlation at lag k says the residual flips sign every
  // k steps — the repetition is at 2k, not k. Ranking by magnitude would
  // let that anti-correlated lag outrank the real period.
  double best_acf = 0.0;
  std::uint32_t best_lag = 1U;
  for (std::uint32_t k = 2U; k <= max_lag_u; ++k) {
    double num = 0.0;
    for (std::size_t i = k; i < n; ++i) {
      num += resid[i] * resid[i - k];
    }
    const double r = num / var;
    if (r > best_acf) {
      best_acf = r;
      best_lag = k;
    }
  }
  if (best_acf >= kAcfDetectionThreshold) {
    return best_lag;
  }
  return 1U;
}

// ---------------------------------------------------------------------------
// Holt-Winters fit
// ---------------------------------------------------------------------------

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

// Runs the Holt-Winters smoothing recurrences over `y` with explicit
// initial state (`init_level`, `init_trend`, `init_season[0..m-1]`),
// populates `*out` (level / trend / season / residuals), and returns the
// in-sample SSE used by the optimiser.
//
// Recurrences (additive seasonality):
//   L_t = alpha (y_t - S_{t-m}) + (1-alpha)(L_{t-1} + B_{t-1})
//   B_t = beta  (L_t - L_{t-1}) + (1-beta)  B_{t-1}
//   S_t = gamma (y_t - L_t)     + (1-gamma) S_{t-m}
// Forecast residual at step t is `y_t - (L_{t-1} + B_{t-1} + S_{t-m})`.
//
// `init_season` may be null when `m == 1`.
double simulate_with_init(const std::vector<double>& y, std::uint32_t m, double alpha, double beta, double gamma,
                          double init_level, double init_trend, const double* init_season, HoltWintersFit* out) {
  const std::size_t n = y.size();
  double level = init_level;
  double trend = init_trend;
  std::vector<double> season(m == 1U ? 0U : static_cast<std::size_t>(m), 0.0);
  if (m != 1U && init_season != nullptr) {
    for (std::uint32_t i = 0; i < m; ++i) {
      season[i] = init_season[i];
    }
  }

  std::vector<double> residuals(n, 0.0);
  double sse = 0.0;
  // The first `start` indices have no one-step-ahead forecast (no prior
  // level / trend / seasonal estimate is available beneath them).
  const std::size_t start = (m == 1U) ? 1U : static_cast<std::size_t>(m);
  for (std::size_t t = start; t < n; ++t) {
    const double prev_level = level;
    const double prev_trend = trend;
    const double s_prev = (m == 1U) ? 0.0 : season[t % m];
    const double y_hat = prev_level + prev_trend + s_prev;
    const double err = y[t] - y_hat;
    residuals[t] = err;
    sse += err * err;
    // Update equations.
    level = alpha * (y[t] - s_prev) + (1.0 - alpha) * (prev_level + prev_trend);
    trend = beta * (level - prev_level) + (1.0 - beta) * prev_trend;
    if (m != 1U) {
      season[t % m] = gamma * (y[t] - level) + (1.0 - gamma) * s_prev;
    }
  }

  if (out != nullptr) {
    out->alpha = alpha;
    out->beta = beta;
    out->gamma = (m == 1U) ? 0.0 : gamma;
    out->level = level;
    out->trend = trend;
    out->season = season;
    out->m = m;
    out->residuals = std::move(residuals);
  }
  return sse;
}

// Computes the textbook two-season-means default initial state used as the
// starting point of the optimiser. With `m == 1` we use Holt's double
// exponential: L_0 = y_0, B_0 = y_1 - y_0 (no seasonal vector).
void compute_default_init(const std::vector<double>& y, std::uint32_t m, double* out_level, double* out_trend,
                          std::vector<double>* out_season) {
  const std::size_t n = y.size();
  out_season->assign(m == 1U ? 0U : static_cast<std::size_t>(m), 0.0);
  if (m == 1U) {
    *out_level = y[0];
    *out_trend = (n >= 2U) ? (y[1] - y[0]) : 0.0;
    return;
  }
  double mean1 = 0.0;
  for (std::uint32_t i = 0; i < m; ++i)
    mean1 += y[i];
  mean1 /= static_cast<double>(m);
  double mean2 = mean1;
  if (n >= 2U * static_cast<std::size_t>(m)) {
    mean2 = 0.0;
    for (std::uint32_t i = 0; i < m; ++i)
      mean2 += y[m + i];
    mean2 /= static_cast<double>(m);
  }
  *out_level = mean1;
  *out_trend = (mean2 - mean1) / static_cast<double>(m);
  double s_sum = 0.0;
  for (std::uint32_t i = 0; i < m; ++i) {
    (*out_season)[i] = y[i] - mean1;
    s_sum += (*out_season)[i];
  }
  const double s_mean = s_sum / static_cast<double>(m);
  for (std::uint32_t i = 0; i < m; ++i)
    (*out_season)[i] -= s_mean;
}

// Backwards-compatible wrapper: runs `simulate_with_init` from the textbook
// two-season-means default initialisation.
double simulate(const std::vector<double>& y, std::uint32_t m, double alpha, double beta, double gamma,
                HoltWintersFit* out) {
  double init_level = 0.0;
  double init_trend = 0.0;
  std::vector<double> init_season;
  compute_default_init(y, m, &init_level, &init_trend, &init_season);
  return simulate_with_init(y, m, alpha, beta, gamma, init_level, init_trend,
                            init_season.empty() ? nullptr : init_season.data(), out);
}

// ---------------------------------------------------------------------------
// Bounded Nelder-Mead in 2D / 3D
// ---------------------------------------------------------------------------

// Clamps each coordinate of `v` (length `dim`) into (kBoundEps, 1 - kBoundEps).
void clamp_vertex(double* v, std::size_t dim) noexcept {
  for (std::size_t i = 0; i < dim; ++i) {
    if (v[i] < kBoundEps)
      v[i] = kBoundEps;
    if (v[i] > 1.0 - kBoundEps)
      v[i] = 1.0 - kBoundEps;
  }
}

// Expands a packed parameter vector of length `dim` into (alpha, beta, gamma).
// `dim == 2` => non-seasonal: gamma is forced to 0.
// `dim == 3` => seasonal: all three are read out.
void unpack_params(const double* p, std::size_t dim, double* alpha, double* beta, double* gamma) noexcept {
  *alpha = p[0];
  *beta = p[1];
  *gamma = (dim == 3U) ? p[2] : 0.0;
}

// SSE objective used by the optimiser. Returns +inf for non-finite SSE
// (which guides the simplex away from pathological corners).
double objective(const double* params, std::size_t dim, const std::vector<double>& y, std::uint32_t m) {
  double alpha = 0.0;
  double beta = 0.0;
  double gamma = 0.0;
  unpack_params(params, dim, &alpha, &beta, &gamma);
  const double sse = simulate(y, m, alpha, beta, gamma, nullptr);
  if (!std::isfinite(sse)) {
    return std::numeric_limits<double>::infinity();
  }
  return sse;
}

// Runs bounded Nelder-Mead over the unit cube of dimension `dim` (2 or 3)
// minimising `objective`. Writes the best (alpha, beta, gamma) into the
// out-pointers and returns the best SSE.
//
// The simplex coefficients are the canonical (Nelder-Mead 1965) values:
//   reflect rho = 1, expand chi = 2, contract psi = 0.5, shrink sigma = 0.5.
double nelder_mead(const std::vector<double>& y, std::uint32_t m, std::size_t dim, double* out_alpha, double* out_beta,
                   double* out_gamma) {
  // Initial simplex: (0.2, 0.1, 0.1) plus unit-perturbations of magnitude
  // 0.1 in each axis. Comfortably inside the bounds.
  const std::size_t k = dim + 1U;
  std::vector<std::vector<double>> simplex(k, std::vector<double>(dim, 0.0));
  // Vertex 0: anchor.
  simplex[0][0] = 0.2;
  simplex[0][1] = 0.1;
  if (dim == 3U)
    simplex[0][2] = 0.1;
  // Vertices 1..dim: perturb the i-th axis.
  for (std::size_t i = 1; i < k; ++i) {
    for (std::size_t j = 0; j < dim; ++j)
      simplex[i][j] = simplex[0][j];
    simplex[i][i - 1U] += 0.1;
  }
  for (auto& v : simplex)
    clamp_vertex(v.data(), dim);

  std::vector<double> fvals(k, 0.0);
  for (std::size_t i = 0; i < k; ++i) {
    fvals[i] = objective(simplex[i].data(), dim, y, m);
  }

  for (int iter = 0; iter < kMaxNelderMeadIters; ++iter) {
    // Sort vertices by function value (ascending).
    std::vector<std::size_t> order(k);
    for (std::size_t i = 0; i < k; ++i)
      order[i] = i;
    std::sort(order.begin(), order.end(),
              [&fvals](std::size_t a, std::size_t b) noexcept { return fvals[a] < fvals[b]; });
    const std::size_t best = order[0];
    const std::size_t worst = order[k - 1U];
    const std::size_t second_worst = order[k - 2U];

    // Convergence check: span of f-values across the simplex.
    if (std::isfinite(fvals[best]) && std::isfinite(fvals[worst]) && (fvals[worst] - fvals[best]) < kNelderMeadFTol) {
      break;
    }

    // Centroid of the best-`dim` vertices (i.e. all except the worst).
    std::vector<double> centroid(dim, 0.0);
    for (std::size_t i = 0; i < k; ++i) {
      if (i == worst)
        continue;
      for (std::size_t j = 0; j < dim; ++j)
        centroid[j] += simplex[i][j];
    }
    for (std::size_t j = 0; j < dim; ++j)
      centroid[j] /= static_cast<double>(dim);

    // Reflect.
    std::vector<double> reflected(dim);
    for (std::size_t j = 0; j < dim; ++j) {
      reflected[j] = centroid[j] + 1.0 * (centroid[j] - simplex[worst][j]);
    }
    clamp_vertex(reflected.data(), dim);
    const double f_reflected = objective(reflected.data(), dim, y, m);

    if (f_reflected < fvals[second_worst] && f_reflected >= fvals[best]) {
      simplex[worst] = reflected;
      fvals[worst] = f_reflected;
      continue;
    }

    if (f_reflected < fvals[best]) {
      // Expand.
      std::vector<double> expanded(dim);
      for (std::size_t j = 0; j < dim; ++j) {
        expanded[j] = centroid[j] + 2.0 * (reflected[j] - centroid[j]);
      }
      clamp_vertex(expanded.data(), dim);
      const double f_expanded = objective(expanded.data(), dim, y, m);
      if (f_expanded < f_reflected) {
        simplex[worst] = expanded;
        fvals[worst] = f_expanded;
      } else {
        simplex[worst] = reflected;
        fvals[worst] = f_reflected;
      }
      continue;
    }

    // Contract.
    std::vector<double> contracted(dim);
    if (f_reflected < fvals[worst]) {
      // Outside contraction.
      for (std::size_t j = 0; j < dim; ++j) {
        contracted[j] = centroid[j] + 0.5 * (reflected[j] - centroid[j]);
      }
    } else {
      // Inside contraction.
      for (std::size_t j = 0; j < dim; ++j) {
        contracted[j] = centroid[j] + 0.5 * (simplex[worst][j] - centroid[j]);
      }
    }
    clamp_vertex(contracted.data(), dim);
    const double f_contracted = objective(contracted.data(), dim, y, m);
    if (f_contracted < std::min(f_reflected, fvals[worst])) {
      simplex[worst] = contracted;
      fvals[worst] = f_contracted;
      continue;
    }

    // Shrink toward best.
    for (std::size_t i = 0; i < k; ++i) {
      if (i == best)
        continue;
      for (std::size_t j = 0; j < dim; ++j) {
        simplex[i][j] = simplex[best][j] + 0.5 * (simplex[i][j] - simplex[best][j]);
      }
      clamp_vertex(simplex[i].data(), dim);
      fvals[i] = objective(simplex[i].data(), dim, y, m);
    }
  }

  // Return the best vertex.
  std::size_t best_idx = 0;
  for (std::size_t i = 1; i < k; ++i) {
    if (fvals[i] < fvals[best_idx])
      best_idx = i;
  }
  unpack_params(simplex[best_idx].data(), dim, out_alpha, out_beta, out_gamma);
  return fvals[best_idx];
}

// ---------------------------------------------------------------------------
// Joint (alpha, beta, gamma, L_0, B_0, S_0..S_{m-1}) optimisation
// ---------------------------------------------------------------------------
//
// The 3D `nelder_mead` above only tunes the smoothing parameters at a fixed
// two-season-means initial state. On strongly cyclic data Mac Excel reaches
// a much tighter fit by jointly tuning the initial state vector along with
// the smoothing parameters; the joint optimum can sit in a basin our 3D
// optimiser cannot reach. The generalised optimiser below packs:
//
//   p = [alpha, beta[, gamma], L_0, B_0, S_0, ..., S_{m-1}]
//
// where the seasonal block is omitted when `m == 1`. The smoothing block is
// clamped to (kBoundEps, 1 - kBoundEps); the initial-state block is bounded
// to a generous data-range envelope so the simplex can move freely without
// running away to non-finite SSE.

struct ParamBounds {
  double lo;
  double hi;
};

// Computes the per-coordinate bounds used by `nelder_mead_full`. The
// smoothing-param block (the first `smooth_dim` coordinates) is `(eps,
// 1 - eps)`; the initial-state block (the remaining `dim - smooth_dim`
// coordinates) is `[y_min - 5 * range, y_max + 5 * range]` so the simplex
// has room to escape its starting basin.
std::vector<ParamBounds> make_full_bounds(const std::vector<double>& y, std::size_t dim, std::size_t smooth_dim) {
  std::vector<ParamBounds> bounds(dim);
  for (std::size_t i = 0; i < smooth_dim; ++i) {
    bounds[i] = ParamBounds{kBoundEps, 1.0 - kBoundEps};
  }
  double y_min = y[0];
  double y_max = y[0];
  for (double v : y) {
    if (v < y_min)
      y_min = v;
    if (v > y_max)
      y_max = v;
  }
  const double range = std::max(y_max - y_min, 1.0);
  const double lo = y_min - 5.0 * range;
  const double hi = y_max + 5.0 * range;
  for (std::size_t i = smooth_dim; i < dim; ++i) {
    bounds[i] = ParamBounds{lo, hi};
  }
  return bounds;
}

void clamp_vertex_bounded(double* v, std::size_t dim, const ParamBounds* bounds) noexcept {
  for (std::size_t i = 0; i < dim; ++i) {
    if (v[i] < bounds[i].lo)
      v[i] = bounds[i].lo;
    if (v[i] > bounds[i].hi)
      v[i] = bounds[i].hi;
  }
}

// SSE objective for the joint optimiser. Reads packed params:
//   smooth_dim == 2 (m == 1): [alpha, beta, L_0, B_0]
//   smooth_dim == 3 (m >= 2): [alpha, beta, gamma, L_0, B_0, S_0..S_{m-1}]
double objective_full(const double* p, std::size_t /*dim*/, std::size_t smooth_dim, const std::vector<double>& y,
                      std::uint32_t m) {
  const double alpha = p[0];
  const double beta = p[1];
  const double gamma = (smooth_dim == 3U) ? p[2] : 0.0;
  const double init_level = p[smooth_dim];
  const double init_trend = p[smooth_dim + 1U];
  const double* init_season = (m >= 2U) ? &p[smooth_dim + 2U] : nullptr;
  const double sse = simulate_with_init(y, m, alpha, beta, gamma, init_level, init_trend, init_season, nullptr);
  return std::isfinite(sse) ? sse : std::numeric_limits<double>::infinity();
}

// Generalised bounded Nelder-Mead with per-coordinate bounds and an
// arbitrary-dimension starting simplex. Identical search logic to the 3D
// `nelder_mead`; the only differences are that the simplex is built around
// the supplied `seed` (length `dim`) and that `clamp_vertex_bounded` uses
// the per-coordinate bounds. Iteration cap and ftol are scaled with `dim`
// so higher-dim landscapes get proportionally more budget.
double nelder_mead_full(const std::vector<double>& y, std::uint32_t m, std::size_t dim, std::size_t smooth_dim,
                        const double* seed, double seed_step, const std::vector<ParamBounds>& bounds,
                        std::vector<double>* out_params) {
  const std::size_t k = dim + 1U;
  std::vector<std::vector<double>> simplex(k, std::vector<double>(dim, 0.0));
  for (std::size_t j = 0; j < dim; ++j) {
    simplex[0][j] = seed[j];
  }
  for (std::size_t i = 1; i < k; ++i) {
    for (std::size_t j = 0; j < dim; ++j) {
      simplex[i][j] = simplex[0][j];
    }
    simplex[i][i - 1U] += seed_step;
  }
  for (auto& v : simplex) {
    clamp_vertex_bounded(v.data(), dim, bounds.data());
  }

  std::vector<double> fvals(k, 0.0);
  for (std::size_t i = 0; i < k; ++i) {
    fvals[i] = objective_full(simplex[i].data(), dim, smooth_dim, y, m);
  }

  // Iteration budget grows with dim — a 9D landscape needs ~2-3x more
  // iterations than the 3D one to converge through the same number of
  // simplex collapses.
  const int iter_cap = static_cast<int>(static_cast<std::size_t>(kMaxNelderMeadIters) * std::max<std::size_t>(1U, dim));

  for (int iter = 0; iter < iter_cap; ++iter) {
    std::vector<std::size_t> order(k);
    for (std::size_t i = 0; i < k; ++i)
      order[i] = i;
    std::sort(order.begin(), order.end(),
              [&fvals](std::size_t a, std::size_t b) noexcept { return fvals[a] < fvals[b]; });
    const std::size_t best = order[0];
    const std::size_t worst = order[k - 1U];
    const std::size_t second_worst = order[k - 2U];

    if (std::isfinite(fvals[best]) && std::isfinite(fvals[worst]) && (fvals[worst] - fvals[best]) < kNelderMeadFTol) {
      break;
    }

    std::vector<double> centroid(dim, 0.0);
    for (std::size_t i = 0; i < k; ++i) {
      if (i == worst)
        continue;
      for (std::size_t j = 0; j < dim; ++j)
        centroid[j] += simplex[i][j];
    }
    for (std::size_t j = 0; j < dim; ++j)
      centroid[j] /= static_cast<double>(dim);

    std::vector<double> reflected(dim);
    for (std::size_t j = 0; j < dim; ++j) {
      reflected[j] = centroid[j] + 1.0 * (centroid[j] - simplex[worst][j]);
    }
    clamp_vertex_bounded(reflected.data(), dim, bounds.data());
    const double f_reflected = objective_full(reflected.data(), dim, smooth_dim, y, m);

    if (f_reflected < fvals[second_worst] && f_reflected >= fvals[best]) {
      simplex[worst] = reflected;
      fvals[worst] = f_reflected;
      continue;
    }

    if (f_reflected < fvals[best]) {
      std::vector<double> expanded(dim);
      for (std::size_t j = 0; j < dim; ++j) {
        expanded[j] = centroid[j] + 2.0 * (reflected[j] - centroid[j]);
      }
      clamp_vertex_bounded(expanded.data(), dim, bounds.data());
      const double f_expanded = objective_full(expanded.data(), dim, smooth_dim, y, m);
      if (f_expanded < f_reflected) {
        simplex[worst] = expanded;
        fvals[worst] = f_expanded;
      } else {
        simplex[worst] = reflected;
        fvals[worst] = f_reflected;
      }
      continue;
    }

    std::vector<double> contracted(dim);
    if (f_reflected < fvals[worst]) {
      for (std::size_t j = 0; j < dim; ++j) {
        contracted[j] = centroid[j] + 0.5 * (reflected[j] - centroid[j]);
      }
    } else {
      for (std::size_t j = 0; j < dim; ++j) {
        contracted[j] = centroid[j] + 0.5 * (simplex[worst][j] - centroid[j]);
      }
    }
    clamp_vertex_bounded(contracted.data(), dim, bounds.data());
    const double f_contracted = objective_full(contracted.data(), dim, smooth_dim, y, m);
    if (f_contracted < std::min(f_reflected, fvals[worst])) {
      simplex[worst] = contracted;
      fvals[worst] = f_contracted;
      continue;
    }

    for (std::size_t i = 0; i < k; ++i) {
      if (i == best)
        continue;
      for (std::size_t j = 0; j < dim; ++j) {
        simplex[i][j] = simplex[best][j] + 0.5 * (simplex[i][j] - simplex[best][j]);
      }
      clamp_vertex_bounded(simplex[i].data(), dim, bounds.data());
      fvals[i] = objective_full(simplex[i].data(), dim, smooth_dim, y, m);
    }
  }

  std::size_t best_idx = 0;
  for (std::size_t i = 1; i < k; ++i) {
    if (fvals[i] < fvals[best_idx])
      best_idx = i;
  }
  *out_params = simplex[best_idx];
  return fvals[best_idx];
}

// ---------------------------------------------------------------------------
// Fit + diagnostics on a resampled, deduplicated, evenly spaced series
// ---------------------------------------------------------------------------

// Fits a Holt-Winters model to `y` with the given seasonality `m`. On
// success populates `*out` (parameters, final state, residuals, MAE,
// RMSE, MASE, SMAPE) and returns `true`. Writes `*out_err` and returns
// `false` on degenerate inputs (n < 2, or seasonal m with n < 2*m).
bool fit_holt_winters(const std::vector<double>& y, std::uint32_t m, HoltWintersFit* out, Value* out_err) {
  const std::size_t n = y.size();
  if (n < 2U) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  if (m >= 2U && n < 2U * static_cast<std::size_t>(m)) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  const std::size_t smooth_dim = (m >= 2U) ? 3U : 2U;
  double alpha = 0.0;
  double beta = 0.0;
  double gamma = 0.0;

  // Stage 1: tune (alpha, beta[, gamma]) at the textbook two-season-means
  // initialisation. This is the conservative path that has shipped since
  // FORECAST.ETS landed; on most series it sits at or near the global
  // optimum and is cheap to compute.
  const double sse_basic = nelder_mead(y, m, smooth_dim, &alpha, &beta, &gamma);

  // Stage 2: jointly tune the smoothing parameters and the initial state
  // (L_0, B_0, S_0..S_{m-1}). The 3D optimum can sit in a basin that the
  // smooth-only stage cannot reach because the initial state is fixed; on
  // strongly cyclic data this stage finds the basin Mac Excel converges
  // to. Multi-start over a small seed grid widens basin coverage further.
  double init_level = 0.0;
  double init_trend = 0.0;
  std::vector<double> init_season;
  compute_default_init(y, m, &init_level, &init_trend, &init_season);
  const std::size_t full_dim = smooth_dim + 2U + ((m >= 2U) ? static_cast<std::size_t>(m) : 0U);
  const std::vector<ParamBounds> full_bounds = make_full_bounds(y, full_dim, smooth_dim);

  // Seed grid. Each entry supplies a (alpha, beta[, gamma]) anchor; the
  // initial-state block is always seeded from the textbook default so the
  // simplex can move it as needed. Seeds cover low/mid/high alpha to
  // hedge against multimodal SSE landscapes.
  static constexpr double kSmoothSeeds[][3] = {
      {0.2, 0.1, 0.1},
      {0.5, 0.1, 0.1},
      {0.8, 0.05, 0.1},
      {0.5, 0.5, 0.5},
  };
  double sse_full = std::numeric_limits<double>::infinity();
  std::vector<double> full_params;
  for (const auto& smooth_seed : kSmoothSeeds) {
    std::vector<double> seed(full_dim, 0.0);
    seed[0] = smooth_seed[0];
    seed[1] = smooth_seed[1];
    if (smooth_dim == 3U)
      seed[2] = smooth_seed[2];
    seed[smooth_dim] = init_level;
    seed[smooth_dim + 1U] = init_trend;
    if (m >= 2U) {
      for (std::uint32_t i = 0; i < m; ++i) {
        seed[smooth_dim + 2U + i] = init_season[i];
      }
    }
    std::vector<double> params;
    const double sse = nelder_mead_full(y, m, full_dim, smooth_dim, seed.data(),
                                        /*seed_step=*/0.1, full_bounds, &params);
    if (sse < sse_full) {
      sse_full = sse;
      full_params = std::move(params);
    }
  }

  // Pick the better fit. The two stages share an objective (SSE on the
  // same series), so a direct comparison is sound. On weakly-structured
  // data the two-stage paths agree to numerical noise; on strongly-cyclic
  // data the joint fit is materially better.
  if (sse_full + 1e-9 < sse_basic && std::isfinite(sse_full)) {
    alpha = full_params[0];
    beta = full_params[1];
    gamma = (smooth_dim == 3U) ? full_params[2] : 0.0;
    init_level = full_params[smooth_dim];
    init_trend = full_params[smooth_dim + 1U];
    if (m >= 2U) {
      for (std::uint32_t i = 0; i < m; ++i) {
        init_season[i] = full_params[smooth_dim + 2U + i];
      }
    }
    (void)simulate_with_init(y, m, alpha, beta, gamma, init_level, init_trend,
                             init_season.empty() ? nullptr : init_season.data(), out);
  } else {
    // Re-run simulation at the fitted parameters to populate residuals
    // and final state.
    (void)simulate(y, m, alpha, beta, gamma, out);
  }

  // Diagnostic statistics. `start` = first index that has a forecast.
  const std::size_t start = (m == 1U) ? 1U : static_cast<std::size_t>(m);
  if (start >= n) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  const std::size_t k = n - start;
  double sum_abs_err = 0.0;
  double sum_sq_err = 0.0;
  double smape_num = 0.0;
  std::size_t smape_n = 0U;
  for (std::size_t t = start; t < n; ++t) {
    const double err = out->residuals[t];
    sum_abs_err += std::fabs(err);
    sum_sq_err += err * err;
    const double y_hat = y[t] - err;
    const double denom = std::fabs(y[t]) + std::fabs(y_hat);
    if (denom > 0.0) {
      // SMAPE in the symmetric form: |F - y| / ((|F| + |y|) / 2).
      smape_num += std::fabs(err) / (denom * 0.5);
      ++smape_n;
    }
  }
  out->mae = sum_abs_err / static_cast<double>(k);
  out->rmse = std::sqrt(sum_sq_err / static_cast<double>(k));
  out->smape = (smape_n > 0U) ? (smape_num / static_cast<double>(smape_n)) : 0.0;

  // MASE: mean absolute error / mean absolute first-difference of the
  // training series.
  double sum_diff = 0.0;
  for (std::size_t i = 1; i < n; ++i) {
    sum_diff += std::fabs(y[i] - y[i - 1U]);
  }
  if (sum_diff > 0.0) {
    const double scale = sum_diff / static_cast<double>(n - 1U);
    out->mase = out->mae / scale;
  } else {
    out->mase = 0.0;
  }
  return true;
}

// ---------------------------------------------------------------------------
// Front-end: shared input preprocessing
// ---------------------------------------------------------------------------

// Reads an optional scalar argument as a non-negative integer (truncated
// toward zero). Blank yields `default_value`. Errors propagate. Out-of-
// domain (negative, non-finite, fractional outside the trunc rule) is
// flagged via `*out_err` with the supplied error code.
bool read_int_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                  std::int64_t default_value, std::int64_t* out, Value* out_err, ErrorCode oo_domain) {
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
  const double d = coerced.value();
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

// Reads an optional scalar number argument. Same semantics as
// `read_int_arg` minus the truncation step.
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

  ForecastOptions opts;
  if (!read_forecast_options(call, arity, 3U, arena, registry, ctx, &opts, &err)) {
    return err;
  }

  Preprocessed pre;
  if (!preprocess(call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx, opts.seasonality,
                  opts.data_completion, opts.aggregation, &pre, &err)) {
    return err;
  }

  // target_date must lie at or after the first timeline value.
  if (target_date < pre.t0) {
    return Value::error(ErrorCode::Num);
  }

  HoltWintersFit fit;
  if (!fit_holt_winters(pre.resampled.y, pre.m, &fit, &err)) {
    return err;
  }

  // Derive forecast horizon. The last training point lives at grid index
  // `n - 1`; `target_date` lives at grid index `target_idx`. Negative h
  // (i.e. target_date inside the training window) returns an interpolated
  // in-sample value; we still treat that as a valid query and compute
  // L + h*B + S correction.
  const std::int64_t n = static_cast<std::int64_t>(pre.resampled.y.size());
  const std::int64_t target_idx = target_step_index(target_date, pre.t0, pre.step);
  const std::int64_t h = target_idx - (n - 1);
  const double level_term = fit.level + static_cast<double>(h) * fit.trend;
  const double seasonal_term = seasonal_correction(fit.season, fit.m, h);
  const double forecast = level_term + seasonal_term;
  if (!std::isfinite(forecast)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(forecast);
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

  ForecastOptions opts;
  if (!read_forecast_options(call, arity, 4U, arena, registry, ctx, &opts, &err)) {
    return err;
  }

  Preprocessed pre;
  if (!preprocess(call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx, opts.seasonality,
                  opts.data_completion, opts.aggregation, &pre, &err)) {
    return err;
  }
  if (target_date < pre.t0) {
    return Value::error(ErrorCode::Num);
  }

  HoltWintersFit fit;
  if (!fit_holt_winters(pre.resampled.y, pre.m, &fit, &err)) {
    return err;
  }

  const std::int64_t n = static_cast<std::int64_t>(pre.resampled.y.size());
  const std::int64_t target_idx = target_step_index(target_date, pre.t0, pre.step);
  const std::int64_t h = target_idx - (n - 1);
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
  if (!std::isfinite(hw)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(hw);
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
  if (!std::isfinite(result)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(result);
}

}  // namespace eval
}  // namespace formulon
