// Copyright 2026 libraz. Licensed under the MIT License.
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

// Materialises an AST argument as an `ArrayValue`. Scalar inputs wrap to
// 1x1. On failure returns the propagating error via `*out_err`.
bool resolve_array(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   const ArrayValue** out, Value* out_err) {
  const Value v = eval_node_as_array(arg, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  if (!v.is_array()) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *out = v.as_array();
  return true;
}

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
  // Compute mean and total variance once.
  double mean = 0.0;
  for (double v : y)
    mean += v;
  mean /= static_cast<double>(n);
  double var = 0.0;
  for (double v : y) {
    const double d = v - mean;
    var += d * d;
  }
  if (var == 0.0) {
    return 1U;
  }
  const std::uint32_t max_lag_u = std::min<std::uint32_t>(static_cast<std::uint32_t>(n / 2U), kMaxSeasonalityDetectLag);
  if (max_lag_u < 2U) {
    return 1U;
  }
  double best_acf = 0.0;
  std::uint32_t best_lag = 1U;
  for (std::uint32_t k = 2U; k <= max_lag_u; ++k) {
    double num = 0.0;
    for (std::size_t i = k; i < n; ++i) {
      num += (y[i] - mean) * (y[i - k] - mean);
    }
    const double r = num / var;
    if (std::fabs(r) > std::fabs(best_acf)) {
      best_acf = r;
      best_lag = k;
    }
  }
  if (std::fabs(best_acf) >= kAcfDetectionThreshold) {
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

// Runs the Holt-Winters smoothing recurrences with the given parameters
// over `y`, populates `*out` (level / trend / season / residuals), and
// returns the in-sample SSE used by the optimiser.
//
// Recurrences (additive seasonality):
//   L_t = alpha (y_t - S_{t-m}) + (1-alpha)(L_{t-1} + B_{t-1})
//   B_t = beta  (L_t - L_{t-1}) + (1-beta)  B_{t-1}
//   S_t = gamma (y_t - L_t)     + (1-gamma) S_{t-m}
// Forecast residual at step t is `y_t - (L_{t-1} + B_{t-1} + S_{t-m})`.
double simulate(const std::vector<double>& y, std::uint32_t m, double alpha, double beta, double gamma,
                HoltWintersFit* out) {
  const std::size_t n = y.size();
  // Two-season-means initialisation. With m == 1 we use Holt's double
  // exponential: L_0 = y_0, B_0 = y_1 - y_0, no seasonal vector.
  double level = 0.0;
  double trend = 0.0;
  std::vector<double> season(m == 1U ? 0U : static_cast<std::size_t>(m), 0.0);
  if (m == 1U) {
    level = y[0];
    trend = (n >= 2U) ? (y[1] - y[0]) : 0.0;
  } else {
    double mean1 = 0.0;
    for (std::uint32_t i = 0; i < m; ++i)
      mean1 += y[i];
    mean1 /= static_cast<double>(m);
    double mean2 = 0.0;
    if (n >= 2U * static_cast<std::size_t>(m)) {
      for (std::uint32_t i = 0; i < m; ++i)
        mean2 += y[m + i];
      mean2 /= static_cast<double>(m);
    } else {
      // Defensive -- caller should already have rejected this case.
      mean2 = mean1;
    }
    level = mean1;
    trend = (mean2 - mean1) / static_cast<double>(m);
    double s_sum = 0.0;
    for (std::uint32_t i = 0; i < m; ++i) {
      season[i] = y[i] - mean1;
      s_sum += season[i];
    }
    const double s_mean = s_sum / static_cast<double>(m);
    for (std::uint32_t i = 0; i < m; ++i)
      season[i] -= s_mean;
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
  const std::size_t dim = (m >= 2U) ? 3U : 2U;
  double alpha = 0.0;
  double beta = 0.0;
  double gamma = 0.0;
  (void)nelder_mead(y, m, dim, &alpha, &beta, &gamma);
  // Re-run simulation at the fitted parameters to populate residuals
  // and final state.
  (void)simulate(y, m, alpha, beta, gamma, out);

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

// Aggregated state from the preprocessing pipeline. Returned to each
// front-end after timeline / values have been paired, sorted, aggregated,
// and resampled onto an evenly spaced grid.
struct Preprocessed {
  Series resampled;      // evenly spaced (t, y) pairs.
  double step = 0.0;     // grid step (median delta-t of original timeline).
  double t0 = 0.0;       // first original timeline value (== resampled.t[0]).
  std::uint32_t m = 1U;  // resolved seasonality length.
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
  if (!resolve_array(timeline_node, arena, registry, ctx, &timeline_arr, out_err)) {
    return false;
  }
  if (!resolve_array(values_node, arena, registry, ctx, &values_arr, out_err)) {
    return false;
  }
  const std::size_t n_t = static_cast<std::size_t>(timeline_arr->rows) * timeline_arr->cols;
  const std::size_t n_v = static_cast<std::size_t>(values_arr->rows) * values_arr->cols;
  if (n_t != n_v) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  if (n_t < 2U) {
    *out_err = Value::error(ErrorCode::NA);
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
  if (node == nullptr) {
    *out = AggregationMode::kAverage;
    return true;
  }
  std::int64_t v = 1;
  if (!read_int_arg(*node, arena, registry, ctx, 1, &v, out_err, ErrorCode::Value)) {
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
  if (node == nullptr) {
    *out = 1;
    return true;
  }
  std::int64_t v = 1;
  if (!read_int_arg(*node, arena, registry, ctx, 1, &v, out_err, ErrorCode::Value)) {
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
  if (node == nullptr) {
    *out = 1;
    return true;
  }
  std::int64_t v = 1;
  if (!read_int_arg(*node, arena, registry, ctx, 1, &v, out_err, ErrorCode::Num)) {
    return false;
  }
  if (v < 0 || v > static_cast<std::int64_t>(kMaxSeasonalityArg)) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  *out = v;
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
  const Value target_v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (target_v.is_error()) {
    return target_v;
  }
  auto target_coerced = coerce_to_number(target_v);
  if (!target_coerced) {
    return Value::error(target_coerced.error());
  }
  const double target_date = target_coerced.value();
  if (!std::isfinite(target_date)) {
    return Value::error(ErrorCode::Num);
  }

  // arg 3..5: optional integer args.
  std::int64_t seasonality = 1;
  int data_completion = 1;
  AggregationMode aggregation = AggregationMode::kAverage;
  Value err = Value::blank();
  if (!read_seasonality(arity > 3U ? &call.as_call_arg(3) : nullptr, arena, registry, ctx, &seasonality, &err)) {
    return err;
  }
  if (!read_data_completion(arity > 4U ? &call.as_call_arg(4) : nullptr, arena, registry, ctx, &data_completion,
                            &err)) {
    return err;
  }
  if (!read_aggregation(arity > 5U ? &call.as_call_arg(5) : nullptr, arena, registry, ctx, &aggregation, &err)) {
    return err;
  }

  Preprocessed pre;
  if (!preprocess(call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx, seasonality, data_completion,
                  aggregation, &pre, &err)) {
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

  const Value target_v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (target_v.is_error()) {
    return target_v;
  }
  auto target_coerced = coerce_to_number(target_v);
  if (!target_coerced) {
    return Value::error(target_coerced.error());
  }
  const double target_date = target_coerced.value();
  if (!std::isfinite(target_date)) {
    return Value::error(ErrorCode::Num);
  }

  Value err = Value::blank();
  double confidence = 0.95;
  if (arity > 3U) {
    if (!read_double_arg(call.as_call_arg(3), arena, registry, ctx, 0.95, &confidence, &err)) {
      return err;
    }
  }
  if (!std::isfinite(confidence) || confidence <= 0.0 || confidence >= 1.0) {
    return Value::error(ErrorCode::Num);
  }

  std::int64_t seasonality = 1;
  int data_completion = 1;
  AggregationMode aggregation = AggregationMode::kAverage;
  if (!read_seasonality(arity > 4U ? &call.as_call_arg(4) : nullptr, arena, registry, ctx, &seasonality, &err)) {
    return err;
  }
  if (!read_data_completion(arity > 5U ? &call.as_call_arg(5) : nullptr, arena, registry, ctx, &data_completion,
                            &err)) {
    return err;
  }
  if (!read_aggregation(arity > 6U ? &call.as_call_arg(6) : nullptr, arena, registry, ctx, &aggregation, &err)) {
    return err;
  }

  Preprocessed pre;
  if (!preprocess(call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx, seasonality, data_completion,
                  aggregation, &pre, &err)) {
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
  std::int64_t h = target_idx - (n - 1);
  // Half-width grows with sqrt(h); for h <= 0 (target inside the
  // training window) clamp the horizon to 1 so the half-width stays
  // strictly positive. ORACLE-PENDING: Excel may return 0 or RMSE
  // here; calibrate against the oracle.
  if (h < 1)
    h = 1;

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
  int data_completion = 1;
  AggregationMode aggregation = AggregationMode::kAverage;
  if (!read_data_completion(arity > 2U ? &call.as_call_arg(2) : nullptr, arena, registry, ctx, &data_completion,
                            &err)) {
    return err;
  }
  if (!read_aggregation(arity > 3U ? &call.as_call_arg(3) : nullptr, arena, registry, ctx, &aggregation, &err)) {
    return err;
  }

  // Force auto-detect by passing seasonality = 1 to preprocess; the
  // resampling path is unchanged.
  Preprocessed pre;
  if (!preprocess(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx, /*seasonality=*/1, data_completion,
                  aggregation, &pre, &err)) {
    return err;
  }
  return Value::number(static_cast<double>(pre.m));
}

Value eval_forecast_ets_stat_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                  const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 6U) {
    return Value::error(ErrorCode::Value);
  }

  // arg 2: statistic_type scalar.
  const Value stat_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
  if (stat_v.is_error()) {
    return stat_v;
  }
  auto stat_coerced = coerce_to_number(stat_v);
  if (!stat_coerced) {
    return Value::error(stat_coerced.error());
  }
  const double stat_d = stat_coerced.value();
  if (!std::isfinite(stat_d)) {
    return Value::error(ErrorCode::Num);
  }
  const std::int64_t stat_type = static_cast<std::int64_t>(stat_d);  // truncate
  if (stat_type < 1 || stat_type > 8) {
    return Value::error(ErrorCode::Num);
  }

  Value err = Value::blank();
  std::int64_t seasonality = 1;
  int data_completion = 1;
  AggregationMode aggregation = AggregationMode::kAverage;
  if (!read_seasonality(arity > 3U ? &call.as_call_arg(3) : nullptr, arena, registry, ctx, &seasonality, &err)) {
    return err;
  }
  if (!read_data_completion(arity > 4U ? &call.as_call_arg(4) : nullptr, arena, registry, ctx, &data_completion,
                            &err)) {
    return err;
  }
  if (!read_aggregation(arity > 5U ? &call.as_call_arg(5) : nullptr, arena, registry, ctx, &aggregation, &err)) {
    return err;
  }

  Preprocessed pre;
  if (!preprocess(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx, seasonality, data_completion,
                  aggregation, &pre, &err)) {
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
