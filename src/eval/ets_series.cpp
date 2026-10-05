//
// Series preparation for the FORECAST.ETS family; the contract is declared in
// `eval/ets_series.h`.

#include "eval/ets_series.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "utils/error.h"
#include "utils/index_sort.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace ets {
namespace {

// Step-size deviation tolerance for the resample grid. A consecutive delta-t
// is required to fall within +/-`kStepDeviation` of the median step
// (Excel's documented +/-30% tolerance). ORACLE-PENDING.
constexpr double kStepDeviation = 0.3;

// Cap on autocorrelation lags inspected by the seasonality detector. Excel
// allows a maximum seasonality of 8760 (hours in a year), but for the ACF
// scan the detection ceiling is 52 (weeks per year) which matches the
// typical "weekly seasonality on weekly data" calibration. ORACLE-PENDING.
constexpr std::uint32_t kMaxSeasonalityDetectLag = 52U;

// Minimum |ACF| magnitude required to declare a detected seasonality.
// Below this the auto-detector returns m = 1. ORACLE-PENDING.
constexpr double kAcfDetectionThreshold = 0.3;

// Relative floor on the linear-trend residual below which the detrended
// series is treated as pure rounding noise and no period is reported. The
// comparison is against the raw mean-centred variance, so it scales with
// the data instead of assuming a magnitude.
constexpr double kDetrendedResidualEpsilon = 1e-20;

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

}  // namespace

// Sorts `src` by t (stable) and collapses runs of identical t via `mode`.
// Returns the new (deduped, sorted) series.
Series sort_and_aggregate(const Series& src, AggregationMode mode) {
  const std::size_t n = src.t.size();
  const auto by_t = [&src](std::uint32_t a, std::uint32_t b) { return src.t[a] < src.t[b]; };
  std::vector<std::uint32_t> idx;
  sorted_index_order(idx, static_cast<std::uint32_t>(n), make_index_less(by_t));
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

}  // namespace ets
}  // namespace eval
}  // namespace formulon
