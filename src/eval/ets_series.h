//
// Series preparation shared by the FORECAST.ETS family: pairing, ordering and
// aggregating the (timeline, value) samples, resampling them onto an evenly
// spaced grid, and detecting the seasonality length.

#ifndef FORMULON_EVAL_ETS_SERIES_H_
#define FORMULON_EVAL_ETS_SERIES_H_

#include <cstdint>
#include <vector>

#include "value.h"

namespace formulon {
namespace eval {
namespace ets {

/// Excel's documented aggregation enumeration for repeated timeline entries.

enum class AggregationMode : std::uint32_t {
  kAverage = 1U,
  kCount = 2U,
  kCountA = 3U,
  kMax = 4U,
  kMedian = 5U,
  kMin = 6U,
  kSum = 7U,
};

/// Parallel (t, y) sample vectors.
struct Series {
  std::vector<double> t;
  std::vector<double> y;
};

/// Stable-sorts `src` by t and collapses runs of identical t via `mode`.
Series sort_and_aggregate(const Series& src, AggregationMode mode);

/// Median of consecutive deltas in a sorted timeline of at least two points.
double median_delta(const std::vector<double>& t);

/// Resamples a sorted, deduplicated series onto a grid of spacing `step`,
/// filling gaps per `data_completion` (0 zero-fill, 1 linear interpolation).
/// Writes `#NUM!` to `*out_err` and returns false on an irregular timeline.
bool resample_to_grid(const Series& sorted, double step, int data_completion, Series* out, Value* out_err);

/// Auto-detected seasonality length (>= 1); 1 means no seasonality.
std::uint32_t detect_seasonality(const std::vector<double>& y);

}  // namespace ets
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_ETS_SERIES_H_
