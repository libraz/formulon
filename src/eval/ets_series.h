//
// Series preparation shared by the FORECAST.ETS family: reading the
// (timeline, value) cells, ordering them, detecting the timeline step,
// building the regular grid, locating forecast targets on it and detecting
// the seasonality length.

#ifndef FORMULON_EVAL_ETS_SERIES_H_
#define FORMULON_EVAL_ETS_SERIES_H_

#include <cmath>
#include <cstdint>
#include <optional>
#include <vector>

#include "value.h"

namespace formulon {
namespace eval {
namespace ets {

/// Bound on the step and day counts the family turns into integers: below
/// 2^53 every integer is exact in a double, and the calendar and variance
/// arithmetic on such counts stays inside the int64 range.
inline constexpr double kMaxWholeCount = 9007199254740992.0;

/// Truncates `v` toward zero, or returns nullopt when `v` is non-finite or its
/// magnitude is not below `kMaxWholeCount`. The family's only narrowing of a
/// double to an integer count.
inline std::optional<std::int64_t> whole_count(double v) noexcept {
  if (!(std::fabs(v) < kMaxWholeCount)) {
    return std::nullopt;
  }
  return static_cast<std::int64_t>(v);
}

/// Aggregation applied to values sharing a grid node (option 1..7).
enum class AggregationMode : std::uint32_t {
  kAverage = 1U,
  kCount = 2U,
  kCountA = 3U,
  kMax = 4U,
  kMedian = 5U,
  kMin = 6U,
  kSum = 7U,
};

/// One input cell reduced to what the series reader distinguishes.
struct SeriesCell {
  enum class Kind : std::uint8_t { kNumber, kBlank, kText, kBool, kError };
  Kind kind = Kind::kBlank;
  double number = 0.0;
  ErrorCode error = ErrorCode::NA;
};

/// Timeline step and calendar classification.
struct Timeline {
  /// Step between grid nodes (days, or the histogram step for calendar data).
  double step = 0.0;
  /// True when the first and last timestamps are valid date serials.
  bool is_date = false;
  /// Months per node on a calendar timeline, otherwise 0 (or -1 before
  /// classification).
  std::int64_t months = 0;
  /// Number of steps in the chosen histogram window.
  std::int64_t count = 0;
  /// Timestamp of the last grid node.
  double last = 0.0;
};

/// Regularised series ready for fitting.
struct PreparedSeries {
  std::vector<double> y;
  Timeline timeline;
};

/// Reads, sorts and regularises the series. `data_completion` is 0 (fill
/// with zero) or 1 (interpolate). On failure writes `*err` and returns false.
bool prepare_series(const std::vector<SeriesCell>& values, const std::vector<SeriesCell>& timeline, int data_completion,
                    AggregationMode aggregation, PreparedSeries* out, ErrorCode* err);

/// Locates `target` on the grid as a (possibly fractional) step count from
/// the last node; `*interpolated` is set when it falls between two steps. A
/// calendar target beyond `kMaxWholeCount` days is `#NUM!`.
bool target_position(const Timeline& timeline, double target, double* position, bool* interpolated, ErrorCode* err);

/// Auto-detected seasonality length (>= 1); 1 means no seasonality.
std::uint32_t detect_seasonality(const std::vector<double>& y);

}  // namespace ets
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_ETS_SERIES_H_
