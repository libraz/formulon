//
// Internal contract shared by the translation units behind `eval/ets_series.h`:
// bit-level numeric helpers, the calendar walk, pair sorting and the timeline
// step / grid stages.

#ifndef FORMULON_EVAL_ETS_SERIES_INTERNAL_H_
#define FORMULON_EVAL_ETS_SERIES_INTERNAL_H_

#include <cfloat>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <vector>

#include "eval/ets_series.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace ets {
namespace series_detail {

inline double div_zero(double a, double b) noexcept {
  if (std::fabs(b) < DBL_MIN) {
    return 0.0;
  }
  const double r = a / b;
  return (std::fabs(r) > DBL_MAX || std::fabs(r) < DBL_MIN) ? 0.0 : r;
}

inline double mul_flush(double a, double b) noexcept {
  return flush_subnormal(a * b);
}

inline bool top_bits_zero(double v) noexcept {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof bits);
  return (bits >> 48U) == 0U;
}

// Three-way compare on the sign-magnitude integer image: -0.0 sorts below
// +0.0 and NaN above every number.
inline int bit_compare(double a, double b) noexcept {
  std::int64_t ia = 0;
  std::int64_t ib = 0;
  std::memcpy(&ia, &a, sizeof ia);
  std::memcpy(&ib, &b, sizeof ib);
  if (ia == ib) {
    return 0;
  }
  int r = ia < ib ? -1 : 1;
  if ((ia & ib) < 0) {
    r = -r;
  }
  return r;
}

inline int sign_of(double v) noexcept {
  return bit_compare(v, 0.0);
}

// Magnitude truncated to an unsigned 32-bit value, saturating.
inline std::uint32_t to_uint32(double v) noexcept {
  v = std::fabs(v);
  if (v != v) {
    return 0U;
  }
  if (v >= 4294967295.0) {
    return 0xffffffffU;
  }
  return static_cast<std::uint32_t>(v);
}

inline bool sign_bit_set(double v) noexcept {
  std::int64_t bits = 0;
  std::memcpy(&bits, &v, sizeof bits);
  return bits < 0;
}

template <typename V>
auto& elem(V& v, std::int64_t i) {
  return v[static_cast<std::size_t>(i)];
}

// Statistics of the consecutive timeline differences.
struct StepStats {
  double min = DBL_MAX;
  double max = 0.0;
  // Non-positive steps (duplicates, and negative steps when `needs_sort`).
  std::int64_t nonpositive = 0;
  bool needs_sort = false;

  void add(double d) {
    if (bit_compare(d, min) == -1) {
      if (top_bits_zero(d)) {
        ++nonpositive;
      } else if (!std::isfinite(d)) {
        // Not counted.
      } else if (sign_bit_set(d)) {
        ++nonpositive;
        needs_sort = true;
      } else {
        min = d;
      }
    }
    if (bit_compare(d, max) == 1) {
      max = d;
    }
  }
};

/// Calendar date held as a (year, month, day) triple of a serial.
struct Ymd {
  std::int64_t year = 0;
  std::int64_t month = 0;
  std::int64_t day = 0;
};

/// Floor division for a positive or negative dividend.
std::int64_t floor_div(std::int64_t a, std::int64_t b) noexcept;

/// True when `x` is a non-negative serial within the 9999-12-31 calendar.
bool is_date_serial(double x) noexcept;

/// Calendar date of a serial that `whole_count` admitted.
Ymd ymd_of(double serial);

/// Moves the (year, month) pair in place by `delta` months, keeping the day,
/// and returns the serial of the result.
double add_months(Ymd* v, std::int64_t delta);

/// Nearest month-step node from `a` toward `b` and the signed step count.
bool month_walk(std::int64_t months, double a, double b, double* node, std::int64_t* count, ErrorCode* err);

/// Sorts `keys` ascending over [lo, hi], permuting `vals` (when non-null)
/// alongside; the order of equal keys is fixed.
void sort_pairs(std::vector<double>* keys, std::vector<double>* vals, std::int64_t lo, std::int64_t hi);

/// Detects the timeline step. `*irregular` is set when the series has gaps or
/// duplicate timestamps and must be rebuilt.
bool detect_step(const std::vector<double>& t, const StepStats& st, Timeline* tl, bool* irregular, ErrorCode* err);

/// Rebuilds the series on a regular grid, filling skipped nodes per the data
/// completion option. `*missing` counts the filled nodes.
bool build_grid(std::vector<double>* tv, std::vector<double>* yv, const Timeline& tl, int completion,
                AggregationMode agg, std::vector<double>* out_t, std::vector<double>* out_y, std::int64_t* missing,
                ErrorCode* err);

}  // namespace series_detail
}  // namespace ets
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_ETS_SERIES_INTERNAL_H_
