// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the date/time primitives declared in date_time.h. The
// heavy lifting is Howard Hinnant's civil_from_days / days_from_civil pair
// (epoch 1970-01-01), wrapped by two Excel-aware offset constants that
// absorb the 1900 leap-year bug.

#include "eval/date_time.h"

#include <cmath>
#include <cstdint>

namespace formulon {
namespace eval {
namespace date_time {
namespace {

// Excel's day 1 is 1900-01-01. If we use the (y, m, d) -> days-since-1970
// mapping from Hinnant's paper, we need offsets that convert in both
// directions.
//
// * `kExcelBaseBeforeGhost` = -days_from_civil(1899, 12, 31) = 25568.
//   Used for serials 1..59 (strictly before the ghost 1900-02-29).
//   serial = days_from_civil(y, m, d) + kExcelBaseBeforeGhost.
//
// * `kExcelBaseAfterGhost` = -days_from_civil(1899, 12, 30) = 25569.
//   Used for serials >= 61 (strictly after the ghost). The one-day gap
//   is what absorbs the fictitious 1900-02-29.
constexpr std::int64_t kExcelBaseBeforeGhost = 25568;
constexpr std::int64_t kExcelBaseAfterGhost = 25569;

// days_from_civil(1900, 2, 28) in the 1970 epoch. Any serial mapped through
// `kExcelBaseBeforeGhost` must produce a civil day <= this value; anything
// past it has crossed the ghost 1900-02-29 and must use the post-ghost
// base.
constexpr std::int64_t kCivilDays1900Feb28 = -25509;

// Seconds in a day. The fractional part of a serial is scaled by this to
// recover wall-clock seconds.
constexpr double kSecondsPerDay = 86400.0;

}  // namespace

std::int64_t days_from_civil(int y, unsigned m, unsigned d) noexcept {
  // Hinnant, "chrono-Compatible Low-Level Date Algorithms", algorithm (7).
  // Normalises any out-of-range month/day as a side effect.
  y -= static_cast<int>(m <= 2);
  const int era = (y >= 0 ? y : y - 399) / 400;
  const unsigned yoe = static_cast<unsigned>(y - era * 400);                 // [0, 399]
  const unsigned doy = (153u * (m > 2 ? m - 3 : m + 9) + 2u) / 5u + d - 1u;  // [0, 365]
  const unsigned doe = yoe * 365u + yoe / 4u - yoe / 100u + doy;             // [0, 146096]
  return static_cast<std::int64_t>(era) * 146097 + static_cast<std::int64_t>(doe) - 719468;
}

YMD civil_from_days(std::int64_t z) noexcept {
  // Hinnant's inverse algorithm (matching (7)).
  z += 719468;
  const std::int64_t era = (z >= 0 ? z : z - 146096) / 146097;
  const unsigned doe = static_cast<unsigned>(z - era * 146097);                    // [0, 146096]
  const unsigned yoe = (doe - doe / 1460u + doe / 36524u - doe / 146096u) / 365u;  // [0, 399]
  const int y = static_cast<int>(yoe) + static_cast<int>(era) * 400;
  const unsigned doy = doe - (365u * yoe + yoe / 4u - yoe / 100u);  // [0, 365]
  const unsigned mp = (5u * doy + 2u) / 153u;                       // [0, 11]
  const unsigned d = doy - (153u * mp + 2u) / 5u + 1u;              // [1, 31]
  const unsigned m = mp < 10u ? mp + 3u : mp - 9u;                  // [1, 12]
  return YMD{y + static_cast<int>(m <= 2u), m, d};
}

YMD ymd_from_serial(double serial_floor) noexcept {
  const std::int64_t s = static_cast<std::int64_t>(std::floor(serial_floor));
  if (s == 60) {
    // Preserve Excel's ghost day.
    return YMD{1900, 2u, 29u};
  }
  const std::int64_t base = (s < 60) ? kExcelBaseBeforeGhost : kExcelBaseAfterGhost;
  return civil_from_days(s - base);
}

double serial_from_ymd(int y, unsigned m, unsigned d) noexcept {
  // Excel reserves serial 60 for the fictitious 1900-02-29 that the 1900
  // leap-year bug retains. `days_from_civil` would normalise (1900, 2, 29)
  // to the real civil day 1900-03-01, which maps through
  // `kExcelBaseAfterGhost` to serial 61 (the correct answer for 1900-03-01).
  // Intercept the literal input shape here BEFORE normalisation so
  // DATE(1900, 2, 29) still returns 60. Callers that reach the ghost day
  // via month/day overflow (e.g. DATE(1900, 1, 60) -> 1900-03-01) are
  // unaffected because their raw (y, m, d) triple is not (1900, 2, 29).
  if (y == 1900 && m == 2u && d == 29u) {
    return 60.0;
  }
  const std::int64_t civil = days_from_civil(y, m, d);
  // Dates strictly before 1900-03-01 use the "before-ghost" base; the rest
  // (including negative civil days for pre-1900 inputs, which the caller
  // will typically reject upstream) use the "after-ghost" base so the 1900
  // leap-year bug is preserved for every valid Excel date.
  const std::int64_t base = (civil <= kCivilDays1900Feb28) ? kExcelBaseBeforeGhost : kExcelBaseAfterGhost;
  return static_cast<double>(civil + base);
}

HMS hms_from_fraction(double serial) noexcept {
  // Extract the positive fractional part: `fmod` preserves the sign of the
  // dividend, which we compensate for when serial < 0. The date-aware
  // builtins reject negative serials, so this branch is only exercised when
  // the fractional part is exactly 0 after the rejection.
  double frac = serial - std::floor(serial);
  if (frac < 0.0) {
    frac += 1.0;
  }
  // Round to the nearest second so `HOUR(TIME(h, m, s))` is exact for
  // integer h/m/s triples. Total seconds are modulo 86,400 so we never
  // leak into the next day.
  std::int64_t total = static_cast<std::int64_t>(std::llround(frac * kSecondsPerDay));
  total %= 86400;
  if (total < 0) {
    total += 86400;
  }
  const unsigned h = static_cast<unsigned>(total / 3600);
  const unsigned m = static_cast<unsigned>((total / 60) % 60);
  const unsigned s = static_cast<unsigned>(total % 60);
  return HMS{h, m, s};
}

}  // namespace date_time
}  // namespace eval
}  // namespace formulon
