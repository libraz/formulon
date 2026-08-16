//
// Implementation of the date/time primitives declared in date_time.h. The
// heavy lifting is Howard Hinnant's civil_from_days / days_from_civil pair
// (epoch 1970-01-01), wrapped by two Excel-aware offset constants that
// absorb the 1900 leap-year bug.

#include "eval/date_time.h"

#include <chrono>
#include <cmath>
#include <cstdint>
#include <ctime>

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

// 1904 date system (Excel "Use 1904 date system"): serial 0 is
// 1904-01-01 and there is no fictitious 1900-02-29, so a single linear
// base suffices with no ghost handling. 1904-01-01 is serial 1462 in the
// 1900 system, hence this base is exactly `kExcelBaseAfterGhost - 1462`;
// every 1904 serial is 1462 less than the 1900 serial for the same day.
constexpr std::int64_t kExcel1904Base = kExcelBaseAfterGhost - 1462;

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

YMD ymd_from_serial(double serial_floor, bool date1904) noexcept {
  const std::int64_t s = static_cast<std::int64_t>(std::floor(serial_floor));
  if (date1904) {
    // 1904 system: no ghost day, single linear base.
    return civil_from_days(s - kExcel1904Base);
  }
  if (s == 60) {
    // Preserve Excel's ghost day.
    return YMD{1900, 2u, 29u};
  }
  const std::int64_t base = (s < 60) ? kExcelBaseBeforeGhost : kExcelBaseAfterGhost;
  return civil_from_days(s - base);
}

double serial_from_ymd(int y, unsigned m, unsigned d, bool date1904) noexcept {
  if (date1904) {
    // 1904 system: no ghost day; single linear base off 1904-01-01.
    return static_cast<double>(days_from_civil(y, m, d) + kExcel1904Base);
  }
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

double yearfrac_us30_360(int y1, unsigned m1, unsigned d1, int y2, unsigned m2, unsigned d2) noexcept {
  // NASD rule set (Excel's implementation):
  //   if d1 == 31                         -> d1 = 30
  //   if d2 == 31 and d1 >= 30            -> d2 = 30
  //   if last-day-of-Feb(d1)              -> d1 = 30
  //     and last-day-of-Feb(d2) too       -> d2 = 30
  auto last_day_of_feb = [](int y, unsigned m, unsigned d) {
    if (m != 2u) {
      return false;
    }
    const bool leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
    return d == (leap ? 29u : 28u);
  };
  const bool d1_last_feb = last_day_of_feb(y1, m1, d1);
  const bool d2_last_feb = last_day_of_feb(y2, m2, d2);
  if (d1_last_feb && d2_last_feb) {
    d2 = 30;
  }
  if (d1_last_feb) {
    d1 = 30;
  }
  if (d2 == 31u && d1 >= 30u) {
    d2 = 30;
  }
  if (d1 == 31u) {
    d1 = 30;
  }
  const double days = 360.0 * (y2 - y1) + 30.0 * (static_cast<double>(m2) - static_cast<double>(m1)) +
                      (static_cast<double>(d2) - static_cast<double>(d1));
  return days / 360.0;
}

double yearfrac_eu30_360(int y1, unsigned m1, unsigned d1, int y2, unsigned m2, unsigned d2) noexcept {
  if (d1 > 30u) {
    d1 = 30;
  }
  if (d2 > 30u) {
    d2 = 30;
  }
  const double days = 360.0 * (y2 - y1) + 30.0 * (static_cast<double>(m2) - static_cast<double>(m1)) +
                      (static_cast<double>(d2) - static_cast<double>(d1));
  return days / 360.0;
}

double yearfrac_actual_actual(int y1, unsigned m1, unsigned d1, int y2, unsigned m2, unsigned d2) noexcept {
  const std::int64_t days1 = days_from_civil(y1, m1, d1);
  const std::int64_t days2 = days_from_civil(y2, m2, d2);
  const double actual_days = static_cast<double>(days2 - days1);
  auto is_leap_year = [](int y) { return (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0); };
  // Excel's basis=1 (actual/actual): the denominator depends on whether the
  // span is contained in a single calendar year.
  //
  // Same-year span: denominator is 366 iff the year is a leap year, 365
  // otherwise -- the position of Feb 29 relative to the range does not
  // matter. (Observed behaviour: YEARFRAC("2008-03-01","2008-08-31",1) = 0.5
  // even though Feb 29 2008 sits before the start of the range.)
  //
  // Span of at most one year that crosses a single year boundary
  // (`y2 == y1 + 1` and end falls on or before start's anniversary): Excel
  // uses a SINGLE-year denominator, not the multi-year average. It is 366 iff
  // a Feb 29 falls in the closed interval [start, end], else 365. This is why
  // YEARFRAC("2020-01-01","2021-01-01",1) is exactly 1.0 (366/366) rather
  // than 366/365.5.
  //
  // Longer cross-year spans: denominator is the average calendar-year length
  // over the inclusive year range [y1, y2], where a leap year contributes 366
  // ONLY if its Feb 29 falls in the closed interval [start, end]. A span that
  // straddles a leap-year boundary without crossing Feb 29 averages as if
  // that leap year were a normal 365-day year.
  double denom = 0.0;
  if (y1 == y2) {
    denom = is_leap_year(y1) ? 366.0 : 365.0;
  } else if (y2 == y1 + 1 && (m1 > m2 || (m1 == m2 && d1 >= d2))) {
    bool has_feb29 = false;
    for (int y = y1; y <= y2; ++y) {
      if (!is_leap_year(y)) {
        continue;
      }
      const std::int64_t feb29 = days_from_civil(y, 2, 29);
      if (feb29 >= days1 && feb29 <= days2) {
        has_feb29 = true;
        break;
      }
    }
    denom = has_feb29 ? 366.0 : 365.0;
  } else {
    const int total_years = y2 - y1 + 1;
    int leap_years_in_span = 0;
    for (int y = y1; y <= y2; ++y) {
      if (!is_leap_year(y)) {
        continue;
      }
      const std::int64_t feb29 = days_from_civil(y, 2, 29);
      if (feb29 >= days1 && feb29 <= days2) {
        ++leap_years_in_span;
      }
    }
    const int non_leap_years = total_years - leap_years_in_span;
    denom = (static_cast<double>(leap_years_in_span) * 366.0 + static_cast<double>(non_leap_years) * 365.0) /
            static_cast<double>(total_years);
  }
  return actual_days / denom;
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

unsigned days_in_month(int y, unsigned m) noexcept {
  static constexpr unsigned kTable[] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
  if (m < 1u || m > 12u) {
    return 31u;
  }
  if (m == 2u) {
    const bool leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
    return leap ? 29u : 28u;
  }
  return kTable[m - 1u];
}

double basis_days_between(double a, double b, int basis) noexcept {
  if (basis == 0 || basis == 4) {
    const YMD ya = ymd_from_serial(a);
    const YMD yb = ymd_from_serial(b);
    const double yf = basis == 0 ? yearfrac_us30_360(ya.y, ya.m, ya.d, yb.y, yb.m, yb.d)
                                 : yearfrac_eu30_360(ya.y, ya.m, ya.d, yb.y, yb.m, yb.d);
    return yf * 360.0;
  }
  // Bases 1, 2, 3: actual days.
  return b - a;
}

CivilTime host_civil_time() noexcept {
  // Excel's clock functions are locale-bound: a worksheet shows the user's
  // local calendar, not UTC. `localtime_r` / `localtime_s` is therefore the
  // decomposition to use, and the result is handed back as civil fields so
  // no caller has to re-derive a timezone.
  using std::chrono::system_clock;
  const std::time_t stamp = system_clock::to_time_t(system_clock::now());
  std::tm local{};
#if defined(_WIN32)
  localtime_s(&local, &stamp);
#else
  localtime_r(&stamp, &local);
#endif
  CivilTime out{};
  out.date.y = local.tm_year + 1900;
  out.date.m = static_cast<unsigned>(local.tm_mon + 1);
  out.date.d = static_cast<unsigned>(local.tm_mday);
  out.time.h = static_cast<unsigned>(local.tm_hour);
  out.time.m = static_cast<unsigned>(local.tm_min);
  out.time.s = static_cast<unsigned>(local.tm_sec);
  return out;
}

}  // namespace date_time
}  // namespace eval
}  // namespace formulon
