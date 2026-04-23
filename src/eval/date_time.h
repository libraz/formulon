// Copyright 2026 libraz. Licensed under the MIT License.
//
// Date / time primitives shared by the calendar-aware built-ins (DATE, YEAR,
// MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY, EDATE, EOMONTH, DAYS, ...).
//
// Excel represents date-times as double-precision serial numbers: the integer
// part is a day count anchored at 1900-01-01 (serial 1), the fractional part
// is the time-of-day as a fraction of 86,400 seconds.
//
// Two peculiarities of the serial scheme must be preserved:
//
//   1. Excel keeps Lotus 1-2-3's 1900 leap-year bug: serial 60 is the
//      non-existent "1900-02-29". Every serial >= 61 is therefore one greater
//      than the proleptic-Gregorian day offset from 1899-12-31.
//   2. Negative serials are not accepted by the date-aware functions; the
//      calling builtins map them to `#NUM!`.
//
// The helpers below are intentionally minimal: each calendar conversion is
// implemented via Howard Hinnant's civil_from_days algorithm (public domain,
// see howardhinnant.github.io/date_algorithms.html). The same routine is used
// by the C++20 <chrono> year_month_day calendar; we port it directly to keep
// the engine C++17 and dependency-free.

#ifndef FORMULON_EVAL_DATE_TIME_H_
#define FORMULON_EVAL_DATE_TIME_H_

#include <cstdint>

namespace formulon {
namespace eval {
namespace date_time {

/// Calendar date triple used by the conversion helpers. Fields are *not*
/// range-checked: the civil <-> days routines accept arbitrary month/day
/// values and normalise them.
struct YMD {
  int y;
  unsigned m;
  unsigned d;
};

/// Time-of-day triple produced by `hms_from_fraction`.
struct HMS {
  unsigned h;
  unsigned m;
  unsigned s;
};

/// Converts a Gregorian (y, m, d) to days since 1970-01-01.
///
/// Accepts any `y` and any `m` / `d`; out-of-range months/days are normalised
/// (e.g. month = 13 rolls into January of the following year). This is
/// Howard Hinnant's canonical `days_from_civil` algorithm, valid for all
/// representable `int` years.
std::int64_t days_from_civil(int y, unsigned m, unsigned d) noexcept;

/// Inverse of `days_from_civil`: converts a day count relative to
/// 1970-01-01 back into a calendar triple. The returned `m` is 1..12 and
/// `d` is the last valid day-of-month for the result.
YMD civil_from_days(std::int64_t days) noexcept;

/// Converts an Excel serial (integer part) to a calendar triple, honouring
/// the 1900 leap-year bug.
///
/// * `serial_floor == 60` is returned as the fictitious `{1900, 2, 29}`.
/// * `serial_floor < 60` maps against base 1899-12-31.
/// * `serial_floor > 60` maps against base 1899-12-30 (the off-by-one
///   correction that absorbs the ghost day).
///
/// Callers must ensure `serial_floor >= 0`; negative serials are invalid
/// and should be rejected at the builtin boundary.
YMD ymd_from_serial(double serial_floor) noexcept;

/// Inverse of `ymd_from_serial`. The input triple need not be a valid
/// calendar date: out-of-range months/days are normalised first (via
/// `days_from_civil`), so `serial_from_ymd(2026, 13, 1)` returns the serial
/// for `2027-01-01`.
///
/// The 1900 leap-year bug is reinserted for any result strictly after
/// 1900-02-28 (i.e. the Excel serial is one greater than the proleptic
/// offset for all dates from 1900-03-01 onward).
double serial_from_ymd(int y, unsigned m, unsigned d) noexcept;

/// Extracts hour/minute/second from the fractional part of a serial. The
/// total seconds in the day are rounded to the nearest integer; this
/// matches observed Excel behaviour for common values like `TIME(0,0,1)`
/// and keeps `HOUR(serial + TIME(h,m,s))` round-trip stable.
HMS hms_from_fraction(double serial) noexcept;

}  // namespace date_time
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DATE_TIME_H_
