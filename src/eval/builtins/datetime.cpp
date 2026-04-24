// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's calendar built-in functions: DATE, TIME, YEAR,
// MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY, EDATE, EOMONTH, DAYS, DAYS360,
// WEEKNUM, ISOWEEKNUM, YEARFRAC, DATEDIF, DATEVALUE, TIMEVALUE, NOW, TODAY.
//
// Most entries are eager, scalar-only (no range expansion). NOW / TODAY read
// the host's local wall clock via `std::chrono::system_clock` +
// `localtime_r` / `localtime_s` and return an Excel serial directly; volatile
// recalc semantics are the scheduler's responsibility, not this file's.
// DATEVALUE / TIMEVALUE parse a common subset of Excel's ja-JP-locale
// date/time strings (ISO, slash, and kanji forms — see the parser helpers
// below); wareki (Reiwa/Heisei/Showa/...) and current-year-defaulting
// partials are intentionally out of scope and documented as divergences in
// tests/divergence.yaml. Shared calendar helpers (serial <-> y/m/d, weekday
// arithmetic, time-of-day decomposition) live in `eval/date_time.h`; this
// file only layers Excel's argument-shape and error-handling rules on top.

#include "eval/builtins/datetime.h"

#include <chrono>
#include <cmath>
#include <cstdint>
#include <ctime>
#include <string_view>

#include "eval/coerce.h"
#include "eval/date_text_parse.h"
#include "eval/date_time.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// All of the calendar builtins below share two helpers:
//
//   * `coerce_serial` reads a single scalar argument, coerces it to a
//     number, and rejects negative values as `#NUM!`. The returned value
//     is the raw serial (fractional part intact); callers floor it
//     themselves when only the date component is relevant.
//   * `days_in_month` returns the last valid day-of-month for a given
//     (year, month) pair, respecting the Gregorian leap rule. Used by
//     EDATE / EOMONTH when clamping the day field after a month shift.
// Excel's maximum serial value: 2958465 = 9999-12-31. Serial 2958466 would
// map to 10000-01-01 which is outside the representable range.
constexpr double kExcelMaxSerial = 2958465.0;

Expected<double, ErrorCode> coerce_serial(const Value& v) {
  auto n = coerce_to_number(v);
  if (!n) {
    return n.error();
  }
  if (n.value() < 0.0 || n.value() >= kExcelMaxSerial + 1.0) {
    return ErrorCode::Num;
  }
  return n.value();
}

unsigned days_in_month(int y, unsigned m) noexcept {
  static constexpr unsigned kTable[] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
  if (m < 1u || m > 12u) {
    return 31u;  // Defensive; callers always normalise first.
  }
  if (m == 2u) {
    const bool leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
    return leap ? 29u : 28u;
  }
  return kTable[m - 1u];
}

/// DATE(year, month, day). Each argument is truncated (Excel floors toward
/// zero for date components). Years in `[0, 1900)` are expanded by adding
/// 1900; years outside `[0, 9999]` yield `#NUM!`. Month/day are passed
/// through the normalisation baked into `days_from_civil`, so rolls like
/// `DATE(2026, 13, 1) = 2027-01-01` fall out naturally.
Value Date_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto year_c = coerce_to_number(args[0]);
  if (!year_c) {
    return Value::error(year_c.error());
  }
  auto month_c = coerce_to_number(args[1]);
  if (!month_c) {
    return Value::error(month_c.error());
  }
  auto day_c = coerce_to_number(args[2]);
  if (!day_c) {
    return Value::error(day_c.error());
  }
  double y = std::trunc(year_c.value());
  const double m = std::trunc(month_c.value());
  const double d = std::trunc(day_c.value());
  // Excel's two-digit-year convention: `0 <= year < 1900` expands by 1900.
  if (y >= 0.0 && y < 1900.0) {
    y += 1900.0;
  }
  if (y < 1900.0 || y > 9999.0) {
    return Value::error(ErrorCode::Num);
  }
  // Normalise month into [1, 12] so the ymd -> days routine stays within a
  // representable range (it accepts any month, but keeping things tidy also
  // helps when building the Excel serial afterwards).
  // The conversion itself tolerates arbitrary (m, d); we only constrain y.
  const int yi = static_cast<int>(y);
  // `days_from_civil` accepts months in [1, 12] exclusively but tolerates
  // day-in-month overflow. Funnel month overflow into a year shift using
  // Python-style floor-division so negative months wrap correctly
  // (C++ integer division truncates toward zero, which would map month 0
  // to the *same* year instead of the previous December).
  long long mm = static_cast<long long>(m);
  long long yy = yi;
  long long mm0 = mm - 1;  // 0-based month in "signed" space
  long long year_shift = mm0 / 12;
  long long rem = mm0 % 12;
  if (rem < 0) {
    rem += 12;
    year_shift -= 1;
  }
  yy += year_shift;
  mm = rem + 1;
  // Excel rejects DATE whose post-month-shift year falls outside [1900, 9999]
  // even if a negative `day` argument would rewind it back into range.
  // `DATE(9999, 13, -1)` nominally resolves to 9999-12-30 after day rewind,
  // but the intermediate month step pushed the year to 10000, so Excel
  // returns #NUM! before day normalisation runs.
  if (yy < 1900 || yy > 9999) {
    return Value::error(ErrorCode::Num);
  }
  const double serial = date_time::serial_from_ymd(static_cast<int>(yy), static_cast<unsigned>(mm),
                                                   static_cast<unsigned>(static_cast<long long>(d)));
  if (serial < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(serial);
}

/// TIME(hour, minute, second). Each component is truncated; out-of-range
/// inputs are accepted and normalised modulo 24/60/60. A negative result
/// (e.g. `TIME(-1, 0, 0)`) yields `#NUM!` per Excel.
Value Time_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto h_c = coerce_to_number(args[0]);
  if (!h_c) {
    return Value::error(h_c.error());
  }
  auto m_c = coerce_to_number(args[1]);
  if (!m_c) {
    return Value::error(m_c.error());
  }
  auto s_c = coerce_to_number(args[2]);
  if (!s_c) {
    return Value::error(s_c.error());
  }
  const double total = std::trunc(h_c.value()) * 3600.0 + std::trunc(m_c.value()) * 60.0 + std::trunc(s_c.value());
  if (total < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  // Normalise modulo a full day so `TIME(25, 0, 0) == TIME(1, 0, 0)`.
  const double day_fraction = std::fmod(total / 86400.0, 1.0);
  return Value::number(day_fraction);
}

/// YEAR(serial). Returns the 4-digit Gregorian year. Extracts the date
/// portion (fractional part discarded). Serial 0 is Excel's fictitious
/// "1/0/1900": YEAR=1900, MONTH=1, DAY=0.
Value Year_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
  }
  if (std::floor(serial.value()) == 0.0) {
    return Value::number(1900.0);
  }
  const date_time::YMD ymd = date_time::ymd_from_serial(serial.value());
  return Value::number(static_cast<double>(ymd.y));
}

/// MONTH(serial). Returns 1..12.
Value Month_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
  }
  if (std::floor(serial.value()) == 0.0) {
    return Value::number(1.0);
  }
  const date_time::YMD ymd = date_time::ymd_from_serial(serial.value());
  return Value::number(static_cast<double>(ymd.m));
}

/// DAY(serial). Returns 0..31 (0 only for the fictitious serial 0 alias).
Value Day_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
  }
  if (std::floor(serial.value()) == 0.0) {
    return Value::number(0.0);
  }
  const date_time::YMD ymd = date_time::ymd_from_serial(serial.value());
  return Value::number(static_cast<double>(ymd.d));
}

/// HOUR(serial). Returns 0..23.
Value Hour_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
  }
  const date_time::HMS hms = date_time::hms_from_fraction(serial.value());
  return Value::number(static_cast<double>(hms.h));
}

/// MINUTE(serial). Returns 0..59.
Value Minute_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
  }
  const date_time::HMS hms = date_time::hms_from_fraction(serial.value());
  return Value::number(static_cast<double>(hms.m));
}

/// SECOND(serial). Returns 0..59.
Value Second_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
  }
  const date_time::HMS hms = date_time::hms_from_fraction(serial.value());
  return Value::number(static_cast<double>(hms.s));
}

/// WEEKDAY(serial, [return_type]). `return_type` selects between Excel's
/// three original encodings (1, 2, 3) and the 2010+ additions (11..17),
/// which simply rotate the week start. Anything else is `#NUM!`.
Value Weekday_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
  }
  int return_type = 1;
  if (arity >= 2) {
    auto rt = coerce_to_number(args[1]);
    if (!rt) {
      return Value::error(rt.error());
    }
    const double t = std::trunc(rt.value());
    return_type = static_cast<int>(t);
  }
  const date_time::YMD ymd = date_time::ymd_from_serial(serial.value());
  // `days_from_civil` at epoch 1970-01-01 (a Thursday). A Thursday has
  // weekday index 4 in the Sunday=0..Saturday=6 convention, so
  // `(days + 4) mod 7` recovers the 0..6 weekday where 0 = Sunday.
  const std::int64_t days = date_time::days_from_civil(ymd.y, ymd.m, ymd.d);
  const int sun0 = static_cast<int>(((days + 4) % 7 + 7) % 7);  // 0..6, Sun=0
  const int mon0 = (sun0 + 6) % 7;                              // 0..6, Mon=0
  // Excel types 11..17 start the week on Mon..Sun respectively and always
  // return 1..7.
  if (return_type >= 11 && return_type <= 17) {
    const int start_offset = return_type - 11;  // 0 = Mon, ..., 6 = Sun
    const int shifted = (mon0 - start_offset + 7) % 7;
    return Value::number(static_cast<double>(shifted + 1));
  }
  switch (return_type) {
    case 1:
      return Value::number(static_cast<double>(sun0 + 1));  // Sun=1..Sat=7
    case 2:
      return Value::number(static_cast<double>(mon0 + 1));  // Mon=1..Sun=7
    case 3:
      return Value::number(static_cast<double>(mon0));  // Mon=0..Sun=6
    default:
      return Value::error(ErrorCode::Num);
  }
}

// EDATE / EOMONTH share the same month-shift preamble: parse the base
// serial, the month delta, then compute the target (year, month) pair
// with the day clamped to the new month's length.
struct ShiftedMonth {
  int y;
  unsigned m;
  unsigned d;      // EDATE result (clamped original day)
  unsigned eom_d;  // EOMONTH result (last day of target month)
};

Expected<ShiftedMonth, ErrorCode> shift_months(const Value* args) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return serial.error();
  }
  auto months_c = coerce_to_number(args[1]);
  if (!months_c) {
    return months_c.error();
  }
  const long long months = static_cast<long long>(std::trunc(months_c.value()));
  const date_time::YMD base = date_time::ymd_from_serial(serial.value());
  long long total_months = static_cast<long long>(base.y) * 12 + static_cast<long long>(base.m - 1) + months;
  // Python-style floor-division so negative deltas work correctly.
  long long new_y = total_months / 12;
  long long new_m = total_months % 12;
  if (new_m < 0) {
    new_m += 12;
    new_y -= 1;
  }
  ShiftedMonth out;
  out.y = static_cast<int>(new_y);
  out.m = static_cast<unsigned>(new_m + 1);
  const unsigned dim = days_in_month(out.y, out.m);
  out.d = base.d > dim ? dim : base.d;
  out.eom_d = dim;
  return out;
}

/// EDATE(start, months). Returns the serial of the same day-of-month shifted
/// by `months` calendar months; day is clamped to the new month's length.
Value Edate_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto shifted = shift_months(args);
  if (!shifted) {
    return Value::error(shifted.error());
  }
  const double serial = date_time::serial_from_ymd(shifted.value().y, shifted.value().m, shifted.value().d);
  if (serial < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(serial);
}

/// EOMONTH(start, months). Returns the serial of the last day of the
/// target month.
Value Eomonth_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  // Excel 365 rejects Boolean arguments to EOMONTH (unlike EDATE, which
  // coerces them). Guard both the start-date and month-offset arguments.
  if (args[0].is_boolean() || args[1].is_boolean()) {
    return Value::error(ErrorCode::Value);
  }
  auto shifted = shift_months(args);
  if (!shifted) {
    return Value::error(shifted.error());
  }
  const double serial = date_time::serial_from_ymd(shifted.value().y, shifted.value().m, shifted.value().eom_d);
  if (serial < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(serial);
}

// ---------------------------------------------------------------------------
// Week / year-fraction / date-diff helpers shared by WEEKNUM, ISOWEEKNUM,
// YEARFRAC and DATEDIF.
// ---------------------------------------------------------------------------

// Returns the Sun=0..Sat=6 weekday index of the Gregorian date (y, m, d).
// Uses the `(days_from_civil + 4) mod 7` trick anchored at the 1970-01-01
// Thursday epoch; the `+ 7) % 7` guard protects against negative remainders
// for pre-1970 proleptic dates (not exercised by Excel, but cheap insurance).
int weekday_sun0(int y, unsigned m, unsigned d) noexcept {
  const std::int64_t days = date_time::days_from_civil(y, m, d);
  return static_cast<int>(((days % 7) + 4 + 7) % 7);
}

// Computes the ISO 8601 week number for an arbitrary Gregorian date.
// Implements the standard "Thursday of the current week" algorithm:
// shift to the Thursday (Mon=1..Sun=7 -> +3 / +2 / ... / -3), then the
// ISO week number is `(day-of-year(Thursday) - 1) / 7 + 1`, and the ISO
// year is that Thursday's calendar year.
int iso_week_number(int y, unsigned m, unsigned d) noexcept {
  const std::int64_t days = date_time::days_from_civil(y, m, d);
  // ISO weekday: Mon=0..Sun=6. The 1970-01-01 epoch falls on a Thursday
  // (days=0 -> Thursday -> mon0=3), so `(days + 3) mod 7` recovers the
  // zero-based Monday-first index. The `+ 7) % 7` guard keeps the
  // result non-negative for proleptic pre-1970 inputs.
  const int mon0 = static_cast<int>(((days + 3) % 7 + 7) % 7);
  // Thursday of this ISO week is `3 - mon0` days away (-3 .. +3).
  const std::int64_t thu_days = days + (3 - mon0);
  const date_time::YMD thu = date_time::civil_from_days(thu_days);
  const std::int64_t jan1_days = date_time::days_from_civil(thu.y, 1, 1);
  const int doy = static_cast<int>(thu_days - jan1_days) + 1;  // 1-based
  return (doy - 1) / 7 + 1;
}

// Completed calendar months between the (y1,m1,d1) -> (y2,m2,d2) pair,
// assuming end >= start (caller enforces). Excel counts a month as
// "completed" iff d2 >= d1.
long long completed_months(int y1, unsigned m1, unsigned d1, int y2, unsigned m2, unsigned d2) noexcept {
  long long months = (static_cast<long long>(y2) - y1) * 12 + (static_cast<long long>(m2) - m1);
  if (d2 < d1) {
    months -= 1;
  }
  return months;
}

/// WEEKNUM(serial, [return_type]). Computes the calendar week number
/// using one of Excel's ten supported return-type codes. Codes 1/2 are
/// the classic Sunday- and Monday-start forms; 11..17 are the "new"
/// codes that rotate the week start across all seven days; 21 is ISO
/// 8601 (delegated to ISOWEEKNUM). Any other type yields `#NUM!`.
///
/// Algorithm for the non-ISO variants: let `ws` be the day-of-week the
/// week starts on (Sun=0..Sat=6). The week number of any date d is
/// `floor((doy(d) + (jan1_dow - ws + 7) % 7 - 1) / 7) + 1`, where
/// `doy` is 1-based day-of-year and `jan1_dow` is the weekday of Jan 1.
/// This is equivalent to the "roll Jan 1 to its week start, roll d to
/// its week start, divide by 7" formulation but avoids two extra
/// conversions through serial space.
Value Weeknum_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
  }
  int return_type = 1;
  if (arity >= 2) {
    auto rt = coerce_to_number(args[1]);
    if (!rt) {
      return Value::error(rt.error());
    }
    return_type = static_cast<int>(std::trunc(rt.value()));
  }
  // Map return_type -> week-start weekday (Sun=0..Sat=6). -1 means invalid.
  int ws = -1;
  switch (return_type) {
    case 1:
    case 17:
      ws = 0;  // Sun
      break;
    case 2:
    case 11:
      ws = 1;  // Mon
      break;
    case 12:
      ws = 2;  // Tue
      break;
    case 13:
      ws = 3;  // Wed
      break;
    case 14:
      ws = 4;  // Thu
      break;
    case 15:
      ws = 5;  // Fri
      break;
    case 16:
      ws = 6;  // Sat
      break;
    case 21: {
      // ISO 8601.
      const date_time::YMD ymd = date_time::ymd_from_serial(std::floor(serial.value()));
      return Value::number(static_cast<double>(iso_week_number(ymd.y, ymd.m, ymd.d)));
    }
    default:
      return Value::error(ErrorCode::Num);
  }
  const date_time::YMD ymd = date_time::ymd_from_serial(std::floor(serial.value()));
  const std::int64_t jan1_days = date_time::days_from_civil(ymd.y, 1, 1);
  const std::int64_t d_days = date_time::days_from_civil(ymd.y, ymd.m, ymd.d);
  const int jan1_sun0 = weekday_sun0(ymd.y, 1, 1);
  // Offset of Jan 1 from its week start (0..6). Adding this to the zero-
  // based day-of-year normalises Jan 1 so the first day of week 1 is day 0.
  const int jan1_offset = (jan1_sun0 - ws + 7) % 7;
  const int doy0 = static_cast<int>(d_days - jan1_days);  // 0-based day-of-year
  const int week = (doy0 + jan1_offset) / 7 + 1;
  return Value::number(static_cast<double>(week));
}

/// ISOWEEKNUM(serial). Returns 1..53 per ISO 8601.
Value Isoweeknum_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
  }
  const date_time::YMD ymd = date_time::ymd_from_serial(std::floor(serial.value()));
  return Value::number(static_cast<double>(iso_week_number(ymd.y, ymd.m, ymd.d)));
}

// ---------------------------------------------------------------------------
// YEARFRAC day-count helpers live in `eval/date_time.h` so the financial
// rate builtins (DISC / INTRATE / RECEIVED / ...) can share them without
// pulling in the rest of the datetime module.
// ---------------------------------------------------------------------------

/// YEARFRAC(start, end, [basis]). Returns the positive (sign-stripped)
/// fraction of a year between two dates under one of five day-count
/// conventions. Invalid basis codes yield `#NUM!`.
Value Yearfrac_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto start = coerce_serial(args[0]);
  if (!start) {
    return Value::error(start.error());
  }
  auto end = coerce_serial(args[1]);
  if (!end) {
    return Value::error(end.error());
  }
  int basis = 0;
  if (arity >= 3) {
    auto b = coerce_to_number(args[2]);
    if (!b) {
      return Value::error(b.error());
    }
    basis = static_cast<int>(std::trunc(b.value()));
  }
  double s = std::floor(start.value());
  double e = std::floor(end.value());
  if (s > e) {
    const double tmp = s;
    s = e;
    e = tmp;
  }
  const date_time::YMD a = date_time::ymd_from_serial(s);
  const date_time::YMD b = date_time::ymd_from_serial(e);
  switch (basis) {
    case 0:
      return Value::number(date_time::yearfrac_us30_360(a.y, a.m, a.d, b.y, b.m, b.d));
    case 1:
      return Value::number(date_time::yearfrac_actual_actual(a.y, a.m, a.d, b.y, b.m, b.d));
    case 2:
      // Actual/360 -- raw day diff honours the 1900 leap-bug so we compute
      // it from the Excel serials (not days_from_civil) for parity.
      return Value::number((e - s) / 360.0);
    case 3:
      return Value::number((e - s) / 365.0);
    case 4:
      return Value::number(date_time::yearfrac_eu30_360(a.y, a.m, a.d, b.y, b.m, b.d));
    default:
      return Value::error(ErrorCode::Num);
  }
}

/// DATEDIF(start, end, unit). Returns the signed difference between two
/// dates expressed in the requested unit. Unit strings are case-sensitive
/// and uppercase-only in Excel's documented surface:
///
///   "Y"  - complete years between start and end
///   "M"  - complete months between start and end
///   "D"  - day difference (`floor(end) - floor(start)`)
///   "YM" - complete months between, ignoring the year part
///   "YD" - days between, treating start as if it fell in end's year
///   "MD" - days between, ignoring months and years (known buggy in Excel)
///
/// `end < start` yields `#NUM!`; an unknown unit also yields `#NUM!`.
Value Datedif_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto start = coerce_serial(args[0]);
  if (!start) {
    return Value::error(start.error());
  }
  auto end = coerce_serial(args[1]);
  if (!end) {
    return Value::error(end.error());
  }
  auto unit = coerce_to_text(args[2]);
  if (!unit) {
    return Value::error(unit.error());
  }
  const double s = std::floor(start.value());
  const double e = std::floor(end.value());
  if (e < s) {
    return Value::error(ErrorCode::Num);
  }
  const date_time::YMD a = date_time::ymd_from_serial(s);
  const date_time::YMD b = date_time::ymd_from_serial(e);
  const std::string& u = unit.value();
  if (u == "D") {
    return Value::number(e - s);
  }
  if (u == "Y") {
    const long long months = completed_months(a.y, a.m, a.d, b.y, b.m, b.d);
    return Value::number(static_cast<double>(months / 12));
  }
  if (u == "M") {
    return Value::number(static_cast<double>(completed_months(a.y, a.m, a.d, b.y, b.m, b.d)));
  }
  if (u == "YM") {
    const long long months = completed_months(a.y, a.m, a.d, b.y, b.m, b.d);
    return Value::number(static_cast<double>(months % 12));
  }
  if (u == "YD") {
    // Shift `start` into the same or preceding year as `end` so the
    // calendar anniversary logic falls out directly. If start's (m,d)
    // has already passed in end's year, the anniversary for the "partial
    // year" is in end's year; otherwise in the year before.
    int anniversary_y = b.y;
    // Calendar comparison (m,d) of start vs end.
    if (a.m > b.m || (a.m == b.m && a.d > b.d)) {
      anniversary_y = b.y - 1;
    }
    const std::int64_t anniv = date_time::days_from_civil(anniversary_y, a.m, a.d);
    const std::int64_t end_days = date_time::days_from_civil(b.y, b.m, b.d);
    return Value::number(static_cast<double>(end_days - anniv));
  }
  if (u == "MD") {
    // Days ignoring months and years: pretend both dates share end's
    // (year, month) and return `b.d - a.d`; if that would be negative,
    // borrow a month (one full calendar month of end's previous month).
    // This mirrors Excel's classic (and buggy) formulation rather than
    // a "correct" calendar-difference algorithm.
    long long diff = static_cast<long long>(b.d) - static_cast<long long>(a.d);
    if (diff < 0) {
      // Borrow the previous month's length relative to the END date.
      int prev_y = b.y;
      int prev_m = static_cast<int>(b.m) - 1;
      if (prev_m == 0) {
        prev_m = 12;
        prev_y -= 1;
      }
      diff += static_cast<long long>(days_in_month(prev_y, static_cast<unsigned>(prev_m)));
    }
    return Value::number(static_cast<double>(diff));
  }
  return Value::error(ErrorCode::Num);
}

/// DAYS(end, start). Both arguments are truncated to their date component;
/// the result is the signed day difference and may be negative.
Value Days_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto end = coerce_to_number(args[0]);
  if (!end) {
    return Value::error(end.error());
  }
  auto start = coerce_to_number(args[1]);
  if (!start) {
    return Value::error(start.error());
  }
  const double diff = std::floor(end.value()) - std::floor(start.value());
  return Value::number(diff);
}

/// DAYS360(start, end, [method]). Counts days under a 360-day-year calendar.
/// `method` is a boolean-coercible selector: FALSE (default) selects the
/// US/NASD rule set, TRUE selects the European rule set. The result is an
/// integer (possibly negative) computed as
/// `360*(ey - sy) + 30*(em - sm) + (ed - sd)` after the day components have
/// been adjusted per the selected rule set.
///
/// US/NASD adjustments (order matters: conditions use the pre-adjusted start
/// day):
///   1. end day 31 and start day < 30  -> end becomes 1, end month += 1
///   2. end day 31 and start day >= 30 -> end day becomes 30
///   3. start day 31                    -> start day becomes 30
///
/// European adjustments: any day of 31 is capped at 30 on both sides.
Value Days360_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto start_n = coerce_to_number(args[0]);
  if (!start_n) {
    return Value::error(start_n.error());
  }
  auto end_n = coerce_to_number(args[1]);
  if (!end_n) {
    return Value::error(end_n.error());
  }
  bool european = false;
  if (arity >= 3) {
    auto m = coerce_to_bool(args[2]);
    if (!m) {
      return Value::error(m.error());
    }
    european = m.value();
  }
  const double start_f = std::floor(start_n.value());
  const double end_f = std::floor(end_n.value());
  // `ymd_from_serial`'s contract requires a non-negative input; reject
  // pre-1900 serials rather than triggering undefined behaviour downstream.
  if (start_f < 0.0 || end_f < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const date_time::YMD s = date_time::ymd_from_serial(start_f);
  const date_time::YMD e = date_time::ymd_from_serial(end_f);
  int sy = s.y;
  int sm = static_cast<int>(s.m);
  int sd = static_cast<int>(s.d);
  int ey = e.y;
  int em = static_cast<int>(e.m);
  int ed = static_cast<int>(e.d);
  if (european) {
    if (sd == 31) {
      sd = 30;
    }
    if (ed == 31) {
      ed = 30;
    }
  } else {
    const bool start_is_last = (sd == 31);
    if (ed == 31) {
      if (sd < 30) {
        ed = 1;
        em += 1;
        if (em > 12) {
          em = 1;
          ey += 1;
        }
      } else {
        ed = 30;
      }
    }
    if (start_is_last) {
      sd = 30;
    }
  }
  const long long result = 360LL * static_cast<long long>(ey - sy) + 30LL * static_cast<long long>(em - sm) +
                           static_cast<long long>(ed - sd);
  return Value::number(static_cast<double>(result));
}

// ---------------------------------------------------------------------------
// NOW / TODAY
//
// Both read the host wall clock. Excel returns local time (the worksheet is
// locale-bound), so we convert `system_clock::now()` via `localtime_r` /
// `localtime_s` to decompose the timestamp into a calendar date + time of
// day, then build the Excel serial from those fields. NOW includes the
// fractional time-of-day; TODAY discards it. Both are marked volatile in
// Excel; the engine's recalc scheduling is out of scope for this impl — every
// invocation re-reads the clock.
// ---------------------------------------------------------------------------

// Returns (date_serial, tod_fraction) for the current local time. Called
// independently by NOW and TODAY so neither routes through a dummy-arena
// call. Defined here rather than in `date_time.cpp` because the clock read
// is only needed by these two builtins and we want to keep `date_time.h`
// clock-agnostic.
struct LocalNow {
  double date_serial;
  double tod_fraction;
};

LocalNow read_local_now() noexcept {
  using std::chrono::system_clock;
  const std::time_t tt = system_clock::to_time_t(system_clock::now());
  std::tm local{};
#if defined(_WIN32)
  localtime_s(&local, &tt);
#else
  localtime_r(&tt, &local);
#endif
  const int y = local.tm_year + 1900;
  const unsigned m = static_cast<unsigned>(local.tm_mon + 1);
  const unsigned d = static_cast<unsigned>(local.tm_mday);
  const double date_serial = date_time::serial_from_ymd(y, m, d);
  const double tod = (static_cast<double>(local.tm_hour) * 3600.0 + static_cast<double>(local.tm_min) * 60.0 +
                      static_cast<double>(local.tm_sec)) /
                     86400.0;
  return LocalNow{date_serial, tod};
}

/// NOW(). Returns the current local date-time as an Excel serial
/// (integer = days since the Excel epoch; fraction = time-of-day / 86400).
/// Zero-argument; `localtime_r` / `localtime_s` decompose the wall clock.
Value Now_(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const LocalNow now = read_local_now();
  return Value::number(now.date_serial + now.tod_fraction);
}

/// TODAY(). Returns the current local date as an integer Excel serial.
/// Equivalent to `std::floor(NOW())` but reads the clock once to avoid any
/// second-boundary drift between the two halves.
Value Today_(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const LocalNow now = read_local_now();
  return Value::number(now.date_serial);
}

// ---------------------------------------------------------------------------
// DATEVALUE / TIMEVALUE text parsing
//
// The parse grammar (ISO dashed, slash, and kanji date forms; time-of-day
// with optional fractional seconds and AM/PM markers) is implemented in
// `src/eval/date_text_parse.{h,cpp}` so DATEVALUE / TIMEVALUE / VALUE all
// funnel through the same recognizer. The helpers below only add the
// Excel-level argument-shape rules on top of that shared parser.
// ---------------------------------------------------------------------------

/// DATEVALUE(text). Parses `text` as a date string and returns the Excel
/// serial for the date part (integer). Embedded time components are
/// accepted and ignored — a string like "2024-03-15 13:30" returns the
/// serial for 2024-03-15. Booleans / numbers coerce to text first, so
/// `DATEVALUE(TRUE)` tries to parse the literal "TRUE" and fails with
/// `#VALUE!`.
Value Datevalue_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string_view trimmed = date_parse::trim_date_text(text.value());
  if (trimmed.empty()) {
    return Value::error(ErrorCode::Value);
  }
  double serial = 0.0;
  double frac = 0.0;
  bool has_date = false;
  bool has_time = false;
  if (!date_parse::parse_date_time_text(trimmed, &serial, &frac, &has_date, &has_time)) {
    return Value::error(ErrorCode::Value);
  }
  if (!has_date) {
    // DATEVALUE requires an explicit date component; time-only input is
    // rejected as #VALUE! (Excel defaults to "today", which isn't
    // reproducible without a clock — see the divergence note in the file
    // banner comment).
    return Value::error(ErrorCode::Value);
  }
  return Value::number(serial);
}

/// TIMEVALUE(text). Parses `text` as a time string and returns the
/// fractional-day value. Embedded date components are accepted and ignored —
/// a string like "2024-03-15 13:30" returns 0.5625. Hour values exceeding 24
/// wrap around the day (`TIMEVALUE("25:00")` returns `1/24` after modular
/// reduction).
Value Timevalue_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string_view trimmed = date_parse::trim_date_text(text.value());
  if (trimmed.empty()) {
    return Value::error(ErrorCode::Value);
  }
  double serial = 0.0;
  double frac = 0.0;
  bool has_date = false;
  bool has_time = false;
  if (!date_parse::parse_date_time_text(trimmed, &serial, &frac, &has_date, &has_time)) {
    return Value::error(ErrorCode::Value);
  }
  if (!has_time) {
    // TIMEVALUE requires an explicit time component; a date-only input
    // returns `#VALUE!` (matches Excel's behaviour for "2024-03-15" as an
    // argument to TIMEVALUE).
    return Value::error(ErrorCode::Value);
  }
  // Normalise the fractional day modulo a full day so TIMEVALUE("25:00")
  // returns the same value as TIMEVALUE("1:00") = 1/24.
  double out = std::fmod(frac, 1.0);
  if (out < 0.0) {
    out += 1.0;
  }
  return Value::number(out);
}

}  // namespace

void register_datetime_builtins(FunctionRegistry& registry) {
  // Date / time. All eager, scalar-only (no range expansion). NOW / TODAY
  // read the host's local wall clock via `std::chrono::system_clock` and
  // `localtime_r` (Windows: `localtime_s`). Recalc semantics (marking them
  // volatile) are handled by the scheduler, not this registration.
  registry.register_function(FunctionDef{"DATE", 3u, 3u, &Date_});
  registry.register_function(FunctionDef{"TIME", 3u, 3u, &Time_});
  registry.register_function(FunctionDef{"YEAR", 1u, 1u, &Year_});
  registry.register_function(FunctionDef{"MONTH", 1u, 1u, &Month_});
  registry.register_function(FunctionDef{"DAY", 1u, 1u, &Day_});
  registry.register_function(FunctionDef{"HOUR", 1u, 1u, &Hour_});
  registry.register_function(FunctionDef{"MINUTE", 1u, 1u, &Minute_});
  registry.register_function(FunctionDef{"SECOND", 1u, 1u, &Second_});
  registry.register_function(FunctionDef{"WEEKDAY", 1u, 2u, &Weekday_});
  registry.register_function(FunctionDef{"EDATE", 2u, 2u, &Edate_});
  registry.register_function(FunctionDef{"EOMONTH", 2u, 2u, &Eomonth_});
  registry.register_function(FunctionDef{"DAYS", 2u, 2u, &Days_});
  registry.register_function(FunctionDef{"DAYS360", 2u, 3u, &Days360_});
  registry.register_function(FunctionDef{"WEEKNUM", 1u, 2u, &Weeknum_});
  registry.register_function(FunctionDef{"ISOWEEKNUM", 1u, 1u, &Isoweeknum_});
  registry.register_function(FunctionDef{"YEARFRAC", 2u, 3u, &Yearfrac_});
  registry.register_function(FunctionDef{"DATEDIF", 3u, 3u, &Datedif_});
  registry.register_function(FunctionDef{"DATEVALUE", 1u, 1u, &Datevalue_});
  registry.register_function(FunctionDef{"TIMEVALUE", 1u, 1u, &Timevalue_});
  registry.register_function(FunctionDef{"NOW", 0u, 0u, &Now_});
  registry.register_function(FunctionDef{"TODAY", 0u, 0u, &Today_});
}

}  // namespace eval
}  // namespace formulon
