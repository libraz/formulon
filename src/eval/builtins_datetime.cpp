// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's calendar built-in functions: DATE, TIME, YEAR,
// MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY, EDATE, EOMONTH, DAYS.
//
// All entries are eager, scalar-only (no range expansion). NOW / TODAY are
// deferred until the evaluator gains a clock-injection surface; DATEVALUE
// / TIMEVALUE are deferred until locale-aware text parsing lands. Shared
// calendar helpers (serial <-> y/m/d, weekday arithmetic, time-of-day
// decomposition) live in `eval/date_time.h`; this file only layers Excel's
// argument-shape and error-handling rules on top.

#include "eval/builtins_datetime.h"

#include <cmath>
#include <cstdint>

#include "eval/coerce.h"
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
Expected<double, ErrorCode> coerce_serial(const Value& v) {
  auto n = coerce_to_number(v);
  if (!n) {
    return n.error();
  }
  if (n.value() < 0.0) {
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
/// portion (fractional part discarded).
Value Year_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
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
  const date_time::YMD ymd = date_time::ymd_from_serial(serial.value());
  return Value::number(static_cast<double>(ymd.m));
}

/// DAY(serial). Returns 1..31.
Value Day_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto serial = coerce_serial(args[0]);
  if (!serial) {
    return Value::error(serial.error());
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

}  // namespace

void register_datetime_builtins(FunctionRegistry& registry) {
  // Date / time. All eager, scalar-only (no range expansion). NOW / TODAY
  // are deferred until the evaluator gains a clock-injection surface;
  // DATEVALUE / TIMEVALUE are deferred until locale-aware text parsing
  // lands.
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
}

}  // namespace eval
}  // namespace formulon
