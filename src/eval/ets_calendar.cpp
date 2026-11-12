//
// Calendar arithmetic behind the monthly timelines of the FORECAST.ETS family:
// serial to (year, month, day) conversion and the month-step walk.

#include <cassert>
#include <cmath>
#include <cstdint>
#include <optional>

#include "eval/ets_series.h"
#include "eval/ets_series_internal.h"
#include "utils/date_time.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace ets {
namespace series_detail {

namespace {

// @limit: excel
constexpr std::uint32_t kMaxDateSerial = static_cast<std::uint32_t>(date_time::kMaxSerial1900);

// Days since 1970-01-01 of a proleptic Gregorian date.
std::int64_t days_from_civil(std::int64_t y, std::int64_t m, std::int64_t d) noexcept {
  y -= m <= 2 ? 1 : 0;
  const std::int64_t era = floor_div(y, 400);
  const std::int64_t yoe = y - era * 400;
  const std::int64_t doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1;
  const std::int64_t doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
  return era * 146097 + doe - 719468;
}

void civil_from_days(std::int64_t z, std::int64_t* y, std::int64_t* m, std::int64_t* d) noexcept {
  z += 719468;
  const std::int64_t era = floor_div(z, 146097);
  const std::int64_t doe = z - era * 146097;
  const std::int64_t yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
  const std::int64_t doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
  const std::int64_t mp = (5 * doy + 2) / 153;
  *d = doy - (153 * mp + 2) / 5 + 1;
  *m = mp + (mp < 10 ? 3 : -9);
  *y = yoe + era * 400 + (*m <= 2 ? 1 : 0);
}

// Serial 0 is 1899-12-30.
std::int64_t spreadsheet_epoch_days() noexcept {
  return days_from_civil(1899, 12, 30);
}

// A day past the month end rolls into the next month.
double serial_of(const Ymd& v) {
  return static_cast<double>(days_from_civil(v.year, v.month, 1) - spreadsheet_epoch_days() + (v.day - 1));
}

}  // namespace

std::int64_t floor_div(std::int64_t a, std::int64_t b) noexcept {
  std::int64_t q = a / b;
  if ((a % b != 0) && ((a < 0) != (b < 0))) {
    --q;
  }
  return q;
}

bool is_date_serial(double x) noexcept {
  return sign_of(x) != -1 && to_uint32(x) <= kMaxDateSerial;
}

// `serial` is a grid node or a target already admitted by `whole_count`.
Ymd ymd_of(double serial) {
  const std::optional<std::int64_t> days = whole_count(std::floor(serial));
  assert(days.has_value());
  Ymd r;
  civil_from_days(*days + spreadsheet_epoch_days(), &r.year, &r.month, &r.day);
  return r;
}

// Moves the (year, month) pair in place, keeping the day.
double add_months(Ymd* v, std::int64_t delta) {
  const std::int64_t total = v->year * 12 + (v->month + delta) - 1;
  v->year = floor_div(total, 12);
  v->month = total - v->year * 12 + 1;
  return serial_of(*v);
}

// Nearest month-step node from `a` toward `b` and the signed step count.
bool month_walk(std::int64_t months, double a, double b, double* node, std::int64_t* count, ErrorCode* err) {
  std::int64_t delta = 0;
  std::int64_t sign = 0;
  if (bit_compare(a, b) == -1) {
    delta = months;
    sign = 1;
  } else {
    delta = -months;
    sign = -1;
  }
  if (months <= 0) {
    *err = ErrorCode::Value;
    return false;
  }
  std::int64_t cnt = 0;
  double cur = a;
  Ymd ymd = ymd_of(a);
  double d = std::fabs(a - b);
  while (true) {
    const double next = add_months(&ymd, delta);
    const double d2 = std::fabs(next - b);
    if (bit_compare(d2, d) >= 0) {
      break;
    }
    cur = next;
    cnt += sign;
    d = d2;
  }
  *node = cur;
  *count = cnt;
  return true;
}

}  // namespace series_detail
}  // namespace ets
}  // namespace eval
}  // namespace formulon
