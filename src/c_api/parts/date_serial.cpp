//
// C ABI - civil time <-> Excel date serial conversion.

#include <cmath>
#include <cstdint>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "utils/date_time.h"

using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;

namespace {

constexpr double kSecondsPerDay = static_cast<double>(formulon::date_time::kSecondsPerDay);
constexpr double kGhostDaySerial = 60.0;

bool in_civil_range(const fm_civil_time_t& c, bool date1904) {
  const bool ghost_day = !date1904 && c.year == 1900 && c.month == 2 && c.day == 29;
  if (c.year < 1900 || c.year > 9999 || c.month < 1 || c.month > 12 || c.day < 1 || c.hour < 0 || c.hour > 23 ||
      c.minute < 0 || c.minute > 59 || c.second < 0 || c.second > 59) {
    return false;
  }
  if (date1904 && c.year < 1904) {
    return false;
  }
  const auto month_days =
      static_cast<std::int32_t>(formulon::date_time::days_in_month(c.year, static_cast<unsigned>(c.month)));
  return ghost_day || c.day <= month_days;
}

}  // namespace

extern "C" fm_status_t fm_convert_civil_to_serial(const fm_civil_time_t* civil, int32_t date1904, double* out_serial) {
  clear_last_error();
  if (civil == nullptr || out_serial == nullptr) {
    return null_argument();
  }
  const bool is_1904 = date1904 != 0;
  if (!in_civil_range(*civil, is_1904)) {
    return invalid_argument("civil field out of range");
  }
  const bool ghost_day = !is_1904 && civil->year == 1900 && civil->month == 2 && civil->day == 29;
  const double day = ghost_day ? kGhostDaySerial
                               : formulon::date_time::serial_from_ymd(civil->year, static_cast<unsigned>(civil->month),
                                                                      static_cast<unsigned>(civil->day), is_1904);
  const double seconds = static_cast<double>(civil->hour) * 3600.0 + static_cast<double>(civil->minute) * 60.0 +
                         static_cast<double>(civil->second);
  *out_serial = day + seconds / kSecondsPerDay;
  return 0;
}

extern "C" fm_status_t fm_convert_serial_to_civil(double serial, int32_t date1904, fm_civil_time_t* out_civil) {
  clear_last_error();
  if (out_civil == nullptr) {
    return null_argument("out_civil");
  }
  const bool is_1904 = date1904 != 0;
  const double max_serial = formulon::date_time::max_date_serial(is_1904);
  if (!std::isfinite(serial) || serial < (is_1904 ? 0.0 : 1.0) || serial >= max_serial + 1.0) {
    return invalid_argument("serial out of range");
  }
  // Round the whole serial to a second as SECOND() does, then split off the day.
  const double total_seconds = std::round(serial * kSecondsPerDay);
  const double day = std::floor(total_seconds / kSecondsPerDay);
  const double second_of_day = total_seconds - day * kSecondsPerDay;
  if (day > max_serial) {
    return invalid_argument("serial rounds past 9999-12-31");
  }
  const formulon::date_time::YMD ymd = formulon::date_time::ymd_from_serial(day, is_1904);
  const auto secs = static_cast<std::int32_t>(second_of_day);
  out_civil->year = ymd.y;
  out_civil->month = static_cast<std::int32_t>(ymd.m);
  out_civil->day = static_cast<std::int32_t>(ymd.d);
  out_civil->hour = secs / 3600;
  out_civil->minute = (secs % 3600) / 60;
  out_civil->second = secs % 60;
  return 0;
}
