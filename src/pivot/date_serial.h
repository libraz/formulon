// Shared validation for Excel date serials used by pivot grouping and filters.

#ifndef FORMULON_PIVOT_DATE_SERIAL_H_
#define FORMULON_PIVOT_DATE_SERIAL_H_

#include <array>
#include <cmath>
#include <string>

#include "utils/date_time.h"
#include "utils/number_text.h"

namespace formulon::pivot {

/// Excel's maximum supported date serial (9999-12-31). Fractions on that day
/// remain valid.
inline double last_valid_date_serial(bool date1904) noexcept {
  return date_time::serial_from_ymd(9999, 12U, 31U, date1904);
}

/// Returns whether `serial` belongs to the finite, non-negative Excel date
/// domain. Callers should convert the serial to a calendar date only after
/// this check succeeds.
inline bool is_valid_date_serial(double serial, bool date1904) noexcept {
  return std::isfinite(serial) && serial >= 0.0 && std::floor(serial) <= last_valid_date_serial(date1904);
}

/// Days before each month in a leap year, with the year length as the last entry: a days group numbers 2/29
/// as day 60 in every year, so 3/1 is day 61 whether or not the year has a 2/29.
inline constexpr std::array<unsigned, 13> kLeapDaysBefore{0U,   31U,  60U,  91U,  121U, 152U, 182U,
                                                          213U, 244U, 274U, 305U, 335U, 366U};

/// Appends `v` to `out`, zero-padded to `min_width` digits.
inline void append_padded(std::string& out, long long v, int min_width) {
  char buf[24];
  format_signed(buf, sizeof buf, v, min_width);
  out += buf;
}

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_DATE_SERIAL_H_
