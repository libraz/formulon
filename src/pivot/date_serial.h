// Shared validation for Excel date serials used by pivot grouping and filters.

#ifndef FORMULON_PIVOT_DATE_SERIAL_H_
#define FORMULON_PIVOT_DATE_SERIAL_H_

#include <cmath>

#include "utils/date_time.h"

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

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_DATE_SERIAL_H_
