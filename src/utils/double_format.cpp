// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the shared shortest-form `double` formatter. See
// `double_format.h` for the behavioural contract.

#include "utils/double_format.h"

#include <cmath>
#include <cstdint>
#include <string>

namespace formulon {

void format_double(std::string& out, double v) {
  if (std::isnan(v)) {
    out.append("nan");
    return;
  }
  if (std::isinf(v)) {
    out.append(v < 0.0 ? "-inf" : "inf");
    return;
  }
  // Negative zero collapses to plain "0" so callers do not have to worry
  // about a stray sign byte sneaking into goldens or concat results.
  if (v == 0.0) {
    out.push_back('0');
    return;
  }
  // Integer fast path: any double in (-1e16, 1e16) whose fractional part is
  // exactly zero round-trips to its int64 value, so we can print it without
  // a decimal point.
  if (std::abs(v) < 1e16) {
    const double truncated = std::trunc(v);
    if (truncated == v) {
      const std::int64_t as_int = static_cast<std::int64_t>(truncated);
      out.append(std::to_string(as_int));
      return;
    }
  }
  // Fallback: std::to_string is locale-dependent in principle but produces
  // the C-locale "1.234" form on every libc Formulon targets. Trim trailing
  // zeros after the decimal point, then any stranded trailing dot.
  std::string s = std::to_string(v);
  const auto dot = s.find('.');
  if (dot != std::string::npos) {
    std::size_t last = s.size();
    while (last > dot + 1 && s[last - 1] == '0') {
      --last;
    }
    if (last > 0 && s[last - 1] == '.') {
      --last;
    }
    s.resize(last);
  }
  out.append(s);
}

}  // namespace formulon
