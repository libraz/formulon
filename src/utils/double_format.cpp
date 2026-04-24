// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the shared shortest-form `double` formatter. See
// `double_format.h` for the behavioural contract.

#include "utils/double_format.h"

#include <cmath>
#include <cstdint>
#include <cstdio>
#include <string>

namespace formulon {

namespace {

// Uppercases any exponent marker in a %g-formatted buffer in place. Excel's
// General format renders exponents as "1E+16" (uppercase E), while printf
// emits "1e+16" on every platform. Only the single lowercase `e` that sits
// between the mantissa and the exponent sign / digits is affected; any
// digits or signs in the rest of the buffer stay put.
void uppercase_exponent(char* buf) noexcept {
  for (char* p = buf; *p != '\0'; ++p) {
    if (*p == 'e') {
      *p = 'E';
      return;
    }
  }
}

}  // namespace

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

  // Excel's General format rounds to 15 significant digits on display,
  // deliberately hiding IEEE-754 artifacts like 0.1 + 0.2 ("0.3" instead
  // of "0.30000000000000004"). `%.15g` matches that; this is a display
  // / text-coercion formatter, not a roundtrip serializer. Callers that
  // need full-precision roundtrip should route through a separate path.
  char buf[32];
  int n = std::snprintf(buf, sizeof(buf), "%.15g", v);
  if (n <= 0 || static_cast<std::size_t>(n) >= sizeof(buf)) {
    // Extreme corner case (should not happen for finite doubles): fall back
    // to the previous std::to_string path with trailing-zero trimming.
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
    return;
  }
  uppercase_exponent(buf);
  out.append(buf);
}

}  // namespace formulon
