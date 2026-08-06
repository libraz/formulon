//
// Implementation of the shared shortest-form `double` formatter. See
// `double_format.h` for the behavioural contract.

#include "utils/double_format.h"

#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <string>

namespace formulon {

namespace {

// Maximum decimal-form display width for Excel's General format. Probes
// against Mac Excel 365 (ja-JP, build 16.108) show that VALUETOTEXT
// preserves decimal form as long as the unsigned magnitude string is at
// most 20 characters; beyond that it switches to scientific notation.
// The leading minus sign (when present) is NOT counted toward this limit.
//
// Examples:
//   1E-18 -> "0.000000000000000001"  (20 chars,  decimal)
//   1E-19 -> "1E-19"                  (sci, decimal would be 21 chars)
//   1E+19 -> "10000000000000000000"   (20 chars,  decimal)
//   1E+20 -> "1E+20"                  (sci, decimal would be 21 chars)
constexpr std::size_t kGeneralDecimalMaxChars = 20;

// Append `n` ASCII '0' characters to `out`.
void append_zeros(std::string& out, int n) {
  for (int i = 0; i < n; ++i) {
    out.push_back('0');
  }
}

// Renders the Excel-style exponent suffix for scientific notation:
//   "E+NN" / "E-NN"
// where the exponent magnitude is zero-padded to at least two digits and
// expands naturally to three or more digits for |exp| >= 100.
void append_excel_exponent(std::string& out, int exp) {
  out.push_back('E');
  out.push_back(exp >= 0 ? '+' : '-');
  int abs_exp = exp < 0 ? -exp : exp;
  if (abs_exp < 10) {
    out.push_back('0');
    out.push_back(static_cast<char>('0' + abs_exp));
  } else {
    char tmp[16];
    std::snprintf(tmp, sizeof(tmp), "%d", abs_exp);
    out.append(tmp);
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
  // Negative zero collapses to plain "0" so callers never see a stray sign.
  if (v == 0.0) {
    out.push_back('0');
    return;
  }

  const bool negative = v < 0.0;
  const double abs_v = negative ? -v : v;

  // Excel's General format renders to at most 15 significant digits;
  // %.15g performs IEEE round-to-nearest-even on the shortest form,
  // matching Mac Excel for the overwhelming majority of values. (Sub-ULP
  // divergences at 16+ sig digits — e.g. literal `1234567890123456` —
  // are documented in tests/divergence.yaml.)
  char buf[40];
  int n = std::snprintf(buf, sizeof(buf), "%.15g", abs_v);
  if (n <= 0 || static_cast<std::size_t>(n) >= sizeof(buf)) {
    // Defensive fallback: should not happen for finite doubles.
    if (negative) {
      out.push_back('-');
    }
    out.append(buf);
    return;
  }

  // Locate the optional 'e' marker that %g uses for scientific form, and
  // the optional '.' decimal point.
  char* e_pos = nullptr;
  char* dot_pos = nullptr;
  for (char* p = buf; *p != '\0'; ++p) {
    if (*p == 'e' && e_pos == nullptr) {
      e_pos = p;
    } else if (*p == '.' && dot_pos == nullptr) {
      dot_pos = p;
    }
  }

  // Extract the digits (mantissa, no sign, no decimal point) and the
  // 10-exponent of the leading digit, i.e. v = digits[0].digits[1..] * 10^exp.
  std::string digits;
  int leading_exp;
  {
    int g_exp = 0;
    char* mantissa_end = (e_pos != nullptr) ? e_pos : buf + n;
    if (e_pos != nullptr) {
      g_exp = std::atoi(e_pos + 1);
    }
    int int_part_len;
    if (dot_pos != nullptr && dot_pos < mantissa_end) {
      digits.assign(buf, dot_pos);
      digits.append(dot_pos + 1, mantissa_end);
      int_part_len = static_cast<int>(dot_pos - buf);
    } else {
      digits.assign(buf, mantissa_end);
      int_part_len = static_cast<int>(mantissa_end - buf);
    }
    leading_exp = (int_part_len - 1) + g_exp;
  }

  // Normalise: strip leading zeros (adjusting `leading_exp`) and trailing
  // zeros (cosmetic). After this, `digits` has no extraneous zeros.
  while (digits.size() > 1 && digits.front() == '0') {
    digits.erase(0, 1);
    --leading_exp;
  }
  while (digits.size() > 1 && digits.back() == '0') {
    digits.pop_back();
  }

  // Compute the natural decimal form's character count (unsigned).
  // - If leading_exp >= 0:
  //     integer part has (leading_exp + 1) chars.
  //     fractional part exists iff digits.size() > integer_part_len, then
  //     contributes 1 ('.') + (digits.size() - integer_part_len) chars.
  // - If leading_exp < 0:
  //     "0." + (-leading_exp - 1) leading zeros + digits.size() chars.
  std::size_t decimal_len;
  if (leading_exp >= 0) {
    const int int_chars = leading_exp + 1;
    if (static_cast<int>(digits.size()) > int_chars) {
      decimal_len = digits.size() + 1;  // includes '.'
    } else {
      decimal_len = static_cast<std::size_t>(int_chars);
    }
  } else {
    decimal_len = 2u + static_cast<std::size_t>(-leading_exp - 1) + digits.size();
  }

  if (negative) {
    out.push_back('-');
  }

  if (decimal_len <= kGeneralDecimalMaxChars) {
    // Decimal form. Negative sign (when present) is NOT counted against
    // the 20-char ceiling, matching Mac Excel.
    if (leading_exp >= 0) {
      const int int_chars = leading_exp + 1;
      const int digits_len = static_cast<int>(digits.size());
      if (digits_len >= int_chars) {
        out.append(digits, 0, static_cast<std::size_t>(int_chars));
        if (digits_len > int_chars) {
          out.push_back('.');
          out.append(digits, static_cast<std::size_t>(int_chars), std::string::npos);
        }
      } else {
        out.append(digits);
        append_zeros(out, int_chars - digits_len);
      }
    } else {
      out.push_back('0');
      out.push_back('.');
      append_zeros(out, -leading_exp - 1);
      out.append(digits);
    }
  } else {
    // Scientific form. Mantissa = digits[0] '.' digits[1..]; exponent in
    // Excel style ("E+NN" / "E-NN", min 2-digit exponent).
    out.push_back(digits.front());
    if (digits.size() > 1) {
      out.push_back('.');
      out.append(digits, 1, std::string::npos);
    }
    append_excel_exponent(out, leading_exp);
  }
}

}  // namespace formulon
