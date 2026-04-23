// Copyright 2026 libraz. Licensed under the MIT License.
//
// Renderer for the Excel TEXT() format-string engine. The tokenizer lives
// in `number_format_tokenizer.cpp`; the public entry point `apply_format`
// lives in `number_format.cpp`. Shared types are declared in
// `number_format_types.h`.
//
// Numeric rendering pre-computes the integer / fractional digit counts,
// scales the value for `%` and trailing-comma divides, rounds to the
// fractional precision with half-away-from-zero, then walks the tokens
// once interleaving the pre-formatted digits and literal runs.
//
// Date/time rendering converts the serial via the shared `date_time`
// helpers and substitutes each `y/m/d/h/s` token by its textual form.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>

#include "eval/date_time.h"
#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace eval {
namespace number_format_detail {
namespace {

const char* month_short(unsigned m) noexcept {
  // Mac Excel ja-JP surprisingly renders `mmm` in English (Jan/Feb/...).
  // The Japanese `N月` form is reserved for `[DBNum2]` and friends, which
  // are out of scope here.
  static const char* kTable[12] = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
  if (m < 1u || m > 12u) {
    return "";
  }
  return kTable[m - 1u];
}

const char* month_long(unsigned m) noexcept {
  // Matches Mac Excel ja-JP: `mmmm` renders as the English full name.
  static const char* kTable[12] = {"January", "February", "March",     "April",   "May",      "June",
                                   "July",    "August",   "September", "October", "November", "December"};
  if (m < 1u || m > 12u) {
    return "";
  }
  return kTable[m - 1u];
}

const char* weekday_short(int sun0) noexcept {
  // Mac Excel ja-JP `ddd` returns English 3-letter weekday abbreviations.
  static const char* kTable[7] = {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"};
  if (sun0 < 0 || sun0 > 6) {
    return "";
  }
  return kTable[sun0];
}

const char* weekday_long(int sun0) noexcept {
  // Mac Excel ja-JP `dddd` renders the English full weekday name.
  static const char* kTable[7] = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};
  if (sun0 < 0 || sun0 > 6) {
    return "";
  }
  return kTable[sun0];
}

void append_pad2(std::string& out, unsigned n) {
  if (n < 10u) {
    out.push_back('0');
  }
  out.append(std::to_string(n));
}

// Rounds `v` to `decimals` fractional places, half-away-from-zero. The
// result's integer and fractional digit strings are written into
// `*int_digits` and `*frac_digits` (decimal digits only, no sign, no
// leading zeros on the integer side for zero values — except the integer
// is always at least `"0"`).
//
// This helper does not handle scientific notation; see `render_scientific`.
void format_fixed_digits(double v, int decimals, bool* negative, std::string* int_digits, std::string* frac_digits) {
  *negative = std::signbit(v);
  double abs_v = std::fabs(v);
  // Use snprintf with requested precision so we pick up libc rounding.
  // `%.*f` in the C locale rounds ties away from zero on the two libcs
  // Formulon targets (glibc, musl, and Apple's libc); this matches
  // Excel's observable behaviour for TEXT's common cases.
  char buf[64];
  const int n = std::snprintf(buf, sizeof(buf), "%.*f", decimals < 0 ? 0 : decimals, abs_v);
  if (n < 0 || static_cast<std::size_t>(n) >= sizeof(buf)) {
    // Fallback: fall back to sprintf with a heap buffer (extremely rare).
    std::string out;
    out.resize(static_cast<std::size_t>(decimals < 0 ? 0 : decimals) + 32u);
    const int m = std::snprintf(&out[0], out.size(), "%.*f", decimals < 0 ? 0 : decimals, abs_v);
    if (m > 0) {
      out.resize(static_cast<std::size_t>(m));
    } else {
      out = "0";
    }
    const std::size_t dot = out.find('.');
    if (dot == std::string::npos) {
      *int_digits = out;
      frac_digits->clear();
    } else {
      *int_digits = out.substr(0, dot);
      *frac_digits = out.substr(dot + 1);
    }
    return;
  }
  std::string_view s(buf, static_cast<std::size_t>(n));
  const std::size_t dot = s.find('.');
  if (dot == std::string_view::npos) {
    int_digits->assign(s);
    frac_digits->clear();
  } else {
    int_digits->assign(s.substr(0, dot));
    frac_digits->assign(s.substr(dot + 1));
  }
}

// Insert ASCII thousands separators into `int_digits`. Walks right-to-left
// emitting a comma every three digits. `int_digits` is overwritten.
void insert_thousands(std::string& int_digits) {
  if (int_digits.size() <= 3) {
    return;
  }
  std::string out;
  out.reserve(int_digits.size() + int_digits.size() / 3);
  const std::size_t n = int_digits.size();
  const std::size_t rem = n % 3;
  std::size_t k = 0;
  if (rem > 0) {
    out.append(int_digits, 0, rem);
    k = rem;
  }
  while (k < n) {
    if (k != 0) {
      out.push_back(',');
    }
    out.append(int_digits, k, 3);
    k += 3;
  }
  int_digits = std::move(out);
}

}  // namespace

// Render one numeric section through the walk-tokens pipeline.
//
// Steps:
//   1. Apply scaling (percent, trailing-comma divisions).
//   2. Compute fractional-digit precision = max(zero+pad, opt) ... in Excel
//      practice, the precision equals the total count of digit tokens in
//      the fractional part. Excel uses `zero + opt + pad` for the rounded
//      precision.
//   3. Round the absolute value to that precision; capture sign separately.
//   4. Walk the tokens and weave the integer/fraction digit strings into
//      the output alongside literals, commas, and percent.
void render_numeric(const Section& section, std::string_view fmt, double value, std::string& out) {
  double scaled = value;
  if (section.has_percent) {
    scaled *= 100.0;
  }
  for (int i = 0; i < section.trailing_comma_scale; ++i) {
    scaled /= 1000.0;
  }
  const int frac_digits = section.fraction_zero_digits + section.fraction_opt_digits + section.fraction_pad_digits;
  bool negative = false;
  std::string int_digits;
  std::string frac_digits_str;

  if (section.has_scientific) {
    // Scientific notation: compute mantissa/exponent, format mantissa
    // through the fixed path, then append the exponent with the configured
    // sign behaviour.
    double mantissa = scaled;
    int exponent = 0;
    if (mantissa != 0.0) {
      // Normalise mantissa so the integer part has `integer_digits_total`
      // digits; default is one digit.
      const int int_total = section.integer_zero_digits + section.integer_opt_digits + section.integer_pad_digits;
      const int want_int_digits = int_total > 0 ? int_total : 1;
      const double abs_m = std::fabs(mantissa);
      const double log10_m = std::log10(abs_m);
      // Target: floor(log10(|mantissa|)) == want_int_digits - 1.
      const int target = want_int_digits - 1;
      int shift = static_cast<int>(std::floor(log10_m)) - target;
      // Walk toward target in a stable way (avoid pow() rounding drift for
      // exact powers of 10).
      if (shift > 0) {
        for (int i = 0; i < shift; ++i) {
          mantissa /= 10.0;
        }
      } else if (shift < 0) {
        for (int i = 0; i < -shift; ++i) {
          mantissa *= 10.0;
        }
      }
      exponent = shift;
    }
    format_fixed_digits(mantissa, frac_digits, &negative, &int_digits, &frac_digits_str);
    // Fall through into the standard walker below, but also remember the
    // exponent to emit when we hit the scientific marker.
    // We lay out a specialised walker here rather than reusing the fixed
    // path's literal copier because the scientific token needs to consume
    // both the SciPlus/SciMinus token and the subsequent digit tokens that
    // describe the exponent width.
    std::string result;
    if (negative) {
      result.push_back('-');
    }
    // int_digits is already normalised by the scaling + format_fixed_digits
    // call above; append it verbatim.
    result.append(int_digits);
    if (section.has_point && frac_digits > 0) {
      result.push_back('.');
      // Pad `frac_digits_str` to `frac_digits` characters.
      if (static_cast<int>(frac_digits_str.size()) < frac_digits) {
        frac_digits_str.append(static_cast<std::size_t>(frac_digits) - frac_digits_str.size(), '0');
      } else if (static_cast<int>(frac_digits_str.size()) > frac_digits) {
        frac_digits_str.resize(static_cast<std::size_t>(frac_digits));
      }
      result.append(frac_digits_str);
    }
    // Exponent marker + digits. Sign emission: SciPlus always emits '+' or
    // '-'; SciMinus emits only '-'.
    // Use capital or lowercase 'E' matching the format string? In this
    // simplified implementation we always emit 'E'. (Mac Excel on a ja-JP
    // keyboard renders a capital E by default.)
    result.push_back('E');
    if (exponent >= 0) {
      if (section.sci_plus) {
        result.push_back('+');
      }
    } else {
      result.push_back('-');
    }
    const int abs_exp = exponent >= 0 ? exponent : -exponent;
    std::string exp_str = std::to_string(abs_exp);
    // Pad to at least `sci_digits` digits.
    if (static_cast<int>(exp_str.size()) < section.sci_digits) {
      exp_str.insert(exp_str.begin(), static_cast<std::size_t>(section.sci_digits) - exp_str.size(), '0');
    }
    result.append(exp_str);
    out.append(result);
    (void)fmt;
    return;
  }

  format_fixed_digits(scaled, frac_digits, &negative, &int_digits, &frac_digits_str);

  // Pad integer digits to the required minimum (zero + pad digits). Leading
  // `#` tokens above the actual digit count drop silently.
  const int int_min = section.integer_zero_digits + section.integer_pad_digits;
  if (static_cast<int>(int_digits.size()) < int_min) {
    int_digits.insert(0, static_cast<std::size_t>(int_min) - int_digits.size(), '0');
  }

  // Remove a redundant leading "0" for values with an integer part of zero
  // when the format has only `#` digits in the integer part.
  if (section.integer_zero_digits == 0 && section.integer_pad_digits == 0 && int_digits == "0") {
    int_digits.clear();
  }

  if (section.thousands_separator) {
    insert_thousands(int_digits);
  }

  // Fraction adjustment: the snprintf produced exactly `frac_digits` chars.
  // Trim trailing zeros matching the `#` tokens (scan right-to-left).
  int trim = section.fraction_opt_digits;
  while (trim > 0 && !frac_digits_str.empty() && frac_digits_str.back() == '0') {
    frac_digits_str.pop_back();
    --trim;
  }

  // Now walk tokens and emit output.
  std::string result;
  if (negative) {
    result.push_back('-');
  }
  bool emitted_integer = false;
  std::size_t frac_cursor = 0;

  auto emit_integer_block = [&]() {
    if (!emitted_integer) {
      result.append(int_digits);
      emitted_integer = true;
    }
  };

  // We handle the integer block as a unit: once any digit token in the
  // integer part is reached, we splat `int_digits` into the output and
  // ignore subsequent integer digit tokens (they were accounted for during
  // classification). Literals intermixed with integer digits are emitted
  // in-place; the single-block strategy is accurate for the common cases
  // (`#,##0` / `0.00` / `0%` / `0E+00`) and does not model format-strings
  // that want to interleave literal text between digit positions.
  bool past_point = false;
  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    const Token& tk = section.tokens[i];
    switch (tk.kind) {
      case Tok::DigitZero:
      case Tok::DigitOpt:
      case Tok::DigitPad:
        if (past_point) {
          if (frac_cursor < frac_digits_str.size()) {
            result.push_back(frac_digits_str[frac_cursor]);
            ++frac_cursor;
          } else {
            // Exceeds precision; emit padding based on token kind.
            if (tk.kind == Tok::DigitZero) {
              result.push_back('0');
            } else if (tk.kind == Tok::DigitPad) {
              result.push_back(' ');
            }
          }
        } else {
          emit_integer_block();
        }
        break;
      case Tok::Point:
        if (section.fraction_zero_digits + section.fraction_opt_digits + section.fraction_pad_digits > 0 ||
            !frac_digits_str.empty()) {
          result.push_back('.');
        }
        past_point = true;
        break;
      case Tok::Comma:
        // Only emit as a literal `,` when it is neither a trailing-scale
        // comma nor the thousands-separator marker (both of which were
        // handled during classification / integer block emission).
        break;
      case Tok::Percent:
        result.push_back('%');
        break;
      case Tok::Literal:
        if (tk.lit_end > tk.lit_begin) {
          result.append(fmt.data() + tk.lit_begin, tk.lit_end - tk.lit_begin);
        }
        break;
      case Tok::SciPlus:
      case Tok::SciMinus:
        // Scientific path handled above; should not reach here.
        break;
      case Tok::At:
        // `@` in a numeric context is ignored (Excel emits nothing).
        break;
      default:
        break;
    }
  }
  out.append(result);
}

void render_date(const Section& section, std::string_view fmt, double serial, std::string& out) {
  if (serial < 0.0 || serial > 2958465.0) {
    // Excel rejects out-of-range serials from TEXT.
    return;
  }
  const date_time::YMD ymd = date_time::ymd_from_serial(serial);
  // Weekday Sunday=0..Saturday=6 computed from the civil day count.
  const std::int64_t days = date_time::days_from_civil(ymd.y, ymd.m, ymd.d);
  const int sun0 = static_cast<int>(((days + 4) % 7 + 7) % 7);

  // Decompose the time portion with optional fractional seconds.
  const double frac_day = serial - std::floor(serial);
  // Total seconds (float-precision) so fractional seconds survive.
  double total_seconds_f = frac_day * 86400.0;
  // Round to `frac_sec_digits` if requested, otherwise to whole seconds.
  double rounded = total_seconds_f;
  if (section.frac_sec_digits == 0) {
    rounded = std::floor(total_seconds_f + 0.5);
  } else {
    const double scale = std::pow(10.0, section.frac_sec_digits);
    rounded = std::floor(total_seconds_f * scale + 0.5) / scale;
  }
  // Extract integer h/m/s and fractional remainder.
  long long total_int_seconds = static_cast<long long>(std::floor(rounded));
  const double sub_sec_float = rounded - static_cast<double>(total_int_seconds);
  // If AM/PM is in use, we need to know it before formatting hours.
  bool use_am_pm = false;
  for (const Token& tk : section.tokens) {
    if (tk.kind == Tok::AmPm || tk.kind == Tok::AP) {
      use_am_pm = true;
      break;
    }
  }

  const long long seconds_of_day = ((total_int_seconds % 86400) + 86400) % 86400;
  unsigned hour_24 = static_cast<unsigned>(seconds_of_day / 3600);
  unsigned minute = static_cast<unsigned>((seconds_of_day / 60) % 60);
  unsigned second = static_cast<unsigned>(seconds_of_day % 60);
  bool pm = hour_24 >= 12u;
  unsigned hour_for_render = hour_24;
  if (use_am_pm) {
    hour_for_render = hour_24 % 12u;
    if (hour_for_render == 0u) {
      hour_for_render = 12u;
    }
  }

  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    const Token& tk = section.tokens[i];
    switch (tk.kind) {
      case Tok::DateY2: {
        unsigned y2 = static_cast<unsigned>(((ymd.y % 100) + 100) % 100);
        append_pad2(out, y2);
        break;
      }
      case Tok::DateY4: {
        char buf[16];
        const int n = std::snprintf(buf, sizeof(buf), "%04d", ymd.y);
        if (n > 0) {
          out.append(buf, static_cast<std::size_t>(n));
        }
        break;
      }
      case Tok::DateM:
        out.append(std::to_string(ymd.m));
        break;
      case Tok::DateMM:
        append_pad2(out, ymd.m);
        break;
      case Tok::DateMMM:
        out.append(month_short(ymd.m));
        break;
      case Tok::DateMMMM:
        out.append(month_long(ymd.m));
        break;
      case Tok::DateD:
        out.append(std::to_string(ymd.d));
        break;
      case Tok::DateDD:
        append_pad2(out, ymd.d);
        break;
      case Tok::DateDDD:
        out.append(weekday_short(sun0));
        break;
      case Tok::DateDDDD:
        out.append(weekday_long(sun0));
        break;
      case Tok::DateH:
        out.append(std::to_string(hour_for_render));
        break;
      case Tok::DateHH:
        append_pad2(out, hour_for_render);
        break;
      case Tok::DateMin:
        out.append(std::to_string(minute));
        break;
      case Tok::DateMMMin:
        append_pad2(out, minute);
        break;
      case Tok::DateS:
        out.append(std::to_string(second));
        break;
      case Tok::DateSS:
        append_pad2(out, second);
        break;
      case Tok::DateElapsedH: {
        // Total hours since serial 0 (integer floor).
        const long long total_hours = static_cast<long long>(std::floor(serial * 24.0));
        out.append(std::to_string(total_hours));
        break;
      }
      case Tok::DateElapsedM: {
        const long long total_minutes = static_cast<long long>(std::floor(serial * 1440.0));
        out.append(std::to_string(total_minutes));
        break;
      }
      case Tok::DateElapsedS: {
        const long long total_sec = static_cast<long long>(std::floor(serial * 86400.0));
        out.append(std::to_string(total_sec));
        break;
      }
      case Tok::AmPm:
        out.append(pm ? "PM" : "AM");
        break;
      case Tok::AP:
        out.append(pm ? "P" : "A");
        break;
      case Tok::FracSecDigits: {
        // Render fractional seconds at the requested precision.
        const int digits = static_cast<int>(tk.width);
        if (digits > 0) {
          out.push_back('.');
          double f = sub_sec_float;
          if (f < 0.0) {
            f = 0.0;
          }
          for (int k = 0; k < digits; ++k) {
            f *= 10.0;
            int d = static_cast<int>(std::floor(f));
            if (d > 9) {
              d = 9;
            } else if (d < 0) {
              d = 0;
            }
            out.push_back(static_cast<char>('0' + d));
            f -= static_cast<double>(d);
          }
        }
        break;
      }
      case Tok::Literal:
        if (tk.lit_end > tk.lit_begin) {
          out.append(fmt.data() + tk.lit_begin, tk.lit_end - tk.lit_begin);
        }
        break;
      default:
        break;
    }
  }
}

// Renders the text section by walking the raw format bytes. We do not
// reuse `section.tokens` here because date letters that snuck into literal
// phrases (e.g. the `s` in "text is @") would have been promoted to
// DateS tokens during tokenisation and lost their positional info. Walking
// the raw format bytes avoids that pitfall while still honouring `"..."`
// quoted literals, `\x` / `!x` escapes, and `[...]` bracketed discards
// (e.g. colour markers).
void render_text_section(const Section& /*section*/, std::string_view fmt, std::string_view original,
                         std::string& out) {
  std::size_t i = 0;
  while (i < fmt.size()) {
    const char c = fmt[i];
    if (c == '"') {
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != '"') {
        out.push_back(fmt[j]);
        ++j;
      }
      i = j < fmt.size() ? j + 1 : j;
      continue;
    }
    if ((c == '\\' || c == '!') && i + 1 < fmt.size()) {
      out.push_back(fmt[i + 1]);
      i += 2;
      continue;
    }
    if (c == '[') {
      // Skip to matching `]` (colour / locale markers discarded).
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != ']') {
        ++j;
      }
      i = j < fmt.size() ? j + 1 : j;
      continue;
    }
    if (c == '@') {
      out.append(original);
      ++i;
      continue;
    }
    out.push_back(c);
    ++i;
  }
}

}  // namespace number_format_detail
}  // namespace eval
}  // namespace formulon
