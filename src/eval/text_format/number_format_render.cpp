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
#include <vector>

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

// ja-JP weekday tokens (`aaa` / `aaaa`). Index 0 = Sunday to match the
// `sun0` index used elsewhere in this file.
const char* weekday_ja_short(int sun0) noexcept {
  // 日, 月, 火, 水, 木, 金, 土 (each is a 3-byte UTF-8 code point).
  static const char* kTable[7] = {"\xE6\x97\xA5", "\xE6\x9C\x88", "\xE7\x81\xAB", "\xE6\xB0\xB4",
                                  "\xE6\x9C\xA8", "\xE9\x87\x91", "\xE5\x9C\x9F"};
  if (sun0 < 0 || sun0 > 6) {
    return "";
  }
  return kTable[sun0];
}

const char* weekday_ja_long(int sun0) noexcept {
  // <weekday>曜日 — the suffix is `\xE6\x9B\x9C\xE6\x97\xA5` (曜日).
  static const char* kTable[7] = {
      "\xE6\x97\xA5\xE6\x9B\x9C\xE6\x97\xA5",  // 日曜日
      "\xE6\x9C\x88\xE6\x9B\x9C\xE6\x97\xA5",  // 月曜日
      "\xE7\x81\xAB\xE6\x9B\x9C\xE6\x97\xA5",  // 火曜日
      "\xE6\xB0\xB4\xE6\x9B\x9C\xE6\x97\xA5",  // 水曜日
      "\xE6\x9C\xA8\xE6\x9B\x9C\xE6\x97\xA5",  // 木曜日
      "\xE9\x87\x91\xE6\x9B\x9C\xE6\x97\xA5",  // 金曜日
      "\xE5\x9C\x9F\xE6\x9B\x9C\xE6\x97\xA5",  // 土曜日
  };
  if (sun0 < 0 || sun0 > 6) {
    return "";
  }
  return kTable[sun0];
}

// Japanese era classification. A date is bucketed into one of five eras by
// (year, month, day) using Mac Excel-observable boundaries:
//   * 1868-01-25 -> 1912-07-29: Meiji
//   * 1912-07-30 -> 1926-12-24: Taisho
//   * 1926-12-25 -> 1989-01-07: Showa
//   * 1989-01-08 -> 2019-04-30: Heisei
//   * 2019-05-01 onward:        Reiwa
// For dates before Meiji (rare in practice for Excel TEXT calls — serial 1
// is 1900-01-01), we fall back to Meiji's era anchor and let the year math
// produce a non-positive era year. The 5-bucket table is duplicated from
// `date_text_parse.cpp`'s parser; see that TU for the parser-side anchors.
struct EraInfo {
  int start_year;
  unsigned start_month;
  unsigned start_day;
  int year_anchor;     // Gregorian year corresponding to era year 1.
  const char* roman;   // 1-byte ASCII era abbreviation for `g`.
  const char* kanji1;  // 3-byte UTF-8 single-kanji abbreviation for `gg`.
  const char* kanji2;  // 6-byte UTF-8 full era name for `ggg`.
};

const EraInfo& classify_era(int year, unsigned month, unsigned day) noexcept {
  static const EraInfo kReiwa{2019, 5u, 1u, 2019, "R", "\xE4\xBB\xA4", "\xE4\xBB\xA4\xE5\x92\x8C"};
  static const EraInfo kHeisei{1989, 1u, 8u, 1989, "H", "\xE5\xB9\xB3", "\xE5\xB9\xB3\xE6\x88\x90"};
  static const EraInfo kShowa{1926, 12u, 25u, 1926, "S", "\xE6\x98\xAD", "\xE6\x98\xAD\xE5\x92\x8C"};
  static const EraInfo kTaisho{1912, 7u, 30u, 1912, "T", "\xE5\xA4\xA7", "\xE5\xA4\xA7\xE6\xAD\xA3"};
  static const EraInfo kMeiji{1868, 1u, 25u, 1868, "M", "\xE6\x98\x8E", "\xE6\x98\x8E\xE6\xB2\xBB"};
  // Rank a (Y, M, D) triple for ordered comparison.
  auto cmp = [](int y1, unsigned m1, unsigned d1, int y2, unsigned m2, unsigned d2) {
    if (y1 != y2) {
      return y1 < y2;
    }
    if (m1 != m2) {
      return m1 < m2;
    }
    return d1 < d2;
  };
  if (!cmp(year, month, day, kReiwa.start_year, kReiwa.start_month, kReiwa.start_day)) {
    return kReiwa;
  }
  if (!cmp(year, month, day, kHeisei.start_year, kHeisei.start_month, kHeisei.start_day)) {
    return kHeisei;
  }
  if (!cmp(year, month, day, kShowa.start_year, kShowa.start_month, kShowa.start_day)) {
    return kShowa;
  }
  if (!cmp(year, month, day, kTaisho.start_year, kTaisho.start_month, kTaisho.start_day)) {
    return kTaisho;
  }
  // Pre-Meiji dates: still classified as Meiji (anchor 1868). Mac Excel
  // does not validate, so e.g. an Edo-period serial would emit a
  // negative-or-zero `era_year`; this matches the parser's lenient
  // behaviour in `date_text_parse.cpp`.
  return kMeiji;
}

void append_pad2(std::string& out, unsigned n) {
  if (n < 10u) {
    out.push_back('0');
  }
  out.append(std::to_string(n));
}

// --- DBNum digit substitution -----------------------------------------
//
// `[DBNum1]` / `[DBNum2]` / `[DBNum3]` are *per-digit* substitutions in
// Mac Excel 365 ja-JP. Despite their popular description as "positional
// kanji" formats, the oracle corpus shows that Excel does NOT decompose
// integers into 千 / 百 / 十 groups -- it simply rewrites each ASCII
// digit through a fixed table. Concretely:
//
//   * `=TEXT(1234, "[DBNum1]0")` -> `一二三四` (NOT `一千二百三十四`).
//   * `=TEXT(1234, "[DBNum2]0")` -> `壱弐参四` (NOT `壱阡弐百参拾四`).
//   * `=TEXT(1234, "[DBNum3]0")` -> `１２３４` (full-width Arabic).
//
// Tables:
//   * DBNum1: 〇一二三四五六七八九 (weak-form everyday kanji digits).
//   * DBNum2: 零壱弐参四伍六七捌玖. Digit 4 is the everyday form 四 (not
//     大字 肆) because Mac Excel's `1234` golden ends in 四. The 5/6/7/8/9
//     entries are not exercised by the oracle corpus; we follow common
//     convention (5=伍, 8=捌, 9=玖 are 大字; 6=六, 7=七 left as everyday
//     since 4=四 sets that precedent).
//   * DBNum3: U+FF10..U+FF19 (full-width Arabic digits).

const char* dbnum1_digit(char ascii_digit) noexcept {
  // 〇一二三四五六七八九 (each is a 3-byte UTF-8 code point).
  static const char* kDigits[10] = {
      "\xE3\x80\x87",  // 〇
      "\xE4\xB8\x80",  // 一
      "\xE4\xBA\x8C",  // 二
      "\xE4\xB8\x89",  // 三
      "\xE5\x9B\x9B",  // 四
      "\xE4\xBA\x94",  // 五
      "\xE5\x85\xAD",  // 六
      "\xE4\xB8\x83",  // 七
      "\xE5\x85\xAB",  // 八
      "\xE4\xB9\x9D",  // 九
  };
  if (ascii_digit < '0' || ascii_digit > '9') {
    return "";
  }
  return kDigits[static_cast<std::size_t>(ascii_digit - '0')];
}

const char* dbnum2_digit(char ascii_digit) noexcept {
  // Mac Excel-observed 大字 mapping: ones digits = 零壱弐参四伍六七捌玖.
  // (Note 4=四 not 肆, 7=七 not 漆, per Mac Excel's empirical output.)
  static const char* kDigits[10] = {
      "\xE9\x9B\xB6",  // 零
      "\xE5\xA3\xB1",  // 壱
      "\xE5\xBC\x90",  // 弐
      "\xE5\x8F\x82",  // 参
      "\xE5\x9B\x9B",  // 四 (everyday, matches Mac Excel oracle)
      "\xE4\xBC\x8D",  // 伍
      "\xE5\x85\xAD",  // 六 (everyday, matches Mac Excel oracle)
      "\xE4\xB8\x83",  // 七 (everyday, matches Mac Excel oracle)
      "\xE6\x8D\x8C",  // 捌
      "\xE7\x8E\x96",  // 玖
  };
  if (ascii_digit < '0' || ascii_digit > '9') {
    return "";
  }
  return kDigits[static_cast<std::size_t>(ascii_digit - '0')];
}

// Full-width Arabic digits U+FF10..U+FF19 (each is a 3-byte UTF-8 sequence
// `EF BC 9X` for X in 0..9).
const char* dbnum3_digit(char ascii_digit) noexcept {
  static const char* kDigits[10] = {
      "\xEF\xBC\x90", "\xEF\xBC\x91", "\xEF\xBC\x92", "\xEF\xBC\x93", "\xEF\xBC\x94",
      "\xEF\xBC\x95", "\xEF\xBC\x96", "\xEF\xBC\x97", "\xEF\xBC\x98", "\xEF\xBC\x99",
  };
  if (ascii_digit < '0' || ascii_digit > '9') {
    return "";
  }
  return kDigits[static_cast<std::size_t>(ascii_digit - '0')];
}

// Returns the per-digit substitution for `c` under `mode`, or an empty
// string if no substitution applies (caller falls back to `c` verbatim).
std::string_view dbnum_digit_subst(DbNumMode mode, char c) noexcept {
  if (c < '0' || c > '9') {
    return {};
  }
  switch (mode) {
    case DbNumMode::kDBNum1:
      return dbnum1_digit(c);
    case DbNumMode::kDBNum2:
      return dbnum2_digit(c);
    case DbNumMode::kDBNum3:
      return dbnum3_digit(c);
    case DbNumMode::kNone:
    default:
      return {};
  }
}

// Appends `value` to `out` with each digit substituted per `mode`. Used
// for date components (era year, m, d, h, min, s) where positional kanji
// do NOT apply -- only per-digit substitution.
void append_int_dbnum(std::string& out, long long value, DbNumMode mode) {
  std::string buf = std::to_string(value);
  if (mode == DbNumMode::kNone) {
    out.append(buf);
    return;
  }
  for (char c : buf) {
    if (c == '-') {
      out.push_back('-');
      continue;
    }
    const std::string_view sub = dbnum_digit_subst(mode, c);
    if (!sub.empty()) {
      out.append(sub);
    } else {
      out.push_back(c);
    }
  }
}

// Appends `value` zero-padded to 2 digits, with DBNum substitution applied.
void append_pad2_dbnum(std::string& out, unsigned value, DbNumMode mode) {
  if (mode == DbNumMode::kNone) {
    append_pad2(out, value);
    return;
  }
  if (value < 10u) {
    const std::string_view zero = dbnum_digit_subst(mode, '0');
    if (!zero.empty()) {
      out.append(zero);
    } else {
      out.push_back('0');
    }
  }
  append_int_dbnum(out, static_cast<long long>(value), mode);
}

// Rounds `v` to `decimals` fractional places, half-away-from-zero. The
// result's integer and fractional digit strings are written into
// `*int_digits` and `*frac_digits` (decimal digits only, no sign, no
// leading zeros on the integer side for zero values — except the integer
// is always at least `"0"`).
//
// This helper does not handle scientific notation; see `render_scientific`.
void format_fixed_digits(double v, int decimals, bool* negative, std::string* int_digits, std::string* frac_digits) {
  // Use strict `< 0.0` rather than `signbit` so that `-0.0` rounds to "0" with
  // no sign byte. Mac Excel's two-section formats (`#,##0_);(#,##0)`) expect
  // the positive section to emit "0 " for both `+0` and `-0`; without this
  // guard the minus leaks out via `signbit(-0.0)`.
  *negative = v < 0.0;
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

// Excel's `General` format code produces an ~11-character-wide numeric
// display: fixed-point when the value fits, scientific notation otherwise.
// This matches Mac Excel 365 / ja-JP for TEXT() calls and is the rendering
// chosen for the IronCalc oracle goldens. See the Microsoft reference on
// "General" number formatting:
// https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-...
//
// Rough shape of the algorithm (for non-zero `abs_v`):
//   * Pure-integer values that fit in 11 decimal digits are printed without
//     a decimal point (e.g. `12`, `1234567890`).
//   * Otherwise with exponent e = floor(log10(abs_v)):
//       - If `11 > e >= -4`, use fixed form with `(11 - int_len - 1)`
//         fractional digits (rounded, trailing zeros trimmed).
//       - Else, use scientific with as many mantissa digits as fit in 11
//         characters: `1.<frac>E+<exp>` or `1.<frac>E-<exp>`.
//
// Caller supplies the value already *sign-stripped*: `format_general` only
// emits the magnitude. The sign prefix (if any) is owned by the outer
// `render_numeric` routine.
void format_general(std::string& out, double v) {
  if (v == 0.0) {
    out.push_back('0');
    return;
  }
  const double abs_v = std::fabs(v);
  // Integer fast path: values whose fractional part is exactly zero and
  // whose magnitude fits in roughly 11 decimal digits print verbatim.
  if (abs_v < 1e11) {
    const double truncated = std::trunc(v);
    if (truncated == v) {
      const std::int64_t as_int = static_cast<std::int64_t>(truncated);
      out.append(std::to_string(as_int));
      return;
    }
  }
  const int exp10 = static_cast<int>(std::floor(std::log10(abs_v)));
  // Switch to scientific for large / very-small magnitudes. Mac Excel /
  // IronCalc goldens flip to scientific at `>=1e11` and `<=1e-9`; anything
  // between stays in fixed form even if it grows long (e.g. `0.00000001`).
  const bool use_scientific = (exp10 >= 11 || exp10 <= -9);
  if (use_scientific) {
    // How wide is the `E+XX` / `E-YYY` tail?
    int exp_abs = exp10 < 0 ? -exp10 : exp10;
    int exp_digits = 1;
    if (exp_abs >= 10) {
      exp_digits = 2;
    }
    if (exp_abs >= 100) {
      exp_digits = 3;
    }
    // Budget 11 chars total. Layout: `X.FFF...E+YY` with 1 int digit, a
    // dot, `frac` fractional digits, an `E`, a sign, and `exp_digits`.
    int frac = 11 - 4 - exp_digits;  // 4 = "X." + "E" + sign
    if (frac < 0) {
      frac = 0;
    }
    char buf[64];
    const int n = std::snprintf(buf, sizeof(buf), "%.*e", frac, v);
    if (n <= 0) {
      out.append(std::to_string(v));
      return;
    }
    std::string_view s(buf, static_cast<std::size_t>(n));
    std::size_t epos = s.find('e');
    if (epos == std::string_view::npos) {
      out.append(s);
      return;
    }
    std::string_view mantissa = s.substr(0, epos);
    std::string_view exp_part = s.substr(epos + 1);
    // Trim trailing zeros from the mantissa's fractional part. `2.50000`
    // collapses to `2.5`; `1.00000` collapses to `1`.
    std::size_t dot = mantissa.find('.');
    std::size_t mantissa_end = mantissa.size();
    if (dot != std::string_view::npos) {
      while (mantissa_end > dot + 1 && mantissa[mantissa_end - 1] == '0') {
        --mantissa_end;
      }
      if (mantissa_end > 0 && mantissa[mantissa_end - 1] == '.') {
        --mantissa_end;
      }
    }
    out.append(mantissa.data(), mantissa_end);
    out.push_back('E');
    // Exponent: always emit an explicit sign, and pad the exponent digits
    // to at least two characters (e.g. `E+09`, not `E+9`). Mac Excel /
    // IronCalc both zero-pad the exponent to two digits, matching the
    // printf `%e` convention.
    char sign = '+';
    std::size_t pos = 0;
    if (!exp_part.empty() && (exp_part[0] == '+' || exp_part[0] == '-')) {
      sign = exp_part[0];
      pos = 1;
    }
    // Strip leading zeros down to a minimum of two digits.
    while (exp_part.size() - pos > 2 && exp_part[pos] == '0') {
      ++pos;
    }
    out.push_back(sign);
    out.append(exp_part.data() + pos, exp_part.size() - pos);
    return;
  }
  // Fixed-point branch: size the fractional-digit count so the whole number
  // (integer part + "." + fraction) spans up to 11 characters. Trailing
  // zeros are trimmed afterwards, so e.g. `0.0001` collapses from the
  // 11-char `0.000100000` down to `0.0001`.
  const int integer_digits = (exp10 >= 0) ? (exp10 + 1) : 1;
  int frac = 11 - integer_digits - 1;  // "1" covers the decimal point.
  if (frac < 0) {
    frac = 0;
  }
  char buf[64];
  const int n = std::snprintf(buf, sizeof(buf), "%.*f", frac, v);
  if (n <= 0) {
    out.append(std::to_string(v));
    return;
  }
  std::string_view s(buf, static_cast<std::size_t>(n));
  std::size_t dot = s.find('.');
  std::size_t end = s.size();
  if (dot != std::string_view::npos) {
    while (end > dot + 1 && s[end - 1] == '0') {
      --end;
    }
    if (end > 0 && s[end - 1] == '.') {
      --end;
    }
  }
  out.append(s.data(), end);
}

}  // namespace

// Forward declaration for the fraction-format helper defined below. Used
// from `render_numeric` when `section.is_fraction` is true.
void render_fraction(const Section& section, std::string_view fmt, double value, std::string& out);

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
  if (section.is_fraction) {
    render_fraction(section, fmt, value, out);
    return;
  }
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

  // Fraction adjustment: the snprintf produced exactly `frac_digits` chars.
  // Trim trailing zeros matching the `#` tokens (scan right-to-left).
  int trim = section.fraction_opt_digits;
  while (trim > 0 && !frac_digits_str.empty() && frac_digits_str.back() == '0') {
    frac_digits_str.pop_back();
    --trim;
  }

  // Locate the decimal point inside the token stream so we can partition the
  // digit tokens into integer / fraction stacks. This mirrors the classifier,
  // but we need the position again at walk time to drive the interleaving.
  int point_index = -1;
  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    if (section.tokens[i].kind == Tok::Point) {
      point_index = static_cast<int>(i);
      break;
    }
  }
  // Scientific tokens were handled in the earlier branch; for plain numeric
  // rendering we ignore any post-scientific digit tokens (there are none in
  // practice since the scientific branch returns early).
  auto is_digit_tok = [](Tok k) { return k == Tok::DigitZero || k == Tok::DigitOpt || k == Tok::DigitPad; };

  // Gather integer digit token positions (in token order). We distribute the
  // pure `int_digits` characters across these positions right-to-left, so
  // format strings like `"00-00-00-00"` place literals between digit groups.
  std::vector<std::size_t> int_digit_positions;
  int_digit_positions.reserve(section.tokens.size());
  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    if (point_index >= 0 && static_cast<int>(i) >= point_index) {
      break;
    }
    if (is_digit_tok(section.tokens[i].kind)) {
      int_digit_positions.push_back(i);
    }
  }
  const std::size_t n_int_tokens = int_digit_positions.size();
  const std::size_t n_int_digits = int_digits.size();

  // `int_slot_text[i]` is the plain-digit substring (no thousands commas)
  // that the i-th integer-digit token should emit. Empty slot = fallback.
  std::vector<std::string> int_slot_text(n_int_tokens);
  if (n_int_tokens > 0 && n_int_digits > 0) {
    if (n_int_digits >= n_int_tokens) {
      // First token absorbs the overflow prefix; each later token takes one
      // digit. Total characters placed = n_int_digits.
      const std::size_t prefix_len = n_int_digits - n_int_tokens + 1;
      int_slot_text[0].assign(int_digits, 0, prefix_len);
      for (std::size_t i = 1; i < n_int_tokens; ++i) {
        int_slot_text[i].assign(1, int_digits[prefix_len + i - 1]);
      }
    } else {
      // Fewer digits than tokens: leading tokens fall back, trailing tokens
      // each get one digit.
      const std::size_t fallback_count = n_int_tokens - n_int_digits;
      for (std::size_t i = 0; i < n_int_digits; ++i) {
        int_slot_text[fallback_count + i].assign(1, int_digits[i]);
      }
    }
  }

  // Suppress the sign prefix when the rounded representation is numerically
  // zero. Mac Excel's single-section formats (`"0"`, `"00-00-00-00"`, etc.)
  // render `TEXT(-1/3, "0")` as `"0"`, not `"-0"`: once the rounding has
  // crushed the magnitude below the displayable precision, no sign leaks out.
  if (negative && !section.has_general) {
    bool all_zero = true;
    for (char ch : int_digits) {
      if (ch != '0') {
        all_zero = false;
        break;
      }
    }
    if (all_zero) {
      for (char ch : frac_digits_str) {
        if (ch != '0') {
          all_zero = false;
          break;
        }
      }
    }
    if (all_zero) {
      negative = false;
    }
  }

  // Now walk tokens and emit output.
  std::string result;
  if (negative) {
    result.push_back('-');
  }
  // Cursor into the pure integer digit stream (0-based, left-to-right).
  // Used to decide when to prepend a thousands-separator comma.
  std::size_t int_cursor = 0;
  std::size_t int_token_cursor = 0;
  std::size_t frac_cursor = 0;
  bool past_point = false;

  auto emit_int_digit_char = [&](char digit) {
    if (section.thousands_separator && int_cursor > 0 && (n_int_digits - int_cursor) % 3 == 0) {
      result.push_back(',');
    }
    const std::string_view sub = dbnum_digit_subst(section.dbnum_mode, digit);
    if (!sub.empty()) {
      result.append(sub);
    } else {
      result.push_back(digit);
    }
    ++int_cursor;
  };

  // Helper: emit a single fractional digit with DBNum substitution.
  auto emit_frac_digit_char = [&](char digit) {
    const std::string_view sub = dbnum_digit_subst(section.dbnum_mode, digit);
    if (!sub.empty()) {
      result.append(sub);
    } else {
      result.push_back(digit);
    }
  };

  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    const Token& tk = section.tokens[i];
    switch (tk.kind) {
      case Tok::DigitZero:
      case Tok::DigitOpt:
      case Tok::DigitPad:
        if (past_point) {
          if (frac_cursor < frac_digits_str.size()) {
            emit_frac_digit_char(frac_digits_str[frac_cursor]);
            ++frac_cursor;
          } else {
            // Exceeds precision; emit padding based on token kind.
            if (tk.kind == Tok::DigitZero) {
              emit_frac_digit_char('0');
            } else if (tk.kind == Tok::DigitPad) {
              result.push_back(' ');
            }
          }
        } else {
          const std::string& slot = int_slot_text[int_token_cursor];
          if (!slot.empty()) {
            for (char d : slot) {
              emit_int_digit_char(d);
            }
          } else {
            // Fallback for an integer digit token with no assigned digit.
            if (tk.kind == Tok::DigitZero) {
              // `0` forces a literal '0' when the digit stream is exhausted.
              emit_int_digit_char('0');
            } else if (tk.kind == Tok::DigitPad) {
              result.push_back(' ');
            }
            // `#` emits nothing.
          }
          ++int_token_cursor;
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
        // Thousands-separator markers and trailing-scale commas were
        // already consumed during classification and (for thousands) during
        // per-digit emission. Literal `,` outside that context has no
        // special meaning in Excel's format language.
        break;
      case Tok::Percent:
        result.push_back('%');
        break;
      case Tok::Literal:
        if (tk.lit_end > tk.lit_begin) {
          result.append(fmt.data() + tk.lit_begin, tk.lit_end - tk.lit_begin);
        }
        break;
      case Tok::Space:
        // `_X` underscore-skip: emit a single space placeholder.
        result.push_back(' ');
        break;
      case Tok::GeneralNumber: {
        // Render through the Excel-specific 11-character `General` formatter
        // on the absolute value; the sign prefix was already emitted above
        // (for single-section formats) or stripped by the section selector
        // (two-section formats pass an already-positive `value`).
        format_general(result, std::fabs(value));
        break;
      }
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

// --- Fraction format rendering (`# ?/?`, `# ??/??`, `0/0`, ...) ---------
//
// Implements Excel's bounded best-rational-approximation via a Stern-Brocot
// mediant search. Compared to the more familiar continued-fraction
// algorithm, Stern-Brocot occasionally picks a different mediant when the
// target lies near a Farey-neighbour boundary; Mac Excel's empirical output
// matches Stern-Brocot, so we follow that.
//
// For a value `v` (already absolute) and a denominator cap `max_q` and
// numerator cap `max_p`, the search finds (p, q) minimising |v - p/q|
// with 1 <= q <= max_q and 0 <= p <= max_p. When an integer group is
// present the search runs against `frac = v - floor(v)` only.
namespace {

// Returns 10^N for small non-negative N. Always fits in a long long for
// the digit-counts we accept (Excel caps fraction placeholder runs well
// below 18 digits in practice; for safety we cap at 9 here).
long long fraction_pow10(int n) noexcept {
  long long r = 1;
  for (int i = 0; i < n && i < 18; ++i) {
    r *= 10;
  }
  return r;
}

// Stern-Brocot bounded mediant search. Returns the best (num, den) with
// `1 <= den <= max_q` and `0 <= num <= max_p`, ties broken by smaller den
// then smaller num (the search's natural traversal order).
void best_rational(double target, long long max_p, long long max_q, long long* out_num, long long* out_den) noexcept {
  if (max_q < 1) {
    max_q = 1;
  }
  if (max_p < 0) {
    max_p = 0;
  }
  long long a_num = 0;
  long long a_den = 1;
  long long b_num = 1;
  long long b_den = 0;  // Represents +infinity.
  long long best_num = 0;
  long long best_den = 1;
  double best_err = std::fabs(target - 0.0);
  for (int iter = 0; iter < 10000; ++iter) {
    const long long m_num = a_num + b_num;
    const long long m_den = a_den + b_den;
    if (m_den > max_q || m_num > max_p) {
      break;
    }
    const double m = static_cast<double>(m_num) / static_cast<double>(m_den);
    const double err = std::fabs(target - m);
    if (err < best_err) {
      best_err = err;
      best_num = m_num;
      best_den = m_den;
    }
    if (target < m) {
      b_num = m_num;
      b_den = m_den;
    } else if (target > m) {
      a_num = m_num;
      a_den = m_den;
    } else {
      break;
    }
  }
  *out_num = best_num;
  *out_den = best_den < 1 ? 1 : best_den;
}

// Emit a non-negative integer `value` right-aligned to `width` characters
// using the placeholder kinds in `[begin, end)`. Each placeholder kind
// determines how unused leading positions render: `0` -> '0' pad, `?` ->
// space pad, `#` -> nothing emitted. The DBNum mapping applies to digits
// (and to `0`-pad positions) but never to spaces or absent positions.
void emit_fraction_digits(const Section& section, std::string_view fmt, long long value, int begin, int end,
                          std::string& out) {
  (void)fmt;
  const int width = end - begin;
  if (width <= 0) {
    return;
  }
  std::string digits = std::to_string(value);
  if (static_cast<int>(digits.size()) > width) {
    // Overflow: Excel never produces this for our caps, but be defensive.
    // Emit the digits verbatim (the widest available run cap is enforced
    // by the search above so this case is essentially unreachable).
    for (char c : digits) {
      const std::string_view sub = dbnum_digit_subst(section.dbnum_mode, c);
      if (!sub.empty()) {
        out.append(sub);
      } else {
        out.push_back(c);
      }
    }
    return;
  }
  const int pad = width - static_cast<int>(digits.size());
  // Iterate placeholder kinds left-to-right. The first `pad` placeholders
  // are leading positions with no digit; the remaining `digits.size()`
  // placeholders consume the digit string in order.
  std::size_t digit_cursor = 0;
  for (int k = 0; k < width; ++k) {
    const Tok kind = section.tokens[static_cast<std::size_t>(begin + k)].kind;
    if (k < pad) {
      // Leading-position behaviour by placeholder kind.
      if (kind == Tok::DigitZero) {
        const std::string_view sub = dbnum_digit_subst(section.dbnum_mode, '0');
        if (!sub.empty()) {
          out.append(sub);
        } else {
          out.push_back('0');
        }
      } else if (kind == Tok::DigitPad) {
        out.push_back(' ');
      }
      // `#`: emit nothing.
    } else {
      const char d = digits[digit_cursor++];
      const std::string_view sub = dbnum_digit_subst(section.dbnum_mode, d);
      if (!sub.empty()) {
        out.append(sub);
      } else {
        out.push_back(d);
      }
    }
  }
}

}  // namespace

void render_fraction(const Section& section, std::string_view fmt, double value, std::string& out) {
  // Sign: emit minus prefix on the rendered form for negative values, then
  // operate on the absolute magnitude. Excel's fraction format does not
  // honour `?`/`#` for sign placement -- the leading `-` is unconditional.
  const bool negative = value < 0.0;
  const double abs_v = std::fabs(value);

  const bool has_int_group = section.fraction_int_max_digits > 0;

  // Compute integer part and fraction target. When no integer group is
  // present (`?/?`-style improper fractions), the search runs against the
  // full magnitude and the numerator is allowed to exceed 1.
  long long integer_part = 0;
  double target = abs_v;
  if (has_int_group) {
    integer_part = static_cast<long long>(std::floor(abs_v));
    target = abs_v - static_cast<double>(integer_part);
  }

  // Run the bounded Stern-Brocot search.
  const long long max_p = fraction_pow10(section.fraction_num_max_digits) - 1;
  const long long max_q = fraction_pow10(section.fraction_den_max_digits) - 1;
  long long num = 0;
  long long den = 1;
  // For improper fractions (no integer group) the numerator bound is the
  // raw cap; for proper fractions (target < 1) Stern-Brocot's invariant
  // `m_num <= m_den` means the cap on `p` is implicitly at most `max_q`,
  // so passing the raw `max_p` is also safe.
  best_rational(target, max_p, max_q, &num, &den);

  // Rounding promotion: if the best approximation rounds up to 1 exactly
  // (num == den) and we have an integer group, increment the integer and
  // zero out the fraction.
  if (has_int_group && num == den) {
    integer_part += 1;
    num = 0;
    den = 1;
  }

  std::string result;
  if (negative) {
    result.push_back('-');
  }

  // Walk the section's token stream. Tokens before `fraction_int_begin`
  // (or before `fraction_num_begin` if no integer group) are emitted as
  // literals; the integer group is rendered through `emit_fraction_digits`;
  // the literal between integer and numerator (a single space) is emitted
  // verbatim; numerator group, the slash, denominator group follow; any
  // trailing literals after the denominator group emit verbatim.
  const std::size_t n_tokens = section.tokens.size();
  // Indices.
  const int int_begin = section.fraction_int_begin;
  const int int_end = section.fraction_int_end;
  const int num_begin = section.fraction_num_begin;
  const int num_end = section.fraction_num_end;
  const int den_begin = section.fraction_den_begin;
  const int den_end = section.fraction_den_end;
  const int slash_index = section.fraction_slash_index;

  // Helper to emit a single token verbatim (literals only; non-literal
  // tokens encountered inside fraction sections are skipped).
  auto emit_token_verbatim = [&](std::size_t idx) {
    const Token& tk = section.tokens[idx];
    if (tk.kind == Tok::Literal && tk.lit_end > tk.lit_begin) {
      result.append(fmt.data() + tk.lit_begin, tk.lit_end - tk.lit_begin);
    } else if (tk.kind == Tok::Space) {
      result.push_back(' ');
    } else if (tk.kind == Tok::Percent) {
      result.push_back('%');
    }
  };

  std::size_t i = 0;
  // 1) Pre-integer / pre-numerator literals.
  const int leading_stop = has_int_group ? int_begin : num_begin;
  while (i < static_cast<std::size_t>(leading_stop)) {
    emit_token_verbatim(i);
    ++i;
  }
  // 2) Integer group.
  if (has_int_group) {
    // Excel suppresses the integer when it is zero AND the leading
    // placeholder is `#`; otherwise the leading-pad behaviour from
    // `emit_fraction_digits` handles `0` and `?` correctly.
    const Tok lead_kind = section.tokens[static_cast<std::size_t>(int_begin)].kind;
    if (integer_part == 0 && lead_kind == Tok::DigitOpt) {
      // Emit nothing for the integer group.
    } else {
      emit_fraction_digits(section, fmt, integer_part, int_begin, int_end, result);
    }
    i = static_cast<std::size_t>(int_end);
    // 3) Literals between integer group and numerator group (typically a
    // single space).
    while (i < static_cast<std::size_t>(num_begin)) {
      emit_token_verbatim(i);
      ++i;
    }
  }
  // 4) Numerator group.
  emit_fraction_digits(section, fmt, num, num_begin, num_end, result);
  i = static_cast<std::size_t>(num_end);
  // 5) Literals up to the slash (including the slash itself).
  while (i <= static_cast<std::size_t>(slash_index)) {
    emit_token_verbatim(i);
    ++i;
  }
  // 6) Denominator group.
  emit_fraction_digits(section, fmt, den, den_begin, den_end, result);
  i = static_cast<std::size_t>(den_end);
  // 7) Trailing literals.
  while (i < n_tokens) {
    emit_token_verbatim(i);
    ++i;
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

  const DbNumMode dbnum = section.dbnum_mode;
  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    const Token& tk = section.tokens[i];
    switch (tk.kind) {
      case Tok::DateY2: {
        unsigned y2 = static_cast<unsigned>(((ymd.y % 100) + 100) % 100);
        append_pad2_dbnum(out, y2, dbnum);
        break;
      }
      case Tok::DateY4: {
        char buf[16];
        const int n = std::snprintf(buf, sizeof(buf), "%04d", ymd.y);
        if (n > 0) {
          if (dbnum == DbNumMode::kNone) {
            out.append(buf, static_cast<std::size_t>(n));
          } else {
            for (int k = 0; k < n; ++k) {
              const std::string_view sub = dbnum_digit_subst(dbnum, buf[k]);
              if (!sub.empty()) {
                out.append(sub);
              } else {
                out.push_back(buf[k]);
              }
            }
          }
        }
        break;
      }
      case Tok::DateM:
        append_int_dbnum(out, static_cast<long long>(ymd.m), dbnum);
        break;
      case Tok::DateMM:
        append_pad2_dbnum(out, ymd.m, dbnum);
        break;
      case Tok::DateMMM:
        out.append(month_short(ymd.m));
        break;
      case Tok::DateMMMM:
        out.append(month_long(ymd.m));
        break;
      case Tok::DateMMMMM: {
        // `mmmmm` (run length >= 5) emits the first letter of the English
        // month name. The month-name table only contains ASCII letters, so
        // the first byte is a complete UTF-8 code point.
        const char* name = month_long(ymd.m);
        if (name[0] != '\0') {
          out.push_back(name[0]);
        }
        break;
      }
      case Tok::DateD:
        append_int_dbnum(out, static_cast<long long>(ymd.d), dbnum);
        break;
      case Tok::DateDD:
        append_pad2_dbnum(out, ymd.d, dbnum);
        break;
      case Tok::DateDDD:
        out.append(weekday_short(sun0));
        break;
      case Tok::DateDDDD:
        out.append(weekday_long(sun0));
        break;
      case Tok::DateAaa:
        out.append(weekday_ja_short(sun0));
        break;
      case Tok::DateAaaa:
        out.append(weekday_ja_long(sun0));
        break;
      case Tok::EraG: {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        out.append(era.roman);
        break;
      }
      case Tok::EraGG: {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        out.append(era.kanji1);
        break;
      }
      case Tok::EraGGG: {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        out.append(era.kanji2);
        break;
      }
      case Tok::EraE: {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        const int era_year = ymd.y - era.year_anchor + 1;
        append_int_dbnum(out, static_cast<long long>(era_year), dbnum);
        break;
      }
      case Tok::EraEE: {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        const int era_year = ymd.y - era.year_anchor + 1;
        if (era_year >= 0 && era_year < 100) {
          append_pad2_dbnum(out, static_cast<unsigned>(era_year), dbnum);
        } else {
          append_int_dbnum(out, static_cast<long long>(era_year), dbnum);
        }
        break;
      }
      case Tok::DateH:
        append_int_dbnum(out, static_cast<long long>(hour_for_render), dbnum);
        break;
      case Tok::DateHH:
        append_pad2_dbnum(out, hour_for_render, dbnum);
        break;
      case Tok::DateMin:
        append_int_dbnum(out, static_cast<long long>(minute), dbnum);
        break;
      case Tok::DateMMMin:
        append_pad2_dbnum(out, minute, dbnum);
        break;
      case Tok::DateS:
        append_int_dbnum(out, static_cast<long long>(second), dbnum);
        break;
      case Tok::DateSS:
        append_pad2_dbnum(out, second, dbnum);
        break;
      case Tok::DateElapsedH: {
        // Total hours since serial 0 (integer floor).
        const long long total_hours = static_cast<long long>(std::floor(serial * 24.0));
        append_int_dbnum(out, total_hours, dbnum);
        break;
      }
      case Tok::DateElapsedM: {
        const long long total_minutes = static_cast<long long>(std::floor(serial * 1440.0));
        append_int_dbnum(out, total_minutes, dbnum);
        break;
      }
      case Tok::DateElapsedS: {
        const long long total_sec = static_cast<long long>(std::floor(serial * 86400.0));
        append_int_dbnum(out, total_sec, dbnum);
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
            const char ch = static_cast<char>('0' + d);
            const std::string_view sub = dbnum_digit_subst(dbnum, ch);
            if (!sub.empty()) {
              out.append(sub);
            } else {
              out.push_back(ch);
            }
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
      case Tok::Space:
        // `_X` underscore-skip: emit a single space placeholder.
        out.push_back(' ');
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
    if (c == '_' && i + 1 < fmt.size()) {
      // `_X` underscore-skip: emit a single space and consume both bytes.
      out.push_back(' ');
      i += 2;
      continue;
    }
    out.push_back(c);
    ++i;
  }
}

}  // namespace number_format_detail
}  // namespace eval
}  // namespace formulon
