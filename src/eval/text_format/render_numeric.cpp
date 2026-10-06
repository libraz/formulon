//
// Standard numeric and `General` rendering for the Excel TEXT() engine.
// See `number_format_render.cpp` for shared helpers and the text-section
// walker; `render_date.cpp` for date/time tokens; `render_fraction.cpp`
// for `# ?/?` style formats.

#include "eval/text_format/render_numeric.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <vector>

#include "eval/text_format/number_format_types.h"
#include "eval/text_format/render_common.h"
#include "eval/text_format/render_fraction.h"
#include "eval/text_format/rounding.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
namespace {

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
  const double decimal_rounded = ::formulon::text_format::round_display_decimal(v, decimals);
  // A finite value at the double limit can round beyond its binary range.
  // Convert the original magnitude in that case; the digit-string cap below
  // still rounds its integer part to Excel's 15 significant digits.
  const double rounded = std::isfinite(decimal_rounded) ? decimal_rounded : v;
  *negative = rounded < 0.0;
  const double abs_v = std::fabs(rounded);
  // `snprintf` only converts an already Excel-rounded decimal value to its
  // digit string. It must not decide the tie direction: Apple libc rounds
  // `%.0f` ties to even, while Excel rounds them away from zero.
  const int requested_decimals = decimals < 0 ? 0 : decimals;
  int rendered_decimals = requested_decimals;
  if (abs_v != 0.0) {
    const int exponent = static_cast<int>(std::floor(std::log10(abs_v)));
    // Once the 15 significant digits have been emitted, any extra fixed
    // places are known zeros. Rendering them from the binary double would
    // reintroduce representation residue (0.1 + 0.2 -> ...00004).
    rendered_decimals = std::min(requested_decimals, std::max(0, 14 - exponent));
  }
  char buf[64];
  const int n = std::snprintf(buf, sizeof(buf), "%.*f", rendered_decimals, abs_v);
  if (n < 0 || static_cast<std::size_t>(n) >= sizeof(buf)) {
    // Fallback: fall back to sprintf with a heap buffer (extremely rare).
    std::string out;
    out.resize(static_cast<std::size_t>(rendered_decimals) + 32u);
    const int m = std::snprintf(&out[0], out.size(), "%.*f", rendered_decimals, abs_v);
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
    frac_digits->append(static_cast<std::size_t>(requested_decimals) - frac_digits->size(), '0');
    cap_integer_significant_digits(int_digits);
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
  frac_digits->append(static_cast<std::size_t>(requested_decimals) - frac_digits->size(), '0');
  cap_integer_significant_digits(int_digits);
}

// Length of `s` once trailing fractional zeros, and then a bare '.', are dropped.
std::size_t trimmed_fraction_length(std::string_view s) {
  std::size_t end = s.size();
  const std::size_t dot = s.find('.');
  if (dot != std::string_view::npos) {
    while (end > dot + 1 && s[end - 1] == '0') {
      --end;
    }
    if (end > 0 && s[end - 1] == '.') {
      --end;
    }
  }
  return end;
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
    out.append(mantissa.data(), trimmed_fraction_length(mantissa));
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
  out.append(buf, trimmed_fraction_length(std::string_view(buf, static_cast<std::size_t>(n))));
}

// Appends one 4-digit group (leading zeros allowed, value non-zero) as kanji
// with 千/百/十 place units; a 1 before a place unit is dropped (1234 ->
// 千二百三十四).
void append_kanji_group(std::string& out, std::string_view group) {
  static const char* const kPlaceUnits[4] = {"", "\xE5\x8D\x81", "\xE7\x99\xBE", "\xE5\x8D\x83"};  // 十 百 千
  for (std::size_t i = 0; i < group.size(); ++i) {
    const char digit = group[i];
    const std::size_t place = group.size() - 1U - i;
    if (digit == '0') {
      continue;
    }
    if (place == 0U || digit != '1') {
      append_digit_dbnum(out, DbNumMode::kDBNum1, digit);
    }
    out.append(kPlaceUnits[place]);
  }
}

// Appends the integer `digits` as positional kanji in 4-digit groups under
// 万/億/兆. A part above 兆 longer than one group has no larger unit and is
// spelled digit by digit.
void append_kanji_integer(std::string& out, std::string_view digits) {
  static const char* const kGroupUnits[4] = {"", "\xE4\xB8\x87", "\xE5\x84\x84", "\xE5\x85\x86"};  // 万 億 兆
  constexpr std::size_t kGroupDigits = 4U;
  constexpr std::size_t kTopGroup = 3U;
  const std::size_t first = digits.find_first_not_of('0');
  if (first == std::string_view::npos) {
    append_digit_dbnum(out, DbNumMode::kDBNum1, '0');
    return;
  }
  digits.remove_prefix(first);
  const std::size_t lower_digits = kTopGroup * kGroupDigits;
  if (digits.size() > lower_digits) {
    const std::string_view top = digits.substr(0, digits.size() - lower_digits);
    if (top.size() <= kGroupDigits) {
      append_kanji_group(out, top);
    } else {
      append_chars_dbnum(out, DbNumMode::kDBNum1, top);
    }
    out.append(kGroupUnits[kTopGroup]);
    digits.remove_prefix(top.size());
  }
  std::size_t unit = (digits.size() + kGroupDigits - 1U) / kGroupDigits;
  std::size_t begin = 0;
  while (unit > 0U) {
    --unit;
    const std::size_t end = digits.size() - unit * kGroupDigits;
    const std::string_view group = digits.substr(begin, end - begin);
    if (group.find_first_not_of('0') != std::string_view::npos) {
      append_kanji_group(out, group);
      out.append(kGroupUnits[unit]);
    }
    begin = end;
  }
}

// `[DBNum1]General`: the integer part in positional kanji, the fraction digit
// by digit. A magnitude General would show in scientific form is written out
// to 15 significant digits instead (1E+15 -> 千兆).
void append_dbnum1_general(std::string& out, double abs_v) {
  std::string general;
  format_general(general, abs_v);
  std::string int_digits;
  std::string frac_digits;
  if (general.find('E') == std::string::npos) {
    const std::size_t dot = general.find('.');
    int_digits = general.substr(0, dot);
    frac_digits = dot == std::string::npos ? std::string() : general.substr(dot + 1U);
  } else if (abs_v >= 1.0) {
    const int exp10 = static_cast<int>(std::floor(std::log10(abs_v)));
    bool negative = false;
    format_fixed_digits(abs_v, std::max(0, 14 - exp10), &negative, &int_digits, &frac_digits);
    while (!frac_digits.empty() && frac_digits.back() == '0') {
      frac_digits.pop_back();
    }
  } else {
    append_chars_dbnum(out, DbNumMode::kDBNum1, general);
    return;
  }
  append_kanji_integer(out, int_digits);
  if (!frac_digits.empty()) {
    out.push_back('.');
    append_chars_dbnum(out, DbNumMode::kDBNum1, frac_digits);
  }
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
FormatStatus render_numeric(const Section& section, std::string_view fmt, double value, std::string& out) {
  double scaled = value;
  if (section.has_percent) {
    scaled *= 100.0;
  }
  for (int i = 0; i < section.trailing_comma_scale; ++i) {
    scaled /= 1000.0;
  }
  if (!std::isfinite(scaled)) {
    return FormatStatus::kOverflow;
  }
  if (section.is_fraction) {
    return render_fraction(section, fmt, scaled, out);
  }
  const int frac_digits = section.fraction_zero_digits + section.fraction_opt_digits + section.fraction_pad_digits;
  bool negative = false;
  std::string int_digits;
  std::string frac_digits_str;

  int exponent = 0;
  if (section.has_scientific) {
    // Scientific notation uses engineering groups sized by every integer
    // placeholder before the decimal point. Required `0`/`?` placeholders
    // still act as the minimum padding when the normalized mantissa is
    // shorter than that group.
    const int integer_group =
        std::max(1, section.integer_zero_digits + section.integer_opt_digits + section.integer_pad_digits);
    double mantissa = scaled;
    if (mantissa != 0.0) {
      const double abs_m = std::fabs(mantissa);
      const int raw_exponent = static_cast<int>(std::floor(std::log10(abs_m)));
      // C++ integer division truncates toward zero. Scientific notation needs
      // floor division so values in (-1, 0) stay in the preceding group.
      int quotient = raw_exponent / integer_group;
      if (raw_exponent < 0 && raw_exponent % integer_group != 0) {
        --quotient;
      }
      exponent = quotient * integer_group;
      if (exponent > 0) {
        for (int i = 0; i < exponent; ++i) {
          mantissa /= 10.0;
        }
      } else if (exponent < 0) {
        for (int i = 0; i > exponent; --i) {
          mantissa *= 10.0;
        }
      }
      if (!std::isfinite(mantissa)) {
        return FormatStatus::kOverflow;
      }
    }
    format_fixed_digits(mantissa, frac_digits, &negative, &int_digits, &frac_digits_str);

    // Rounding can promote 99.999 to 100.00 (or 9.999 to 10.00). Carry the
    // engineering exponent by one whole group and format the renormalized
    // mantissa again so the integer side stays within its group.
    if (mantissa != 0.0 && static_cast<int>(int_digits.size()) > integer_group) {
      for (int i = 0; i < integer_group; ++i) {
        mantissa /= 10.0;
      }
      exponent += integer_group;
      if (!std::isfinite(mantissa)) {
        return FormatStatus::kOverflow;
      }
      format_fixed_digits(mantissa, frac_digits, &negative, &int_digits, &frac_digits_str);
    }
  } else {
    format_fixed_digits(scaled, frac_digits, &negative, &int_digits, &frac_digits_str);
  }

  // `General` chooses its own display precision below.  Its sign must
  // therefore come from the unrounded input, not from the zero-decimal
  // placeholder pass above: e.g. TEXT(-1/3, "General") is -0.333333333,
  // not 0.333333333.
  if (section.has_general) {
    negative = scaled < 0.0;
  }

  // Pad integer digits to the required minimum imposed by `0` placeholders.
  // `?` reserves a visual position but must render a space rather than a zero
  // when the value has fewer integer digits (e.g. `TEXT(5,"?0")` -> ` 5`).
  const int int_min = section.integer_zero_digits;
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

  // Locate the decimal point and scientific marker inside the token stream so
  // we can partition the digit tokens into integer / fraction / exponent
  // stacks. This mirrors the classifier, but we need the positions again at
  // walk time to drive the interleaving.
  int point_index = -1;
  int scientific_index = -1;
  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    if (point_index < 0 && section.tokens[i].kind == Tok::Point) {
      point_index = static_cast<int>(i);
    }
    if (scientific_index < 0 && (section.tokens[i].kind == Tok::SciPlus || section.tokens[i].kind == Tok::SciMinus)) {
      scientific_index = static_cast<int>(i);
    }
  }
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
    if (scientific_index >= 0 && static_cast<int>(i) >= scientific_index) {
      break;
    }
    if (is_digit_tok(section.tokens[i].kind)) {
      int_digit_positions.push_back(i);
    }
  }

  // Scientific exponent placeholders consume the prepared exponent digits
  // independently of the mantissa's integer/fraction streams. Keep the same
  // right-aligned slot behavior as the common numeric walker so literals,
  // DBNum substitution, and `_X` spaces remain in their original positions.
  std::vector<std::size_t> exponent_digit_positions;
  std::vector<std::string> exponent_slot_text;
  std::string exponent_digits;
  if (section.has_scientific) {
    const int abs_exponent = exponent < 0 ? -exponent : exponent;
    exponent_digits = std::to_string(abs_exponent);
    bool after_marker = false;
    for (std::size_t i = 0; i < section.tokens.size(); ++i) {
      const Tok kind = section.tokens[i].kind;
      if (kind == Tok::SciPlus || kind == Tok::SciMinus) {
        after_marker = true;
        continue;
      }
      if (after_marker && is_digit_tok(kind)) {
        exponent_digit_positions.push_back(i);
      }
    }
    exponent_slot_text.resize(exponent_digit_positions.size());
    const std::size_t n_exponent_tokens = exponent_digit_positions.size();
    const std::size_t n_exponent_digits = exponent_digits.size();
    if (n_exponent_tokens > 0 && n_exponent_digits > 0) {
      if (n_exponent_digits >= n_exponent_tokens) {
        const std::size_t prefix_len = n_exponent_digits - n_exponent_tokens + 1;
        exponent_slot_text[0].assign(exponent_digits, 0, prefix_len);
        for (std::size_t i = 1; i < n_exponent_tokens; ++i) {
          exponent_slot_text[i].assign(1, exponent_digits[prefix_len + i - 1]);
        }
      } else {
        const std::size_t fallback_count = n_exponent_tokens - n_exponent_digits;
        for (std::size_t i = 0; i < n_exponent_digits; ++i) {
          exponent_slot_text[fallback_count + i].assign(1, exponent_digits[i]);
        }
      }
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
    if (decimal_digits_all_zero(int_digits) && decimal_digits_all_zero(frac_digits_str)) {
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
  std::size_t exponent_token_cursor = 0;
  bool past_point = false;
  bool in_exponent = false;

  auto emit_int_digit_char = [&](char digit) {
    if (section.thousands_separator && int_cursor > 0 && (n_int_digits - int_cursor) % 3 == 0) {
      result.push_back(',');
    }
    append_digit_dbnum(result, section.dbnum_mode, digit);
    ++int_cursor;
  };

  // Helper: emit a single fractional digit with DBNum substitution.
  auto emit_frac_digit_char = [&](char digit) { append_digit_dbnum(result, section.dbnum_mode, digit); };

  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    const Token& tk = section.tokens[i];
    switch (tk.kind) {
      case Tok::DigitZero:
      case Tok::DigitOpt:
      case Tok::DigitPad:
        if (in_exponent) {
          const std::string* slot = nullptr;
          if (exponent_token_cursor < exponent_slot_text.size()) {
            slot = &exponent_slot_text[exponent_token_cursor];
          }
          if (slot != nullptr && !slot->empty()) {
            for (char d : *slot) {
              append_digit_dbnum(result, section.dbnum_mode, d);
            }
          } else if (tk.kind == Tok::DigitZero) {
            append_digit_dbnum(result, section.dbnum_mode, '0');
          } else if (tk.kind == Tok::DigitPad) {
            result.push_back(' ');
          }
          ++exponent_token_cursor;
        } else if (past_point) {
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
        // `[DBNum1]` spells the integer part positionally; `[DBNum2-3]`
        // substitute digit by digit, and non-digit bytes (the decimal
        // point, exponent marker, ...) pass through unchanged.
        if (section.dbnum_mode == DbNumMode::kDBNum1) {
          append_dbnum1_general(result, std::fabs(value));
          break;
        }
        std::string general;
        format_general(general, std::fabs(value));
        append_chars_dbnum(result, section.dbnum_mode, general);
        break;
      }
      case Tok::SciPlus:
      case Tok::SciMinus:
        // The mantissa was prepared above; the rest of the token stream is
        // still walked here so exponent placeholders and surrounding
        // literals retain their format positions.
        result.push_back('E');
        if (exponent >= 0) {
          if (tk.kind == Tok::SciPlus) {
            result.push_back('+');
          }
        } else {
          result.push_back('-');
        }
        in_exponent = true;
        break;
      case Tok::At:
        // Text-only sections are dispatched before numeric rendering.
        break;
      default:
        break;
    }
  }
  out.append(result);
  return FormatStatus::kOk;
}

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
