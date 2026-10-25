//
// Fraction-format rendering for the Excel TEXT() engine. Variable
// denominators use a bounded continued-fraction search; fixed denominators
// round directly. Numerator placeholder width controls padding, not magnitude.

#include "eval/text_format/render_fraction.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/text_format/number_format_scanner.h"
#include "eval/text_format/number_format_types.h"
#include "eval/text_format/render_common.h"
#include "utils/number_text.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
namespace {

// Excel caps variable-denominator approximation at seven digits even when
// the format has more placeholders. The improper numerator is a decimal
// string: a finite double's whole part can run to 309 digits.
constexpr std::uint64_t kMaxVariableDenominator = 9999999ULL;

// Compute 10^N without overflow for the supported placeholder widths.
// best_rational applies Excel's seven-digit variable-denominator cap; wider
// placeholder runs retain their layout.
std::uint64_t fraction_pow10(int n) noexcept {
  std::uint64_t r = 1;
  for (int i = 0; i < n && i < 18; ++i) {
    r *= 10;
  }
  return r;
}

// Excel takes the last convergent of the binary64 continued-fraction expansion
// whose denominator fits, never a semiconvergent, even a closer one: 0.5627
// gives 4/7 rather than 5/9, and 0.6125 gives 3/5 rather than 5/8.
void best_rational(double target, std::uint64_t max_q, std::uint64_t* out_num, std::uint64_t* out_den) noexcept {
  max_q = std::min(max_q, kMaxVariableDenominator);
  if (max_q < 1) {
    max_q = 1;
  }
  std::uint64_t num = 0;
  std::uint64_t den = 1;
  double x = target;
  std::uint64_t p_prev2 = 0;
  std::uint64_t p_prev1 = 1;
  std::uint64_t q_prev2 = 1;
  std::uint64_t q_prev1 = 0;
  for (int iter = 0; iter < 256 && std::isfinite(x); ++iter) {
    const double a_value = std::floor(x);
    if (a_value < 0.0) {
      break;
    }
    // Compare before casting: a_value can exceed UINT64_MAX. The limit is below 2^53, so the bound is exact.
    if (q_prev1 != 0 && a_value > static_cast<double>((max_q - q_prev2) / q_prev1)) {
      break;
    }
    const std::uint64_t a = static_cast<std::uint64_t>(a_value);
    num = p_prev2 + p_prev1 * a;
    den = q_prev2 + q_prev1 * a;
    const double fractional = x - a_value;
    if (fractional == 0.0) {
      break;
    }
    p_prev2 = p_prev1;
    p_prev1 = num;
    q_prev2 = q_prev1;
    q_prev1 = den;
    x = 1.0 / fractional;
  }
  *out_num = num;
  *out_den = den;
}

// Multiply a decimal digit string by an 18-digit-or-smaller denominator. The
// per-digit product plus carry is at most 9e18 + 1e18, safely below UINT64_MAX.
std::string multiply_decimal(std::string_view digits, std::uint64_t multiplier) {
  if (multiplier == 0 || digits.empty() || (digits.size() == 1 && digits[0] == '0')) {
    return "0";
  }
  std::string reversed;
  reversed.reserve(digits.size() + 20);
  std::uint64_t carry = 0;
  for (std::size_t i = digits.size(); i > 0; --i) {
    const std::uint64_t digit = static_cast<std::uint64_t>(digits[i - 1] - '0');
    const std::uint64_t product = digit * multiplier + carry;
    reversed.push_back(static_cast<char>('0' + product % 10U));
    carry = product / 10U;
  }
  while (carry != 0) {
    reversed.push_back(static_cast<char>('0' + carry % 10U));
    carry /= 10U;
  }
  std::reverse(reversed.begin(), reversed.end());
  return reversed;
}

std::string add_decimal(std::string_view digits, std::uint64_t addend) {
  const std::string addend_digits = std::to_string(addend);
  std::string reversed;
  reversed.reserve(std::max(digits.size(), addend_digits.size()) + 1);
  std::size_t i = digits.size();
  std::size_t j = addend_digits.size();
  std::uint64_t carry = 0;
  while (i > 0 || j > 0 || carry != 0) {
    std::uint64_t sum = carry;
    if (i > 0) {
      sum += static_cast<std::uint64_t>(digits[--i] - '0');
    }
    if (j > 0) {
      sum += static_cast<std::uint64_t>(addend_digits[--j] - '0');
    }
    reversed.push_back(static_cast<char>('0' + sum % 10U));
    carry = sum / 10U;
  }
  if (reversed.empty()) {
    reversed.push_back('0');
  }
  std::reverse(reversed.begin(), reversed.end());
  return reversed;
}

bool is_digit_placeholder_tok(Tok k) {
  return k == Tok::DigitZero || k == Tok::DigitOpt || k == Tok::DigitPad;
}

// Emit the non-negative integer `digits` right-aligned to `width` characters
// using the placeholder kinds in `[begin, end)`. Each placeholder kind
// determines how unused leading positions render: `0` -> '0' pad, `?` ->
// space pad, `#` -> nothing emitted. The DBNum mapping applies to digits
// (and to `0`-pad positions) but never to spaces or absent positions.
void emit_fraction_digits(const Section& section, std::string_view digits, int begin, int end, std::string& out,
                          bool trailing_pad = false) {
  std::vector<Tok> kinds;
  bool grouped = false;
  for (int k = begin; k < end; ++k) {
    const Tok kind = section.tokens[static_cast<std::size_t>(k)].kind;
    if (kind == Tok::Comma) {
      grouped = true;
    } else {
      kinds.push_back(kind);
    }
  }
  const int width = static_cast<int>(kinds.size());
  if (width <= 0) {
    return;
  }
  // One piece per emitted position; `is_digit` marks the positions a group separator counts.
  std::vector<std::string> pieces;
  std::vector<bool> is_digit;
  const auto add_digit = [&](char d) {
    pieces.emplace_back();
    append_digit_dbnum(pieces.back(), section.digit_style, d);
    is_digit.push_back(true);
  };
  if (static_cast<int>(digits.size()) > width) {
    // Overflow: the digits show in full (the search above caps their count).
    for (const char d : digits) {
      add_digit(d);
    }
  } else {
    const int pad = width - static_cast<int>(digits.size());
    // The first `pad` placeholders are leading positions with no digit; the rest consume `digits` in order.
    // A trailing `0` after an optional `#` pads nothing.
    bool zero_pad_silent = false;
    if (trailing_pad) {
      for (int k = 0; k < width; ++k) {
        zero_pad_silent = zero_pad_silent || kinds[static_cast<std::size_t>(k)] == Tok::DigitOpt;
      }
    }
    std::size_t digit_cursor = 0;
    for (int k = 0; k < width; ++k) {
      const Tok kind = kinds[static_cast<std::size_t>(k)];
      const bool is_padding = trailing_pad ? k >= static_cast<int>(digits.size()) : k < pad;
      if (is_padding) {
        if (kind == Tok::DigitZero) {
          if (!zero_pad_silent) {
            add_digit('0');
          }
        } else if (kind == Tok::DigitPad) {
          pieces.emplace_back(" ");
          is_digit.push_back(false);
        }
        // `#`: emit nothing.
      } else {
        add_digit(trailing_pad ? digits[static_cast<std::size_t>(k)] : digits[digit_cursor++]);
      }
    }
  }
  // A separator between integer placeholders groups every three digits from the right.
  std::size_t digits_right = 0;
  std::vector<std::string> reversed;
  for (std::size_t k = pieces.size(); k-- > 0;) {
    reversed.push_back(std::move(pieces[k]));
    if (grouped && is_digit[k]) {
      ++digits_right;
      if (digits_right % 3 == 0 && k > 0 && is_digit[k - 1]) {
        reversed.emplace_back(",");
      }
    }
  }
  for (auto it = reversed.rbegin(); it != reversed.rend(); ++it) {
    out += *it;
  }
}

}  // namespace

FormatStatus render_fraction(const Section& section, std::string_view fmt, double value, std::string& out) {
  // Sign: emit minus prefix on the rendered form for negative values, then
  // operate on the absolute magnitude. Excel's fraction format does not
  // honour `?`/`#` for sign placement -- the leading `-` is unconditional.
  const bool negative = value < 0.0;
  const double abs_v = std::fabs(value);

  const bool has_int_group = section.fraction_int_max_digits > 0;

  // Split the magnitude before approximation. The bounded search only needs
  // the fractional part, while an improper numerator is assembled as
  // `whole * denominator + fraction_numerator` below using decimal strings.
  // That preserves values whose whole part is wider than uint64_t.
  double integer_part = std::floor(abs_v);
  const double target = abs_v - integer_part;

  std::uint64_t fraction_num = 0;
  std::uint64_t den = 1;
  if (section.fraction_fixed_denominator) {
    // A fixed denominator rounds `value * denominator` half away from zero.
    den = section.fraction_fixed_denominator_value;
    const double scaled = std::round(abs_v * static_cast<double>(den)) - integer_part * static_cast<double>(den);
    fraction_num = scaled > 0.0 ? static_cast<std::uint64_t>(scaled) : 0U;
  } else {
    // The numerator is uncapped so improper forms such as `?/?` render `2469/2`.
    const std::uint64_t max_q = fraction_pow10(section.fraction_den_max_digits) - 1U;
    best_rational(target, max_q, &fraction_num, &den);
  }

  // Rounding promotion: if the best approximation rounds up to 1 exactly
  // (num == den) and we have an integer group, increment the integer and
  // zero out the fraction.
  if (has_int_group && fraction_num == den) {
    integer_part += 1.0;
    if (!std::isfinite(integer_part)) {
      return FormatStatus::kOverflow;
    }
    fraction_num = 0;
    den = 1;
  }
  const int int_end = section.fraction_int_end;
  const int num_begin = section.fraction_num_begin;
  const int num_end = section.fraction_num_end;
  const int den_begin = section.fraction_den_begin;
  const int den_end = section.fraction_den_end;
  const int slash_index = section.fraction_slash_index;
  const auto kind_at = [&](int index) { return section.tokens[static_cast<std::size_t>(index)].kind; };
  // A group separator between integer placeholders is inert; the placeholders on both sides all count.
  int int_begin = section.fraction_int_begin;
  while (has_int_group && int_begin >= 2 && kind_at(int_begin - 1) == Tok::Comma &&
         kind_at(int_begin - 2) != Tok::Comma && is_digit_placeholder_tok(kind_at(int_begin - 2))) {
    int_begin -= 2;
    while (int_begin > 0 && is_digit_placeholder_tok(kind_at(int_begin - 1))) {
      --int_begin;
    }
  }
  const bool literal_numerator = kind_at(num_begin) == Tok::Literal;
  bool zero_placeholder_numerator = false;
  for (int k = num_begin; k < num_end; ++k) {
    zero_placeholder_numerator = zero_placeholder_numerator || kind_at(k) == Tok::DigitZero;
  }
  bool pad_placeholder = false;
  for (int k = has_int_group ? int_begin : num_begin; k < den_end; ++k) {
    pad_placeholder = pad_placeholder || kind_at(k) == Tok::DigitPad;
  }
  // A whole value hides its fraction: with a `0` numerator placeholder it still
  // reads `0/1`; with a `?` anywhere the group keeps its width in blanks; else
  // the group, separator included, is dropped.
  const bool whole_value = has_int_group && fraction_num == 0 && !zero_placeholder_numerator;
  const bool blank_fraction = whole_value && pad_placeholder;
  const bool drop_fraction = whole_value && !pad_placeholder;
  char int_buf[400];
  const int int_len = format_fixed(int_buf, sizeof(int_buf), integer_part, 0);
  if (int_len < 0 || static_cast<std::size_t>(int_len) >= sizeof(int_buf)) {
    return FormatStatus::kOverflow;
  }
  std::string int_digits(int_buf, int_len > 0 ? static_cast<std::size_t>(int_len) : 0U);
  cap_integer_significant_digits(&int_digits);
  const std::string numerator_digits =
      has_int_group ? std::to_string(fraction_num) : add_decimal(multiply_decimal(int_digits, den), fraction_num);
  const std::string denominator_digits = std::to_string(den);

  std::string result;
  if (negative) {
    result.push_back('-');
  }

  // Walk the section's token stream: literals before the integer group (or the
  // numerator when there is none), the integer group, the literals up to the
  // numerator, the numerator group, the slash, the denominator group, then the
  // trailing literals.
  const std::size_t n_tokens = section.tokens.size();

  // Emits a single token verbatim (literals only; other tokens inside a
  // fraction section are skipped).
  auto emit_token_verbatim = [&](std::size_t idx) {
    const Token& tk = section.tokens[idx];
    if (tk.kind == Tok::Literal && tk.lit_end > tk.lit_begin) {
      result.append(fmt.data() + tk.lit_begin, tk.lit_end - tk.lit_begin);
    } else if (tk.kind == Tok::Space) {
      result.push_back(' ');
    } else if (tk.kind == Tok::Percent) {
      append_percent(result, fmt, tk);
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
    // A zero integer under a visible fraction takes only the placeholders'
    // padding (`0` -> '0', `?` -> blank, `#` -> nothing); a whole value shows
    // its integer, zero included.
    const bool fraction_visible = !blank_fraction && !drop_fraction;
    const std::string_view shown_digits = integer_part == 0.0 && fraction_visible ? "" : std::string_view(int_digits);
    emit_fraction_digits(section, shown_digits, int_begin, int_end, result);
    i = static_cast<std::size_t>(int_end);
    if (blank_fraction) {
      for (; i < static_cast<std::size_t>(den_end); ++i) {
        const Token& tk = section.tokens[i];
        if (tk.kind == Tok::Literal) {
          for (std::size_t k = tk.lit_begin; k < tk.lit_end; k += utf8_scalar_width(fmt, k)) {
            result.push_back(' ');
          }
        } else {
          result.push_back(' ');
        }
      }
    } else if (drop_fraction) {
      i = static_cast<std::size_t>(den_end);
    } else {
      // 3) Literals between the integer group and the numerator (typically a
      // single space). An empty integer before an optional numerator digit
      // takes no separator.
      bool int_all_optional = true;
      for (int k = int_begin; k < int_end; ++k) {
        int_all_optional = int_all_optional && (kind_at(k) == Tok::DigitOpt || kind_at(k) == Tok::Comma);
      }
      const bool skip_separator = integer_part == 0.0 && int_all_optional && kind_at(num_begin) == Tok::DigitOpt;
      while (i < static_cast<std::size_t>(num_begin)) {
        if (!skip_separator) {
          emit_token_verbatim(i);
        }
        ++i;
      }
    }
  }
  if (blank_fraction || drop_fraction) {
    while (i < n_tokens) {
      emit_token_verbatim(i);
      ++i;
    }
    out.append(result);
    return FormatStatus::kOk;
  }
  // 4) Numerator group: literal digits print as written.
  if (literal_numerator) {
    for (int k = num_begin; k < num_end; ++k) {
      emit_token_verbatim(static_cast<std::size_t>(k));
    }
  } else {
    emit_fraction_digits(section, numerator_digits, num_begin, num_end, result);
  }
  i = static_cast<std::size_t>(num_end);
  // 5) Literals up to the slash; the slash itself writes `/` even when spelled `／`.
  while (i < static_cast<std::size_t>(slash_index)) {
    emit_token_verbatim(i);
    ++i;
  }
  result.push_back('/');
  ++i;
  // 6) Denominator group.
  if (section.fraction_fixed_denominator) {
    // A zero-prefixed fixed denominator shows zeros for its significant width, then spaces (`/008` -> `0  `).
    const bool has_zero_prefix =
        den_begin < den_end && section.tokens[static_cast<std::size_t>(den_begin)].kind == Tok::DigitZero;
    if (!has_zero_prefix) {
      result.append(denominator_digits);
    } else {
      for (std::size_t digit = 0; digit < denominator_digits.size(); ++digit) {
        append_digit_dbnum(result, section.digit_style, '0');
      }
      result.append(static_cast<std::size_t>(den_end - den_begin) - denominator_digits.size(), ' ');
    }
  } else {
    // Runs starting with `0` right-align, others left-align (`?0` -> `20`, `0?` -> `02`).
    const bool denominator_trailing_pad =
        den_begin < den_end && section.tokens[static_cast<std::size_t>(den_begin)].kind != Tok::DigitZero;
    emit_fraction_digits(section, denominator_digits, den_begin, den_end, result, denominator_trailing_pad);
  }
  i = static_cast<std::size_t>(den_end);
  // 7) Trailing literals.
  while (i < n_tokens) {
    emit_token_verbatim(i);
    ++i;
  }

  out.append(result);
  return FormatStatus::kOk;
}

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
