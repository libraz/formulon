//
// Fraction-format rendering for the Excel TEXT() engine. Variable
// denominators use a bounded continued-fraction search; fixed denominators
// round directly. Numerator placeholder width controls padding, not magnitude.

#include "eval/text_format/render_fraction.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>

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

std::uint64_t rounded_product(double target, std::uint64_t denominator) noexcept {
  const double product = target * static_cast<double>(denominator);
  if (!(product > 0.0)) {
    return 0;
  }
  const double rounded = std::floor(product + 0.5);
  if (rounded >= static_cast<double>(denominator)) {
    return denominator;
  }
  return static_cast<std::uint64_t>(rounded);
}

struct RationalCandidate {
  std::uint64_t numerator = 0;
  std::uint64_t denominator = 1;
  double error = std::numeric_limits<double>::max();
};

// All arithmetic is binary64, as in Excel: a wider long double (x86-64,
// aarch64 Linux, WASM) rounds 0.015 * 100 below 1.5 and picks other candidates.
void consider_candidate(double target, std::uint64_t numerator, std::uint64_t denominator,
                        RationalCandidate* best) noexcept {
  if (denominator == 0) {
    return;
  }
  const double approximation = static_cast<double>(numerator) / static_cast<double>(denominator);
  const double error = std::fabs(target - approximation);
  // Excel treats adjacent binary64 values at a rational midpoint as ties
  // (17/144 with a one-digit denominator chooses 1/8, not 1/9). Use the
  // input's precision.
  // Scale the bound by approximation error, so an exact rational never ties
  // with a non-exact one and tiny fractions remain distinguishable from zero.
  const double tie_bound = 32.0 * std::numeric_limits<double>::epsilon() * std::max(error, best->error);
  const bool tied = std::fabs(error - best->error) <= tie_bound;
  if ((!tied && error < best->error) || (tied && (denominator < best->denominator ||
                                                  (denominator == best->denominator && numerator < best->numerator)))) {
    best->numerator = numerator;
    best->denominator = denominator;
    best->error = error;
  }
}

// Bounded continued-fraction search. It keeps the convergents in an integer
// denominator and jumps over large partial quotients, avoiding the linear
// Stern-Brocot walk that made values such as 1/20000 hit its iteration cap.
void best_rational(double target, std::uint64_t max_q, std::uint64_t* out_num, std::uint64_t* out_den) noexcept {
  max_q = std::min(max_q, kMaxVariableDenominator);
  if (max_q < 1) {
    max_q = 1;
  }
  RationalCandidate best;

  auto consider_boundary = [&](std::uint64_t denominator) noexcept {
    if (denominator > 0) {
      consider_candidate(target, rounded_product(target, denominator), denominator, &best);
    }
  };

  // A denominator-one endpoint is always a useful initial candidate,
  // including tiny fractional parts whose first continued-fraction term
  // exceeds the denominator range.
  consider_boundary(1);

  // A wider expansion would turn 1/0.1 into 9.99..., not 10, and pick another term.
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
    if (q_prev2 > max_q) {
      consider_boundary(max_q);
      break;
    }
    const std::uint64_t quotient_limit = q_prev1 == 0 ? max_q : (max_q - q_prev2) / q_prev1;
    // Compare before casting: a_value can exceed UINT64_MAX. The limit is below 2^53, so the bound is exact.
    if (a_value >= static_cast<double>(quotient_limit) + 1.0) {
      // When even the first reciprocal is out of range Excel returns zero, not 1/max_q.
      if (p_prev1 == 0 && q_prev1 == 1) {
        break;
      }
      const std::uint64_t t = quotient_limit;
      if (t > 0) {
        consider_candidate(target, p_prev2 + p_prev1 * t, q_prev2 + q_prev1 * t, &best);
      }
      consider_candidate(target, p_prev1, q_prev1, &best);
      consider_boundary(max_q);
      break;
    }
    const std::uint64_t a = static_cast<std::uint64_t>(a_value);

    if (a != 0 && p_prev1 > (std::numeric_limits<std::uint64_t>::max() - p_prev2) / a) {
      consider_boundary(max_q);
      break;
    }
    const std::uint64_t p = p_prev2 + p_prev1 * a;
    const std::uint64_t q = q_prev2 + q_prev1 * a;
    if (q > max_q || q == 0) {
      consider_boundary(max_q);
      break;
    }
    consider_candidate(target, p, q, &best);
    const double fractional = x - a_value;
    if (fractional == 0.0) {
      break;
    }
    p_prev2 = p_prev1;
    p_prev1 = p;
    q_prev2 = q_prev1;
    q_prev1 = q;
    x = 1.0 / fractional;
  }

  *out_num = best.numerator;
  *out_den = best.denominator < 1 ? 1 : best.denominator;
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

// Emit the non-negative integer `digits` right-aligned to `width` characters
// using the placeholder kinds in `[begin, end)`. Each placeholder kind
// determines how unused leading positions render: `0` -> '0' pad, `?` ->
// space pad, `#` -> nothing emitted. The DBNum mapping applies to digits
// (and to `0`-pad positions) but never to spaces or absent positions.
void emit_fraction_digits(const Section& section, std::string_view digits, int begin, int end, std::string& out,
                          bool trailing_pad = false) {
  const int width = end - begin;
  if (width <= 0) {
    return;
  }
  if (static_cast<int>(digits.size()) > width) {
    // Overflow: Excel never produces this for our caps, but be defensive.
    // Emit the digits verbatim (the widest available run cap is enforced
    // by the search above so this case is essentially unreachable).
    append_chars_dbnum(out, section.digit_style, digits);
    return;
  }
  const int pad = width - static_cast<int>(digits.size());
  // Iterate placeholder kinds left-to-right. The first `pad` placeholders
  // are leading positions with no digit; the remaining `digits.size()`
  // placeholders consume the digit string in order.
  std::size_t digit_cursor = 0;
  for (int k = 0; k < width; ++k) {
    const Tok kind = section.tokens[static_cast<std::size_t>(begin + k)].kind;
    const bool is_padding = trailing_pad ? k >= static_cast<int>(digits.size()) : k < pad;
    if (is_padding) {
      // Leading-position behaviour by placeholder kind.
      if (kind == Tok::DigitZero) {
        append_digit_dbnum(out, section.digit_style, '0');
      } else if (kind == Tok::DigitPad) {
        out.push_back(' ');
      }
      // `#`: emit nothing.
    } else {
      const char d = trailing_pad ? digits[static_cast<std::size_t>(k)] : digits[digit_cursor++];
      append_digit_dbnum(out, section.digit_style, d);
    }
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
  double integer_part = 0.0;
  double target = abs_v;
  integer_part = std::floor(abs_v);
  target = abs_v - integer_part;

  std::uint64_t fraction_num = 0;
  std::uint64_t den = 1;
  if (section.fraction_fixed_denominator) {
    den = section.fraction_fixed_denominator_value;
    fraction_num = rounded_product(target, den);
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
  // A whole value shows no fraction, but keeps its width: the separator,
  // numerator, slash and denominator each become blanks (`59    ` for
  // `# ?/?`), and a zero integer is shown even under `#`.
  const bool suppress_fraction = has_int_group && fraction_num == 0;
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
    // Excel suppresses the integer when it is zero AND the leading
    // placeholder is `#`; otherwise the leading-pad behaviour from
    // `emit_fraction_digits` handles `0` and `?` correctly.
    const Tok lead_kind = section.tokens[static_cast<std::size_t>(int_begin)].kind;
    if (integer_part == 0.0 && lead_kind == Tok::DigitOpt && !suppress_fraction) {
      // Emit nothing for the integer group.
    } else {
      emit_fraction_digits(section, int_digits, int_begin, int_end, result);
    }
    i = static_cast<std::size_t>(int_end);
    if (suppress_fraction) {
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
    } else {
      // 3) Literals between integer group and numerator group (typically a
      // single space).
      while (i < static_cast<std::size_t>(num_begin)) {
        emit_token_verbatim(i);
        ++i;
      }
    }
  }
  if (suppress_fraction) {
    while (i < n_tokens) {
      emit_token_verbatim(i);
      ++i;
    }
    out.append(result);
    return FormatStatus::kOk;
  }
  // 4) Numerator group.
  emit_fraction_digits(section, numerator_digits, num_begin, num_end, result);
  i = static_cast<std::size_t>(num_end);
  // 5) Literals up to the slash (including the slash itself).
  while (i <= static_cast<std::size_t>(slash_index)) {
    emit_token_verbatim(i);
    ++i;
  }
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
