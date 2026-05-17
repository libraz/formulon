// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Fraction-format rendering for the Excel TEXT() engine. Covers `# ?/?`,
// `# ??/??`, `0/0`, and friends. Implements Excel's bounded
// best-rational-approximation via a Stern-Brocot mediant search.
//
// Compared to the more familiar continued-fraction algorithm, Stern-Brocot
// occasionally picks a different mediant when the target lies near a
// Farey-neighbour boundary; Mac Excel's empirical output matches
// Stern-Brocot, so we follow that.
//
// For a value `v` (already absolute) and a denominator cap `max_q` and
// numerator cap `max_p`, the search finds (p, q) minimising |v - p/q|
// with 1 <= q <= max_q and 0 <= p <= max_p. When an integer group is
// present the search runs against `frac = v - floor(v)` only.

#include "eval/text_format/render_fraction.h"

#include <cmath>
#include <cstddef>
#include <string>
#include <string_view>

#include "eval/text_format/number_format_types.h"
#include "eval/text_format/render_common.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
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
    append_chars_dbnum(out, section.dbnum_mode, digits);
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
        append_digit_dbnum(out, section.dbnum_mode, '0');
      } else if (kind == Tok::DigitPad) {
        out.push_back(' ');
      }
      // `#`: emit nothing.
    } else {
      const char d = digits[digit_cursor++];
      append_digit_dbnum(out, section.dbnum_mode, d);
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

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
