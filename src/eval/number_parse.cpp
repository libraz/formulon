//
// Implementation of the shared locale-aware numeric parsers declared in
// `number_parse.h`. Extracted verbatim from the VALUE / NUMBERVALUE
// builtins so the same normalisation drives implicit text->number coercion
// (arithmetic operators and the criteria engine).

#include "eval/number_parse.h"

#include <cmath>
#include <cstddef>
#include <cstdint>

#include "eval/eval_profile_scope.h"
#include "excel_locale.h"
#include "utils/double_parse.h"

namespace formulon {
namespace eval {
namespace {

bool is_ascii_ws(unsigned char c) noexcept {
  return c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '\v' || c == '\f';
}

std::string_view trim_ascii(std::string_view s) noexcept {
  while (!s.empty() && is_ascii_ws(static_cast<unsigned char>(s.front()))) {
    s.remove_prefix(1);
  }
  while (!s.empty() && is_ascii_ws(static_cast<unsigned char>(s.back()))) {
    s.remove_suffix(1);
  }
  return s;
}

// Strips a leading currency prefix from the active locale's accepted set.
// Returns the input unchanged when no accepted symbol is present.
std::string_view strip_currency(std::string_view s) noexcept {
  for (const std::string_view symbol : locale_facts(current_eval_profile()).accepted_currency) {
    if (!symbol.empty() && s.substr(0, symbol.size()) == symbol) {
      return s.substr(symbol.size());
    }
  }
  return s;
}

// Strips a trailing Euro suffix (`23€` -> `23`). Mac Excel 365 accepts Euro
// both as prefix and suffix; no other symbol is accepted as a suffix. A
// suffix-currency locale also allows the blank before the symbol.
std::string_view strip_trailing_euro(std::string_view s) noexcept {
  constexpr std::string_view kEuro = "\xE2\x82\xAC";
  if (s.size() < kEuro.size() || s.substr(s.size() - kEuro.size()) != kEuro) {
    return s;
  }
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  bool accepted = false;
  for (const std::string_view symbol : facts.accepted_currency) {
    accepted = accepted || symbol == kEuro;
  }
  if (!accepted) {
    return s;
  }
  s.remove_suffix(kEuro.size());
  return facts.currency.suffix ? trim_ascii(s) : s;
}

}  // namespace

bool parse_numeric(std::string_view s, char decimal_sep, char group_sep, double* out, bool strict_groups) noexcept {
  // Trim ASCII whitespace first.
  s = trim_ascii(s);
  if (s.empty()) {
    return false;
  }
  // Sign and currency may appear in either order at the front, and Excel
  // accepts a currency symbol on the leading OR trailing side but NOT both:
  //   "-$100" / "$-100" / "$ 100" -> ok;  "$100" / "€100" / "100€" -> ok;
  //   "$100€" -> #VALUE! (currency on both ends).
  bool negative = false;
  auto try_sign = [&negative, &s]() -> bool {
    if (!s.empty() && (s.front() == '+' || s.front() == '-')) {
      negative = s.front() == '-';
      s.remove_prefix(1);
      return true;
    }
    return false;
  };
  const bool had_leading_sign = try_sign();
  bool leading_currency = false;
  {
    const std::string_view after = strip_currency(s);
    if (after.size() != s.size()) {
      s = after;
      leading_currency = true;
    }
  }
  if (leading_currency) {
    // A space between the currency symbol and the number is allowed
    // ("$ 100"), and the sign may follow the symbol ("$-100").
    while (!s.empty() && is_ascii_ws(static_cast<unsigned char>(s.front()))) {
      s.remove_prefix(1);
    }
    if (!had_leading_sign) {
      try_sign();
    }
  }
  if (s.empty()) {
    return false;
  }
  // A trailing Euro suffix is a single-sided currency: strip it only when no
  // leading currency was consumed, so "$100€" (currency on both ends) keeps a
  // stray `€` that the numeric scan below rejects. Excel accepts `€` as a
  // suffix (e.g. `"23€"`); `$` and `円` are not accepted as suffixes.
  if (!leading_currency) {
    s = strip_trailing_euro(s);
    if (s.empty()) {
      return false;
    }
  }
  // Trailing percent signs. Each `%` multiplies the parsed value by 0.01,
  // so `"50%%"` yields `0.005` (matching Excel's NUMBERVALUE behavior).
  int percent_count = 0;
  while (!s.empty() && s.back() == '%') {
    ++percent_count;
    s.remove_suffix(1);
  }
  if (s.empty()) {
    return false;
  }
  // Scan and assemble a canonical C-locale numeric string (digits, one
  // optional `.`, optional exponent `e[+/-]digits`). Reject on any
  // unexpected byte.
  std::string canonical;
  canonical.reserve(s.size());
  bool seen_digit = false;
  bool seen_point = false;
  bool seen_exp = false;
  // Thousands-grouping validation state. Mac Excel rejects malformed
  // groupings such as `"12,34"` (2 digits before, 2 after) or `"1,2345"`
  // (final group not exactly 3 digits). The first group (before the first
  // separator) must be 1-3 digits; every subsequent group must be exactly 3.
  bool seen_group_sep = false;
  int digits_in_current_group = 0;
  for (std::size_t i = 0; i < s.size(); ++i) {
    const char c = s[i];
    if (c >= '0' && c <= '9') {
      canonical.push_back(c);
      seen_digit = true;
      if (!seen_point && !seen_exp) {
        ++digits_in_current_group;
      }
      continue;
    }
    if (c == decimal_sep && !seen_point && !seen_exp) {
      // Transitioning out of integer part: validate the final integer group
      // if any group separators were seen.
      if (seen_group_sep && digits_in_current_group != 3) {
        return false;
      }
      canonical.push_back('.');
      seen_point = true;
      continue;
    }
    if (group_sep != '\0' && c == group_sep && !seen_point && !seen_exp && !strict_groups) {
      continue;
    }
    if (group_sep != '\0' && c == group_sep && !seen_point && !seen_exp) {
      // Group separator inside the integer part. Validate the just-finished
      // group: 1-3 digits for the first one, exactly 3 for any subsequent.
      // `group_sep == '\0'` means the caller opted out of grouping.
      if (!seen_group_sep) {
        if (digits_in_current_group < 1 || digits_in_current_group > 3) {
          return false;
        }
      } else {
        if (digits_in_current_group != 3) {
          return false;
        }
      }
      seen_group_sep = true;
      digits_in_current_group = 0;
      continue;
    }
    if ((c == 'e' || c == 'E') && seen_digit && !seen_exp) {
      // Transitioning out of integer part (no decimal seen): validate the
      // final integer group if any group separators were seen.
      if (!seen_point && seen_group_sep && digits_in_current_group != 3) {
        return false;
      }
      canonical.push_back('e');
      seen_exp = true;
      if (i + 1 < s.size() && (s[i + 1] == '+' || s[i + 1] == '-')) {
        canonical.push_back(s[i + 1]);
        ++i;
      }
      continue;
    }
    return false;
  }
  if (!seen_digit) {
    return false;
  }
  // End-of-input: if grouping was used and we never left the integer part,
  // the final group must also be exactly 3 digits.
  if (seen_group_sep && !seen_point && !seen_exp && digits_in_current_group != 3) {
    return false;
  }
  double parsed = 0.0;
  if (!parse_double_exact(canonical, &parsed) || std::isinf(parsed) || numeric_text_above_excel_max(canonical)) {
    return false;
  }
  // Apply percent scaling with division (not multiplication by 0.01) so the
  // result is bit-identical to Mac Excel for clean cases such as
  // `VALUE("23.5%")`. `23.5 * 0.01` is one ulp above the canonical `0.235`
  // that Excel returns; dividing by 100 directly lands on that bit pattern.
  for (int k = 0; k < percent_count; ++k) {
    parsed /= 100.0;
  }
  if (negative) {
    parsed = -parsed;
  }
  *out = parsed;
  return true;
}

std::string normalize_locale_numeric(std::string_view raw, bool* paren_negated) {
  *paren_negated = false;
  const bool fold_fullwidth = locale_facts(current_eval_profile()).fullwidth_numeric_text;
  std::string out;
  out.reserve(raw.size());
  std::size_t i = 0;
  while (i < raw.size()) {
    const unsigned char b0 = static_cast<unsigned char>(raw[i]);
    // 3-byte UTF-8 sequences cover the U+3000 / U+FF00 / U+FFE5 ranges we
    // care about; everything else passes through verbatim so multi-byte
    // tails (e.g. `¥` 0xC2 0xA5, the kanji `円`, etc.) reach `parse_numeric`
    // unchanged.
    if (fold_fullwidth && b0 >= 0xE0u && b0 < 0xF0u && i + 2 < raw.size()) {
      const unsigned char b1 = static_cast<unsigned char>(raw[i + 1]);
      const unsigned char b2 = static_cast<unsigned char>(raw[i + 2]);
      const std::uint32_t cp = (static_cast<std::uint32_t>(b0 & 0x0Fu) << 12) |
                               (static_cast<std::uint32_t>(b1 & 0x3Fu) << 6) | static_cast<std::uint32_t>(b2 & 0x3Fu);
      char ascii = '\0';
      if (cp >= 0xFF10u && cp <= 0xFF19u) {
        ascii = static_cast<char>('0' + (cp - 0xFF10u));
      } else if (cp >= 0xFF21u && cp <= 0xFF3Au) {
        ascii = static_cast<char>('A' + (cp - 0xFF21u));
      } else if (cp >= 0xFF41u && cp <= 0xFF5Au) {
        ascii = static_cast<char>('a' + (cp - 0xFF41u));
      } else if (cp == 0xFF0Eu) {
        ascii = '.';
      } else if (cp == 0xFF0Cu) {
        ascii = ',';
      } else if (cp == 0xFF05u) {
        ascii = '%';
      } else if (cp == 0xFF0Bu) {
        ascii = '+';
      } else if (cp == 0xFF0Du) {
        ascii = '-';
      } else if (cp == 0xFF08u) {
        ascii = '(';
      } else if (cp == 0xFF09u) {
        ascii = ')';
      } else if (cp == 0x3000u) {
        ascii = ' ';
      }
      if (ascii != '\0') {
        out.push_back(ascii);
        i += 3;
        continue;
      }
    }
    out.push_back(raw[i]);
    ++i;
  }
  // Accounting-style outer parens. Trim only ASCII whitespace because
  // `parse_numeric` does the same; a Japanese profile has already folded
  // full-width spaces to ASCII above.
  std::string_view trimmed = trim_ascii(out);
  if (trimmed.size() >= 3 && trimmed.front() == '(' && trimmed.back() == ')') {
    std::string_view inner = trimmed.substr(1, trimmed.size() - 2);
    inner = trim_ascii(inner);
    if (!inner.empty() && inner.front() != '+' && inner.front() != '-') {
      *paren_negated = true;
      return std::string(inner);
    }
  }
  return out;
}

bool numeric_text_above_excel_max(std::string_view text) {
  std::size_t i = 0;
  if (i < text.size() && (text[i] == '+' || text[i] == '-')) {
    ++i;
  }
  long long digits_before_point = 0;
  long long first_nonzero = -1;
  long long seen = 0;
  bool in_fraction = false;
  for (; i < text.size(); ++i) {
    const char c = text[i];
    if (c == '.' && !in_fraction) {
      in_fraction = true;
    } else if (c >= '0' && c <= '9') {
      if (c != '0' && first_nonzero < 0) {
        first_nonzero = seen;
      }
      ++seen;
      if (!in_fraction) {
        ++digits_before_point;
      }
    } else {
      break;
    }
  }
  if (first_nonzero < 0) {
    return false;
  }
  long long exponent = 0;
  if (i < text.size() && (text[i] == 'e' || text[i] == 'E')) {
    ++i;
    bool negative = false;
    if (i < text.size() && (text[i] == '+' || text[i] == '-')) {
      negative = text[i] == '-';
      ++i;
    }
    for (; i < text.size() && text[i] >= '0' && text[i] <= '9' && exponent < 100000; ++i) {
      exponent = exponent * 10 + (text[i] - '0');
    }
    if (negative) {
      exponent = -exponent;
    }
  }
  // The leading digit's decimal exponent decides it: fifteen nines at 307 is
  // the largest accepted value, so anything whose leading digit sits at 308
  // or beyond is over.
  return digits_before_point - 1 - first_nonzero + exponent >= 308;
}

bool parse_excel_number(std::string_view text, double* out) {
  bool paren_negated = false;
  const std::string normalized = normalize_locale_numeric(text, &paren_negated);
  double value = 0.0;
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  if (!parse_numeric(normalized, facts.decimal_separator, facts.group_separator, &value)) {
    return false;
  }
  *out = paren_negated ? -value : value;
  return true;
}

}  // namespace eval
}  // namespace formulon
