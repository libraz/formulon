// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the shared locale-aware numeric parsers declared in
// `number_parse.h`. Extracted verbatim from the VALUE / NUMBERVALUE
// builtins so the same normalisation drives implicit text->number coercion
// (arithmetic operators and the criteria engine).

#include "eval/number_parse.h"

#include <locale.h>  // newlocale / uselocale / freelocale (POSIX 2008)

#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <cstring>

namespace formulon {
namespace eval {
namespace {

// Cached C locale used to make numeric parsing independent of the host
// process's LC_NUMERIC. Created once; never freed (process-lifetime). A
// `(locale_t)0` result (allocation failure) makes callers fall back to the
// plain `std::strtod` path.
locale_t c_numeric_locale() noexcept {
  static const locale_t loc = newlocale(LC_NUMERIC_MASK, "C", static_cast<locale_t>(0));
  return loc;
}

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

// Strips a leading currency prefix. The accepted set is the host profile's
// currency symbols; for the ja-JP profile this is {$, ¥, ￥, €} (ASCII '$',
// UTF-8 '¥' 0xC2 0xA5, full-width '￥' 0xEF 0xBF 0xA5, '€' 0xE2 0x82 0xAC).
// `£` / `¢` / `₩` / the kanji `円` are deliberately excluded (oracle-verified
// against Mac Excel 365 ja-JP; a future en-GB profile would add '£').
// Returns the input unchanged when no accepted symbol is present.
std::string_view strip_currency(std::string_view s) noexcept {
  if (s.empty()) {
    return s;
  }
  if (s.front() == '$') {
    return s.substr(1);
  }
  if (s.size() >= 2 && static_cast<unsigned char>(s[0]) == 0xC2u && static_cast<unsigned char>(s[1]) == 0xA5u) {
    return s.substr(2);
  }
  if (s.size() >= 3 && static_cast<unsigned char>(s[0]) == 0xEFu && static_cast<unsigned char>(s[1]) == 0xBFu &&
      static_cast<unsigned char>(s[2]) == 0xA5u) {
    return s.substr(3);
  }
  if (s.size() >= 3 && static_cast<unsigned char>(s[0]) == 0xE2u && static_cast<unsigned char>(s[1]) == 0x82u &&
      static_cast<unsigned char>(s[2]) == 0xACu) {
    return s.substr(3);
  }
  return s;
}

// Strips a trailing Euro suffix (`23€` -> `23`). Mac Excel 365 accepts Euro
// both as prefix and suffix; the dollar sign and the yen kanji (`円`) are
// NOT accepted as suffix, so this is deliberately Euro-only.
std::string_view strip_trailing_euro(std::string_view s) noexcept {
  if (s.size() >= 3 && static_cast<unsigned char>(s[s.size() - 3]) == 0xE2u &&
      static_cast<unsigned char>(s[s.size() - 2]) == 0x82u && static_cast<unsigned char>(s[s.size() - 1]) == 0xACu) {
    return s.substr(0, s.size() - 3);
  }
  return s;
}

}  // namespace

double parse_double_c_locale(const char* str, char** endptr) noexcept {
  const locale_t loc = c_numeric_locale();
  if (loc == static_cast<locale_t>(0)) {
    return std::strtod(str, endptr);
  }
  // uselocale swaps only the calling thread's locale, so this is safe under
  // the scheduler's worker threads and does not disturb other threads.
  const locale_t previous = uselocale(loc);
  const double value = std::strtod(str, endptr);
  uselocale(previous);
  return value;
}

bool parse_numeric(std::string_view s, char decimal_sep, char group_sep, double* out) noexcept {
  // Trim ASCII whitespace first.
  s = trim_ascii(s);
  if (s.empty()) {
    return false;
  }
  // Sign and currency may appear in either order at the front, and Mac Excel
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
  // stray `€` that the numeric scan below rejects. Mac Excel 365 accepts `€`
  // as a suffix (e.g. `"23€"`); `$` and `円` are not accepted as suffixes.
  if (!leading_currency) {
    s = strip_trailing_euro(s);
    if (s.empty()) {
      return false;
    }
  }
  // Trailing percent signs. Each `%` multiplies the parsed value by 0.01,
  // so `"50%%"` yields `0.005` (matches Mac Excel ja-JP NUMBERVALUE).
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
  // Parse via std::strtod over a NUL-terminated buffer.
  char stack_buf[64];
  char* heap_buf = nullptr;
  const std::size_t n = canonical.size();
  char* buf = stack_buf;
  if (n + 1 > sizeof(stack_buf)) {
    heap_buf = static_cast<char*>(std::malloc(n + 1));
    if (heap_buf == nullptr) {
      return false;
    }
    buf = heap_buf;
  }
  std::memcpy(buf, canonical.data(), n);
  buf[n] = '\0';
  char* end_ptr = nullptr;
  double parsed = parse_double_c_locale(buf, &end_ptr);
  const bool ok = end_ptr == buf + n;
  if (heap_buf != nullptr) {
    std::free(heap_buf);
  }
  if (!ok) {
    return false;
  }
  if (std::isnan(parsed) || std::isinf(parsed)) {
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
  std::string out;
  out.reserve(raw.size());
  std::size_t i = 0;
  while (i < raw.size()) {
    const unsigned char b0 = static_cast<unsigned char>(raw[i]);
    // 3-byte UTF-8 sequences cover the U+3000 / U+FF00 / U+FFE5 ranges we
    // care about; everything else passes through verbatim so multi-byte
    // tails (e.g. `¥` 0xC2 0xA5, the kanji `円`, etc.) reach `parse_numeric`
    // unchanged.
    if (b0 >= 0xE0u && b0 < 0xF0u && i + 2 < raw.size()) {
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
  // `parse_numeric` does the same; full-width spaces have already been
  // folded to ASCII above.
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

bool parse_excel_number(std::string_view text, double* out) {
  bool paren_negated = false;
  const std::string normalized = normalize_locale_numeric(text, &paren_negated);
  double value = 0.0;
  if (!parse_numeric(normalized, '.', ',', &value)) {
    return false;
  }
  *out = paren_negated ? -value : value;
  return true;
}

}  // namespace eval
}  // namespace formulon
