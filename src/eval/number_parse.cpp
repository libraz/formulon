//
// Implementation of the shared locale-aware numeric parsers declared in
// `number_parse.h`. Extracted verbatim from the VALUE / NUMBERVALUE
// builtins so the same normalisation drives implicit text->number coercion
// (arithmetic operators and the criteria engine).

#include "eval/number_parse.h"

#include <cstddef>
#include <cstdint>
#include <string>

#include "eval/eval_profile_scope.h"
#include "excel_locale.h"
#include "utils/double_parse.h"
#include "utils/number_text.h"
#include "utils/utf8_length.h"

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

// Decodes the UTF-8 sequence at `s[i]`; a malformed byte decodes as itself.
std::uint32_t decode_utf8(std::string_view s, std::size_t i, std::size_t* len) noexcept {
  const unsigned char b0 = static_cast<unsigned char>(s[i]);
  const std::size_t n = b0 < 0xC0u ? 1u : b0 < 0xE0u ? 2u : b0 < 0xF0u ? 3u : 4u;
  if (n == 1u || i + n > s.size()) {
    *len = 1;
    return b0;
  }
  std::uint32_t cp = b0 & (0xFFu >> (n + 1u));
  for (std::size_t k = 1; k < n; ++k) {
    cp = (cp << 6) | (static_cast<unsigned char>(s[i + k]) & 0x3Fu);
  }
  *len = n;
  return cp;
}

// Decimal digits (Unicode Nd) of the non-ASCII scripts Excel reads; 0 when
// `cp` is not one. Tamil has no zero, only 1..9.
char fold_decimal_digit(std::uint32_t cp) noexcept {
  constexpr std::uint32_t kZeros[] = {0x0660, 0x06F0, 0x0966, 0x09E6, 0x0A66, 0x0AE6, 0x0B66, 0x0C66,
                                      0x0CE6, 0x0D66, 0x0E50, 0x0ED0, 0x0F20, 0x1040, 0x17E0, 0x1810};
  for (const std::uint32_t zero : kZeros) {
    if (cp >= zero && cp <= zero + 9u) {
      return static_cast<char>('0' + (cp - zero));
    }
  }
  if (cp >= 0x0BE7u && cp <= 0x0BEFu) {
    return static_cast<char>('1' + (cp - 0x0BE7u));
  }
  return '\0';
}

// The decimal digits plus the Ethiopic 1..9 (no zero), which numbers read
// but date and time text does not.
char fold_script_digit(std::uint32_t cp) noexcept {
  const char digit = fold_decimal_digit(cp);
  if (digit != '\0') {
    return digit;
  }
  if (cp >= 0x1369u && cp <= 0x1371u) {
    return static_cast<char>('1' + (cp - 0x1369u));
  }
  return '\0';
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

// Strips a trailing currency suffix the locale accepts (`23€`, `1 ₽`), with
// or without a blank before it.
std::string_view strip_trailing_currency(std::string_view s) noexcept {
  for (const std::string_view symbol : locale_facts(current_eval_profile()).accepted_currency_suffix) {
    if (!symbol.empty() && s.size() >= symbol.size() && s.substr(s.size() - symbol.size()) == symbol) {
      s.remove_suffix(symbol.size());
      return trim_ascii(s);
    }
  }
  return s;
}

// Assembles a canonical C-locale numeric string (digits, one optional `.`,
// optional exponent `e[+/-]digits`), rejecting any other byte. Grouping: the
// first group is 1+ digits with at least one nonzero (all-zero is rejected),
// every later one 3+.
bool canonical_number(std::string_view s, char decimal_sep, char group_sep, std::string* out) {
  std::string& canonical = *out;
  canonical.reserve(s.size());
  bool seen_digit = false;
  bool seen_point = false;
  bool seen_exp = false;
  bool seen_group_sep = false;
  int digits_in_current_group = 0;
  bool first_group_has_nonzero = false;
  const auto group_ok = [&]() {
    return seen_group_sep ? digits_in_current_group >= 3 : first_group_has_nonzero && digits_in_current_group >= 1;
  };
  for (std::size_t i = 0; i < s.size(); ++i) {
    const char c = s[i];
    if (c >= '0' && c <= '9') {
      canonical.push_back(c);
      seen_digit = true;
      if (!seen_point && !seen_exp) {
        first_group_has_nonzero = first_group_has_nonzero || (c != '0' && !seen_group_sep);
        ++digits_in_current_group;
      }
      continue;
    }
    const bool in_integer = !seen_point && !seen_exp;
    if (c == decimal_sep && in_integer) {
      if (seen_group_sep && !group_ok()) {
        return false;
      }
      canonical.push_back('.');
      seen_point = true;
      continue;
    }
    if (group_sep != '\0' && c == group_sep && in_integer) {
      if (!group_ok()) {
        return false;
      }
      seen_group_sep = true;
      digits_in_current_group = 0;
      continue;
    }
    if ((c == 'e' || c == 'E') && seen_digit && !seen_exp) {
      if (in_integer && seen_group_sep && !group_ok()) {
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
  return seen_digit && !(seen_group_sep && !seen_point && !seen_exp && !group_ok());
}

// The `a/b` of a mixed fraction: plain digits, each part at most five digits
// (leading zeros count) and at most 32767, and a non-zero denominator.
bool parse_fraction(std::string_view s, int* numerator, int* denominator) noexcept {
  const std::size_t slash = s.find('/');
  const auto part = [](std::string_view digits, int* value) {
    if (digits.empty() || digits.size() > 5) {
      return false;
    }
    int v = 0;
    for (const char c : digits) {
      if (c < '0' || c > '9') {
        return false;
      }
      v = v * 10 + (c - '0');
      if (v > 32767) {
        return false;
      }
    }
    *value = v;
    return true;
  };
  return slash != std::string_view::npos && part(s.substr(0, slash), numerator) &&
         part(s.substr(slash + 1), denominator) && *denominator > 0;
}

}  // namespace

bool parse_numeric(std::string_view s, char decimal_sep, char group_sep, int percent_count, double* out) noexcept {
  s = trim_ascii(s);
  const auto skip_blanks = [&s]() {
    while (!s.empty() && s.front() == ' ') {
      s.remove_prefix(1);
    }
  };
  const auto trim_back_blanks = [&s]() {
    while (!s.empty() && s.back() == ' ') {
      s.remove_suffix(1);
    }
  };
  // Front: `(` or a sign, then a currency symbol, then (when neither came
  // first) a sign or `(` after the symbol: "(¥1)", "-¥1", "¥-1", "¥(1)".
  // Only blanks may sit between a sign and the digits; a tab there is no number.
  bool negative = false;
  bool paren = false;
  bool signed_front = false;
  const auto take_sign = [&]() {
    if (!s.empty() && (s.front() == '+' || s.front() == '-')) {
      negative = s.front() == '-';
      s.remove_prefix(1);
      skip_blanks();
      return true;
    }
    return false;
  };
  if (!s.empty() && s.front() == '(') {
    paren = true;
    s.remove_prefix(1);
    skip_blanks();
  } else {
    signed_front = take_sign();
  }
  bool currency = false;
  {
    const std::string_view after = strip_currency(s);
    if (after.size() != s.size()) {
      s = after;
      currency = true;
      while (!s.empty() && is_ascii_ws(static_cast<unsigned char>(s.front()))) {
        s.remove_prefix(1);
      }
      if (!paren && !signed_front && !take_sign() && !s.empty() && s.front() == '(') {
        paren = true;
        s.remove_prefix(1);
        skip_blanks();
      }
    }
  }
  // Back: a currency suffix (one currency side only) and the closing `)`, in
  // either order: "(5)€", "(5€)".
  const auto take_suffix = [&]() {
    if (currency) {
      return;
    }
    const std::string_view before = strip_trailing_currency(s);
    if (before.size() != s.size()) {
      s = before;
      currency = true;
      trim_back_blanks();
    }
  };
  trim_back_blanks();
  take_suffix();
  if (paren) {
    if (s.empty() || s.back() != ')') {
      return false;
    }
    s.remove_suffix(1);
    trim_back_blanks();
    take_suffix();
  }
  if (s.empty() || (currency && percent_count > 0)) {
    return false;
  }
  // A mixed fraction `N a/b`: a whole number, one blank, then a fraction.
  // Where the blank groups digits, the numerator takes at most two digits.
  const std::size_t blank = s.rfind(' ');
  if (blank != std::string_view::npos && s.substr(blank).find('/') != std::string_view::npos) {
    int numerator = 0;
    int denominator = 0;
    std::string whole_text;
    double whole = 0.0;
    const std::string_view fraction = s.substr(blank + 1);
    if ((group_sep == ' ' && fraction.find('/') > 2) || !parse_fraction(fraction, &numerator, &denominator) ||
        !canonical_number(s.substr(0, blank), decimal_sep, group_sep, &whole_text) ||
        whole_text.find_first_of(".e") != std::string::npos || !excel_text_number(whole_text, false, 0, &whole)) {
      return false;
    }
    double value = whole + static_cast<double>(numerator) / static_cast<double>(denominator);
    for (int k = 0; k < percent_count; ++k) {
      value /= 100.0;
    }
    *out = negative != paren && value != 0.0 ? -value : value;
    return true;
  }
  std::string canonical;
  double parsed = 0.0;
  return canonical_number(s, decimal_sep, group_sep, &canonical) && parse_double_exact(canonical, &parsed) &&
         excel_text_number(canonical, negative != paren, percent_count, out);
}

std::string normalize_locale_numeric(std::string_view raw, int* percent_count) {
  const bool fold_fullwidth = locale_facts(current_eval_profile()).fullwidth_numeric_text;
  std::string out;
  out.reserve(raw.size());
  for (std::size_t i = 0; i < raw.size();) {
    std::size_t len = 1;
    const std::uint32_t cp = decode_utf8(raw, i, &len);
    char ascii = cp >= 0x80u ? fold_script_digit(cp) : '\0';
    if (ascii == '\0' && fold_fullwidth) {
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
      } else if (cp == 0xFF0Fu) {
        ascii = '/';
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
      } else if (cp == 0xFF04u) {
        ascii = '$';
      } else if (cp == 0x3000u) {
        ascii = ' ';
      }
    }
    if (ascii != '\0') {
      out.push_back(ascii);
    } else {
      out.append(raw.substr(i, len));
    }
    i += len;
  }
  // Excel takes one `%` anywhere outside the digit run (either side of the
  // sign, inside or outside the parentheses); a second one, or one with number
  // characters on both sides, is not a number.
  *percent_count = 0;
  const auto body = [](char c) { return (c >= '0' && c <= '9') || c == '.' || c == ',' || c == 'e' || c == 'E'; };
  for (std::size_t k = 0; k < out.size(); ++k) {
    if (out[k] != '%') {
      continue;
    }
    if (*percent_count != 0 || (k > 0 && body(out[k - 1]) && k + 1 < out.size() && body(out[k + 1]))) {
      *percent_count = -1;
      return out;
    }
    *percent_count = 1;
    out.erase(k, 1);
    --k;
  }
  return out;
}

std::string fold_script_digits(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  for (std::size_t i = 0; i < text.size();) {
    std::size_t len = 1;
    const char digit = fold_decimal_digit(decode_utf8(text, i, &len));
    if (digit != '\0') {
      out.push_back(digit);
    } else {
      out.append(text.substr(i, len));
    }
    i += len;
  }
  return out;
}

std::uint32_t first_utf16_unit(std::string_view text) noexcept {
  if (text.empty()) {
    return 0;
  }
  std::size_t len = 1;
  const std::uint32_t cp = decode_utf8(text, 0, &len);
  return cp >= 0x10000u ? 0xD800u + ((cp - 0x10000u) >> 10) : cp;
}

bool numbervalue_prepass(std::string_view text, std::uint32_t decimal_unit, std::uint32_t group_unit, std::string* out,
                         int* percent_count) {
  if (numeric_text_too_long(text) || decimal_unit == group_unit) {
    return false;
  }
  const char locale_decimal = locale_facts(current_eval_profile()).decimal_separator;
  out->clear();
  bool seen_decimal = false;
  for (std::size_t i = 0; i < text.size();) {
    std::size_t len = 1;
    const std::uint32_t cp = decode_utf8(text, i, &len);
    if (cp == ' ' || cp == '\t' || cp == '\n' || cp == '\r') {
      // Dropped.
    } else if (cp == decimal_unit) {
      out->push_back(locale_decimal);
      seen_decimal = true;
    } else if (cp == group_unit) {
      if (seen_decimal) {
        return false;
      }
    } else {
      out->append(text.substr(i, len));
    }
    i += len;
  }
  *percent_count = 0;
  while (!out->empty() && out->back() == '%') {
    out->pop_back();
    ++*percent_count;
  }
  return true;
}

bool numeric_text_too_long(std::string_view raw) noexcept {
  return utf16_units_in(raw) > kMaxNumericTextLength;
}

bool excel_text_number(std::string_view magnitude, bool negative, int percent_count, double* out) {
  // At most 15 significant digits are kept (truncated); `scale` is the decimal
  // exponent of the kept mantissa.
  constexpr int kKeptDigits = 15;
  constexpr long long kExponentCap = 1000000000LL;
  std::uint64_t mantissa = 0;
  int kept = 0;
  long long scale = 0;
  std::size_t i = 0;
  for (; i < magnitude.size() && magnitude[i] >= '0' && magnitude[i] <= '9'; ++i) {
    const int d = magnitude[i] - '0';
    if (mantissa == 0 && d == 0) {
      continue;
    }
    if (kept < kKeptDigits) {
      mantissa = mantissa * 10U + static_cast<std::uint64_t>(d);
      ++kept;
    } else {
      ++scale;
    }
  }
  if (i < magnitude.size() && magnitude[i] == '.') {
    for (++i; i < magnitude.size() && magnitude[i] >= '0' && magnitude[i] <= '9'; ++i) {
      const int d = magnitude[i] - '0';
      if (mantissa == 0 && d == 0) {
        --scale;
      } else if (kept < kKeptDigits) {
        mantissa = mantissa * 10U + static_cast<std::uint64_t>(d);
        ++kept;
        --scale;
      }
    }
  }
  if (i < magnitude.size() && (magnitude[i] == 'e' || magnitude[i] == 'E')) {
    ++i;
    bool exponent_negative = false;
    if (i < magnitude.size() && (magnitude[i] == '+' || magnitude[i] == '-')) {
      exponent_negative = magnitude[i] == '-';
      ++i;
    }
    long long written = 0;
    for (; i < magnitude.size() && magnitude[i] >= '0' && magnitude[i] <= '9'; ++i) {
      if (written < kExponentCap) {
        written = written * 10 + (magnitude[i] - '0');
      }
    }
    scale += exponent_negative ? -written : written;
  }
  scale -= 2LL * percent_count;
  // The check sits on the leading digit's decimal exponent; zero has none and
  // is judged with no digits.
  if (!excel_leading_exponent_in_range(scale + kept - 1)) {
    return false;
  }
  const double value = excel_scale_decimal(mantissa, scale);
  *out = negative && value != 0.0 ? -value : value;
  return true;
}

bool parse_excel_number(std::string_view text, double* out) {
  if (numeric_text_too_long(text)) {
    return false;
  }
  int percent_count = 0;
  const std::string normalized = normalize_locale_numeric(text, &percent_count);
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  return percent_count >= 0 &&
         parse_numeric(normalized, facts.decimal_separator, facts.group_separator, percent_count, out);
}

}  // namespace eval
}  // namespace formulon
