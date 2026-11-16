//
// Spelling of names inside pivot formulas: a field name as the formula text
// writes it, and the synthetic identifier a parsed formula carries for each
// `FormulaRef`.

#ifndef FORMULON_PIVOT_PIVOT_FORMULA_NAMES_H_
#define FORMULON_PIVOT_PIVOT_FORMULA_NAMES_H_

#include <cstddef>
#include <optional>
#include <string>
#include <string_view>

#include "utils/strings.h"

namespace formulon::pivot {

/// Prefix of the synthetic identifier standing for a `FormulaRef`; reserved in formula text.
inline constexpr std::string_view kFormulaRefPrefix = "_fmpv";

/// Bytes that extend an identifier run; a superset of the tokenizer's name
/// bytes so a run never stops inside a cell reference or number.
inline bool is_formula_name_byte(unsigned char c) noexcept {
  const char ch = static_cast<char>(c);
  return strings::is_ascii_alpha(ch) || strings::is_ascii_digit(ch) || c == '_' || c == '.' || c == '\\' || c == '?' ||
         c == '$' || c >= 0x80;
}

/// Appends the synthetic identifier for `refs[index]`.
inline void append_formula_ref_name(std::string& out, std::size_t index) {
  out.append(kFormulaRefPrefix);
  char digits[20];
  std::size_t n = 0;
  do {
    digits[n++] = static_cast<char>('0' + index % 10);
    index /= 10;
  } while (index != 0);
  while (n > 0) {
    out.push_back(digits[--n]);
  }
}

/// The `refs` index a synthetic identifier stands for, or nullopt when it is not one below `ref_count`.
inline std::optional<std::size_t> formula_ref_slot(std::string_view name, std::size_t ref_count) {
  if (!strings::case_insensitive_starts_with(name, kFormulaRefPrefix) || name.size() == kFormulaRefPrefix.size()) {
    return std::nullopt;
  }
  std::size_t index = 0;
  for (const char ch : name.substr(kFormulaRefPrefix.size())) {
    if (ch < '0' || ch > '9' || index > ref_count) {
      return std::nullopt;
    }
    index = index * 10 + static_cast<std::size_t>(ch - '0');
  }
  return index < ref_count ? std::optional<std::size_t>(index) : std::nullopt;
}

/// Appends a field name as the formula text writes it: bare when it reads back
/// as one identifier run, else `'...'` with quotes doubled.
inline void append_formula_name(std::string& out, std::string_view name) {
  bool quote = name.empty() || strings::case_insensitive_starts_with(name, kFormulaRefPrefix);
  for (std::size_t i = 0; !quote && i < name.size(); ++i) {
    quote = !is_formula_name_byte(static_cast<unsigned char>(name[i]));
  }
  if (!quote) {
    out.append(name);
    return;
  }
  out.push_back('\'');
  for (const char ch : name) {
    out.append(ch == '\'' ? 2U : 1U, ch);
  }
  out.push_back('\'');
}

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_FORMULA_NAMES_H_
