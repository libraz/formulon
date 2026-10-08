// Localized TEXT format string mapping.

#include "eval/text_format/format_localize.h"

#include <cstddef>
#include <string>
#include <string_view>

#include "eval/text_format/number_format_scanner.h"
#include "excel_locale.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
namespace {

constexpr std::string_view kGeneralKeyword = "General";

char ascii_lower(char c) noexcept {
  return c >= 'A' && c <= 'Z' ? static_cast<char>(c + ('a' - 'A')) : c;
}

bool is_ascii_letter(char c) noexcept {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
}

bool is_digit_placeholder(char c) noexcept {
  return c == '0' || c == '#' || c == '?';
}

// ASCII case-insensitive prefix test; non-ASCII bytes compare exactly.
bool starts_with_ci(std::string_view s, std::string_view prefix) noexcept {
  if (s.size() < prefix.size()) {
    return false;
  }
  for (std::size_t i = 0; i < prefix.size(); ++i) {
    if (ascii_lower(s[i]) != ascii_lower(prefix[i])) {
      return false;
    }
  }
  return true;
}

// `word` at `s[i]`, not followed by an ASCII letter when it ends in one.
bool keyword_at(std::string_view s, std::size_t i, std::string_view word) noexcept {
  if (word.empty() || !starts_with_ci(s.substr(i), word)) {
    return false;
  }
  const std::size_t end = i + word.size();
  return !is_ascii_letter(word.back()) || end >= s.size() || !is_ascii_letter(s[end]);
}

// Locale spelling of the indexed colour form (`[色12]`); empty where none
// is measured.
std::string_view color_index_prefix(ExcelLocale locale) noexcept {
  switch (locale) {
    case ExcelLocale::kJaJP:
      return "色";  // text_format.text_color_indexed_discarded
    case ExcelLocale::kEnUS:
      return kStoredColorIndexPrefix;
    default:
      return {};  // unmeasured
  }
}

// Rewrites a bracket body's colour name into the stored spelling. Returns
// false when the body is a stored colour name the locale does not accept.
bool localize_bracket(std::string_view body, const LocaleFacts& facts, const LocaleFacts& english,
                      std::string_view index_prefix, std::string& out) {
  for (std::size_t k = 0; k < facts.color_names.size(); ++k) {
    const std::string_view name = facts.color_names[k];
    if (!name.empty() && starts_with_ci(body, name)) {
      out.append(english.color_names[k]);
      out.append(body.substr(name.size()));
      return true;
    }
  }
  if (!index_prefix.empty() && starts_with_ci(body, index_prefix)) {
    out.append(kStoredColorIndexPrefix);
    out.append(body.substr(index_prefix.size()));
    return true;
  }
  out.append(body);
  for (const std::string_view name : english.color_names) {
    if (starts_with_ci(body, name)) {
      return false;
    }
  }
  return !starts_with_ci(body, kStoredColorIndexPrefix);
}

}  // namespace

LocalizedFormat localize_format(std::string_view fmt, FormatDialect dialect, ExcelProfile profile) {
  const LocaleFacts& facts = locale_facts(profile);
  LocalizedFormat result;
  const std::string folded = facts.fullwidth_syntax_fold ? normalize_ja_jp_format_syntax(fmt) : std::string();
  const std::string_view s = facts.fullwidth_syntax_fold ? std::string_view(folded) : fmt;
  if (dialect == FormatDialect::kStored) {
    result.text.assign(s);
    return result;
  }

  const LocaleFacts& english = locale_facts(mac_365_en_us_profile());
  const std::string_view index_prefix = color_index_prefix(profile.locale);
  const char decimal = facts.decimal_separator;
  const char group = facts.group_separator;
  const bool map_separators = decimal != '.' || group != ',';
  const bool english_general = facts.general_alias.empty();  // locale_tokens.text_general_english
  std::string& out = result.text;
  out.reserve(s.size() + 8U);

  std::size_t i = 0;
  while (i < s.size()) {
    const char c = s[i];
    if (c == '"') {
      const std::size_t close = s.find('"', i + 1);
      const std::size_t end = close == std::string_view::npos ? s.size() : close + 1;
      out.append(s.substr(i, end - i));
      i = end;
      continue;
    }
    if (c == '\\' || (c == '!' && facts.bang_escape) || c == '_' || c == '*') {
      const std::size_t end = i + 1 < s.size() ? i + 1 + utf8_scalar_width(s, i + 1) : s.size();
      out.append(s.substr(i, end - i));
      i = end;
      continue;
    }
    if (c == '[') {
      const std::size_t close = s.find(']', i + 1);
      if (close == std::string_view::npos) {
        out.append(s.substr(i));
        break;
      }
      out.push_back('[');
      if (!localize_bracket(s.substr(i + 1, close - i - 1), facts, english, index_prefix, out)) {
        result.valid = false;
      }
      out.push_back(']');
      i = close + 1;
      continue;
    }
    if (!facts.general_alias.empty() && keyword_at(s, i, facts.general_alias)) {
      out.append(kGeneralKeyword);
      i += facts.general_alias.size();
      continue;
    }
    if (keyword_at(s, i, kGeneralKeyword)) {
      if (!english_general) {
        result.valid = false;
      }
      out.append(s.substr(i, kGeneralKeyword.size()));
      i += kGeneralKeyword.size();
      continue;
    }
    if (map_separators) {
      if (c == decimal) {
        out.push_back('.');
        ++i;
        continue;
      }
      if (c == group) {
        // A blank groups only between digit placeholders (`# ##0`); elsewhere it is text.
        const bool groups = group != ' ' || (i > 0 && i + 1 < s.size() && is_digit_placeholder(s[i - 1]) &&
                                             is_digit_placeholder(s[i + 1]));
        out.push_back(groups ? ',' : c);
        ++i;
        continue;
      }
      if (c == '.' || c == ',') {
        // The invariant separator the locale does not use is plain text.
        out.push_back('\\');
        out.push_back(c);
        ++i;
        continue;
      }
    }
    out.push_back(c);
    ++i;
  }
  return result;
}

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
