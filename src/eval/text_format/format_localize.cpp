// Localized TEXT format string mapping.

#include "eval/text_format/format_localize.h"

#include <cstddef>
#include <string>
#include <string_view>

#include "eval/text_format/number_format_scanner.h"
#include "excel_locale.h"
#include "utils/strings.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
namespace {

constexpr std::string_view kGeneralKeyword = "General";

bool is_ascii_letter(char c) noexcept {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
}

bool is_digit_placeholder(char c) noexcept {
  return c == '0' || c == '#' || c == '?';
}

// `word` at `s[i]`, not followed by an ASCII letter when it ends in one.
bool keyword_at(std::string_view s, std::size_t i, std::string_view word) noexcept {
  if (word.empty() || !strings::case_insensitive_starts_with(s.substr(i), word)) {
    return false;
  }
  const std::size_t end = i + word.size();
  return !is_ascii_letter(word.back()) || end >= s.size() || !is_ascii_letter(s[end]);
}

// Lower-cased letter of a body made of one repeated ASCII letter; '\0' otherwise.
char letter_run(std::string_view body) noexcept {
  if (body.empty() || !is_ascii_letter(body.front())) {
    return '\0';
  }
  const char unit = strings::ascii_to_lower(body.front());
  for (const char c : body) {
    if (strings::ascii_to_lower(c) != unit) {
      return '\0';
    }
  }
  return unit;
}

// Format letter an alias spelling at `s[i]` stands for, with its byte length; '\0' when none starts there.
char alias_at(std::string_view s, std::size_t i, const LocaleFacts& facts, std::size_t* len) noexcept {
  for (const FormatLetterAlias& alias : facts.format_letter_aliases) {
    if (!alias.spelling.empty() && s.compare(i, alias.spelling.size(), alias.spelling) == 0) {
      *len = alias.spelling.size();
      return alias.letter;
    }
  }
  return '\0';
}

// Lower-cased letter of a body made of one repeated alias letter; '\0' otherwise.
char alias_run(std::string_view body, const LocaleFacts& facts) noexcept {
  char unit = '\0';
  std::size_t i = 0;
  while (i < body.size()) {
    std::size_t len = 0;
    const char letter = strings::ascii_to_lower(alias_at(body, i, facts, &len));
    if (letter == '\0' || (unit != '\0' && letter != unit)) {
      return '\0';
    }
    unit = letter;
    i += len;
  }
  return unit;
}

// Rewrites a bracket body's colour name into the stored spelling. Returns
// false when the body is a stored colour name the locale does not accept.
bool localize_bracket(std::string_view body, const LocaleFacts& facts, const LocaleFacts& english, std::string& out) {
  const std::string_view index_prefix = facts.color_index_prefix;
  // An elapsed-time body (`[h]`, `[mm]`) spells its unit with the locale's
  // letter; the invariant letter of a unit the locale spells differently is
  // rejected (text_format.text_elapsed_hours).
  const bool aliased = !facts.format_letter_aliases[0].spelling.empty();
  const char ascii_unit = letter_run(body);
  if (aliased && (ascii_unit == 'h' || ascii_unit == 'm' || ascii_unit == 's')) {
    out.append(body);
    return false;
  }
  if (const char unit = aliased ? alias_run(body, facts) : ascii_unit; unit != '\0') {
    const FormatLetters& letters = facts.format_letters;
    const char stored = unit == strings::ascii_to_lower(letters.hour)     ? 'h'
                        : unit == strings::ascii_to_lower(letters.minute) ? 'm'
                        : unit == strings::ascii_to_lower(letters.second) ? 's'
                                                                          : '\0';
    if (stored != '\0') {
      out.append(body.size(), stored);
      return true;
    }
    if (unit == 'h' || unit == 'm' || unit == 's') {
      out.append(body);
      return false;
    }
  }
  for (std::size_t k = 0; k < facts.color_names.size(); ++k) {
    const std::string_view name = facts.color_names[k];
    if (!name.empty() && strings::case_insensitive_starts_with(body, name)) {
      out.append(english.color_names[k]);
      out.append(body.substr(name.size()));
      return true;
    }
  }
  if (!index_prefix.empty() && strings::case_insensitive_starts_with(body, index_prefix)) {
    out.append(kStoredColorIndexPrefix);
    out.append(body.substr(index_prefix.size()));
    return true;
  }
  out.append(body);
  for (const std::string_view name : english.color_names) {
    if (strings::case_insensitive_starts_with(body, name)) {
      return false;
    }
  }
  return !strings::case_insensitive_starts_with(body, kStoredColorIndexPrefix);
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
  const char decimal = facts.decimal_separator;
  const char group = facts.group_separator;
  const bool map_separators = decimal != '.' || group != ',';
  const bool aliased = !facts.format_letter_aliases[0].spelling.empty();
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
      if (!localize_bracket(s.substr(i + 1, close - i - 1), facts, english, out)) {
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
      if (!facts.english_general) {
        result.valid = false;
      }
      out.append(s.substr(i, kGeneralKeyword.size()));
      i += kGeneralKeyword.size();
      continue;
    }
    if (aliased) {
      // AM/PM and A/P keep their Latin letters; the other Latin date letters are text.
      if (keyword_at(s, i, "AM/PM") || keyword_at(s, i, "A/P")) {
        const std::size_t len = keyword_at(s, i, "AM/PM") ? 5U : 3U;
        out.append(s.substr(i, len));
        i += len;
        continue;
      }
      std::size_t len = 0;
      if (const char letter = alias_at(s, i, facts, &len); letter != '\0') {
        out.push_back(letter);
        i += len;
        continue;
      }
      const char lc = strings::ascii_to_lower(c);
      if (lc == 'y' || lc == 'm' || lc == 'd' || lc == 'h' || lc == 's') {
        out.push_back('\\');
        out.push_back(c);
        ++i;
        continue;
      }
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
      if (c == '.' && facts.format_rejects_dot) {
        result.valid = false;
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
