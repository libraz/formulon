#include "eval/collated_wildcard.h"

#include <cstdint>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "eval/collation.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"

namespace formulon {
namespace eval {
namespace {

// One text code point under the `=` equivalence: `id` names its collation
// class, `cp` / `byte` locate the source character.
struct TextToken {
  std::uint32_t id;
  std::uint32_t cp;
  std::size_t byte;
};

enum class PatternKind : std::uint8_t { kLiteral, kAny, kStar };

struct PatternItem {
  PatternKind kind;
  std::uint32_t id;
};

// Interns each code point's comparison collation key, so two characters are
// one class exactly when `=` calls them equal. Expansions and contractions
// never span characters, matching SEARCH (measured on Mac Excel 365).
class CharClasses {
 public:
  explicit CharClasses(ExcelProfile profile) : profile_(profile) {}

  std::uint32_t id_of(std::uint32_t cp) {
    std::string utf8 = encode_utf8_codepoint(cp);
    const auto cached = by_char_.find(utf8);
    if (cached != by_char_.end()) {
      return cached->second;
    }
    const std::string key = collation_key(utf8, profile_, CollationStrength::kCompare);
    const auto inserted = by_key_.emplace(key, static_cast<std::uint32_t>(by_key_.size()));
    by_char_.emplace(std::move(utf8), inserted.first->second);
    return inserted.first->second;
  }

 private:
  ExcelProfile profile_;
  std::unordered_map<std::string, std::uint32_t> by_char_;
  std::unordered_map<std::string, std::uint32_t> by_key_;
};

// `?` / `*` are wildcards and `~` makes the next character literal; a
// trailing `~` is dropped.
std::vector<PatternItem> parse_pattern(std::string_view pattern, CharClasses& classes) {
  std::vector<PatternItem> out;
  bool escaped = false;
  for (std::size_t i = 0; i < pattern.size();) {
    std::size_t step = 0;
    const std::uint32_t cp = decode_utf8_step(pattern, i, &step);
    i += step;
    if (!escaped && cp == '~') {
      escaped = true;
      continue;
    }
    if (!escaped && cp == '*') {
      out.push_back({PatternKind::kStar, 0U});
    } else if (!escaped && cp == '?') {
      out.push_back({PatternKind::kAny, 0U});
    } else {
      out.push_back({PatternKind::kLiteral, classes.id_of(cp)});
    }
    escaped = false;
  }
  return out;
}

std::vector<TextToken> tokenize(std::string_view text, std::size_t start_byte, CharClasses& classes) {
  std::vector<TextToken> out;
  for (std::size_t i = start_byte; i < text.size();) {
    std::size_t step = 0;
    const std::uint32_t cp = decode_utf8_step(text, i, &step);
    out.push_back({classes.id_of(cp), cp, i});
    i += step;
  }
  return out;
}

// Whether `pattern` matches `text` from `start`: a prefix of it, or under
// `whole` all of it. Under `sbcs_question`, `?` takes only a single-byte
// character.
bool match_at(const std::vector<PatternItem>& pattern, const std::vector<TextToken>& text, std::size_t start,
              bool sbcs_question, bool whole) {
  std::size_t pi = 0;
  std::size_t ti = start;
  std::size_t star_pi = std::string::npos;
  std::size_t star_ti = 0;
  while (pi < pattern.size() || (whole && ti < text.size())) {
    if (pi < pattern.size()) {
      const PatternItem& item = pattern[pi];
      if (item.kind == PatternKind::kStar) {
        star_pi = pi++;
        star_ti = ti;
        continue;
      }
      if (ti < text.size()) {
        const bool hit = item.kind == PatternKind::kLiteral ? item.id == text[ti].id
                                                            : !sbcs_question || byte_count_jajp(text[ti].cp) == 1;
        if (hit) {
          ++pi;
          ++ti;
          continue;
        }
      }
    }
    if (star_pi == std::string::npos || star_ti >= text.size()) {
      return false;
    }
    pi = star_pi + 1;
    ti = ++star_ti;
  }
  return true;
}

}  // namespace

bool collated_wildcard_match(std::string_view pattern, std::string_view text, ExcelProfile profile) {
  CharClasses classes(profile);
  const std::vector<PatternItem> items = parse_pattern(pattern, classes);
  return match_at(items, tokenize(text, 0, classes), 0, /*sbcs_question=*/false, /*whole=*/true);
}

std::size_t collated_wildcard_find(std::string_view pattern, std::string_view text, std::size_t start_byte,
                                   ExcelProfile profile, bool sbcs_question) {
  CharClasses classes(profile);
  const std::vector<PatternItem> items = parse_pattern(pattern, classes);
  const std::vector<TextToken> tokens = tokenize(text, start_byte, classes);
  for (std::size_t start = 0; start <= tokens.size(); ++start) {
    if (match_at(items, tokens, start, sbcs_question, /*whole=*/false)) {
      return start < tokens.size() ? tokens[start].byte : text.size();
    }
  }
  return std::string_view::npos;
}

}  // namespace eval
}  // namespace formulon
