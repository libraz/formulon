#include "eval/collated_wildcard.h"

#include <cstdint>
#include <memory>
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

// Collation class of the second UTF-16 unit of a supplementary-plane
// character; no pattern literal but that same second unit equals it.
constexpr std::uint32_t kSecondUnitId = 0xFFFFFFFFU;

// One text UTF-16 unit under the `=` equivalence: `id` names its collation
// class, `cp` / `byte` locate the source character. A supplementary-plane
// character is two tokens, so `?` matches half of it and `??` all of it.
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
// never span characters, matching SEARCH (Mac Excel 365).
class CharClasses {
 public:
  explicit CharClasses(ExcelProfile profile) : profile_(profile) {}

  std::uint32_t id_of(std::uint32_t cp) {
    const auto cached = by_char_.find(cp);
    if (cached != by_char_.end()) {
      return cached->second;
    }
    const std::string key = collation_key(encode_utf8_codepoint(cp), profile_, CollationStrength::kCompare);
    const auto inserted = by_key_.emplace(key, static_cast<std::uint32_t>(by_key_.size()));
    by_char_.emplace(cp, inserted.first->second);
    return inserted.first->second;
  }

 private:
  ExcelProfile profile_;
  std::unordered_map<std::uint32_t, std::uint32_t> by_char_;
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
      if (cp > 0xFFFFU) {
        out.push_back({PatternKind::kLiteral, kSecondUnitId});
      }
    }
    escaped = false;
  }
  return out;
}

void tokenize(std::string_view text, std::size_t start_byte, CharClasses& classes, std::vector<TextToken>& out) {
  out.clear();
  for (std::size_t i = start_byte; i < text.size();) {
    std::size_t step = 0;
    const std::uint32_t cp = decode_utf8_step(text, i, &step);
    out.push_back({classes.id_of(cp), cp, i});
    if (cp > 0xFFFFU) {
      out.push_back({kSecondUnitId, cp, i});
    }
    i += step;
  }
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

struct CollatedWildcard::Impl {
  Impl(std::string_view pattern, ExcelProfile profile) : classes(profile), items(parse_pattern(pattern, classes)) {}

  CharClasses classes;
  std::vector<PatternItem> items;
  std::vector<TextToken> tokens;  // scratch reused across calls
};

CollatedWildcard::CollatedWildcard(std::string_view pattern, ExcelProfile profile)
    : impl_(std::make_unique<Impl>(pattern, profile)) {}

CollatedWildcard::~CollatedWildcard() = default;
CollatedWildcard::CollatedWildcard(CollatedWildcard&&) noexcept = default;
CollatedWildcard& CollatedWildcard::operator=(CollatedWildcard&&) noexcept = default;

bool CollatedWildcard::match(std::string_view text) const {
  Impl& m = *impl_;
  tokenize(text, 0, m.classes, m.tokens);
  return match_at(m.items, m.tokens, 0, /*sbcs_question=*/false, /*whole=*/true);
}

std::size_t CollatedWildcard::find(std::string_view text, std::size_t start_byte, bool sbcs_question) const {
  Impl& m = *impl_;
  tokenize(text, start_byte, m.classes, m.tokens);
  for (std::size_t start = 0; start <= m.tokens.size(); ++start) {
    if (start < m.tokens.size() && m.tokens[start].id == kSecondUnitId) {
      continue;  // a match never begins inside a character
    }
    if (match_at(m.items, m.tokens, start, sbcs_question, /*whole=*/false)) {
      return start < m.tokens.size() ? m.tokens[start].byte : text.size();
    }
  }
  return std::string_view::npos;
}

bool collated_wildcard_match(std::string_view pattern, std::string_view text, ExcelProfile profile) {
  return CollatedWildcard(pattern, profile).match(text);
}

std::size_t collated_wildcard_find(std::string_view pattern, std::string_view text, std::size_t start_byte,
                                   ExcelProfile profile, bool sbcs_question) {
  return CollatedWildcard(pattern, profile).find(text, start_byte, sbcs_question);
}

}  // namespace eval
}  // namespace formulon
