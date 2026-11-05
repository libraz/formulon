//
// Replacement-template tokens for REGEXREPLACE: the token types and the
// parser that turns a replacement string into a token list. Expansion against
// a match lives in `regex_lazy.cpp`.

#ifndef FORMULON_EVAL_REGEX_REPLACEMENT_H_
#define FORMULON_EVAL_REGEX_REPLACEMENT_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

namespace formulon::eval::detail {

enum class ReplTokenKind : std::uint8_t {
  kLiteral,
  kGroup,    // `number`
  kNamed,    // `text` is the group name
  kBefore,   // text before the match
  kAfter,    // text after the match
  kSubject,  // the whole subject
  kUpperNext,
  kLowerNext,
  kUpperRun,
  kLowerRun,
  kCaseEnd,
  kInvalid,
};

struct ReplToken {
  ReplTokenKind kind = ReplTokenKind::kLiteral;
  std::uint32_t number = 0;
  std::string text;
};

/// Parses a replacement string into tokens; an invalid escape becomes a
/// `kInvalid` token that fails when a match is expanded.
std::vector<ReplToken> parse_replacement(std::string_view r);

}  // namespace formulon::eval::detail

#endif  // FORMULON_EVAL_REGEX_REPLACEMENT_H_
