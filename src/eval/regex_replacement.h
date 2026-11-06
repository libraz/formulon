//
// Replacement-template tokens for REGEXREPLACE: the token types, the parser
// that turns a replacement string into a token list, and the expansion of a
// token list against one match.

#ifndef FORMULON_EVAL_REGEX_REPLACEMENT_H_
#define FORMULON_EVAL_REGEX_REPLACEMENT_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/regex_program.h"

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

/// Appends the expansion of `tokens` for one match to `out`. `ovec` is the
/// PCRE2 ovector of the match. False when a token is invalid or names a group
/// the pattern lacks.
bool expand_replacement(const std::vector<ReplToken>& tokens, const CompiledPattern& program, std::string_view subject,
                        const PCRE2_SIZE* ovec, std::string& out);

}  // namespace formulon::eval::detail

#endif  // FORMULON_EVAL_REGEX_REPLACEMENT_H_
