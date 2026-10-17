//
// Unit tests for the `*` / `?` / `~` wildcard matchers declared in
// `wildcard.h`: codepoint-aware matching and the SEARCHB byte-mode `?`.

#include "eval/wildcard.h"

#include <cstddef>
#include <string_view>

#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// wildcard_match: codepoint-aware `*` backtracking and `?` matching.
// ---------------------------------------------------------------------------

TEST(WildcardMatchCodepoint, StarBacktrackAdvancesWholeCodepoints) {
  // Regression: `*` backtracking used to advance one BYTE, letting a
  // following `?` match a UTF-8 continuation byte and spuriously match.
  // `*??c*` against "あcう" must NOT match (only 3 codepoints: あ, c, う —
  // there is no "c" two codepoints in from any start after consuming two).
  EXPECT_FALSE(wildcard_match("*??c*",
                              "\xE3\x81\x82"
                              "c"
                              "\xE3\x81\x86"));
}

TEST(WildcardMatchCodepoint, QuestionMatchesExactlyOneCodepoint) {
  // "?" matches a single multibyte codepoint.
  EXPECT_TRUE(wildcard_match("?", "\xE3\x81\x82"));  // あ
  EXPECT_FALSE(wildcard_match("?",
                              "\xE3\x81\x82"
                              "\xE3\x81\x84"));        // あい (2 cp)
  EXPECT_FALSE(wildcard_match("??", "\xE3\x81\x82"));  // one cp cannot fill two ?
}

TEST(WildcardMatchCodepoint, SupplementaryCharacterTakesTwoQuestionMarks) {
  // A 4-byte codepoint (U+1F600) is two UTF-16 units, so one "?" matches half of it.
  EXPECT_FALSE(wildcard_match("?", "\xF0\x9F\x98\x80"));
  EXPECT_TRUE(wildcard_match("??", "\xF0\x9F\x98\x80"));
  EXPECT_FALSE(wildcard_match("a?b",
                              "a"
                              "\xF0\x9F\x98\x80"
                              "b"));
  EXPECT_TRUE(wildcard_match("a??b",
                             "a"
                             "\xF0\x9F\x98\x80"
                             "b"));
  // A literal never matches half of a character, but a star can stop inside one.
  EXPECT_TRUE(wildcard_match("*?b",
                             "a"
                             "\xF0\x9F\x98\x80"
                             "b"));
  EXPECT_FALSE(wildcard_match("a?\xF0\x9F\x98\x80", "a\xF0\x9F\x98\x80"));
  EXPECT_TRUE(wildcard_match("a*b",
                             "a"
                             "\xF0\x9F\x98\x80"
                             "b"));
}

TEST(WildcardMatchCodepoint, StarSwallowsMultibyteRun) {
  // `*` skips a run of multibyte codepoints and the trailing literal aligns.
  EXPECT_TRUE(wildcard_match("*c",
                             "\xE3\x81\x82"
                             "\xE3\x81\x84"
                             "c"));  // あいc
  EXPECT_TRUE(wildcard_match("*?c",
                             "\xE3\x81\x82"
                             "c"));  // あc
}

}  // namespace
}  // namespace eval
}  // namespace formulon
