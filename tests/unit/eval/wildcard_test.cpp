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

TEST(WildcardMatchCodepoint, QuestionMatchesSupplementaryPlaneEmoji) {
  // A 4-byte codepoint (U+1F600) is one "?" unit.
  EXPECT_TRUE(wildcard_match("?", "\xF0\x9F\x98\x80"));
  EXPECT_TRUE(wildcard_match("a?b",
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

// ---------------------------------------------------------------------------
// wildcard_find_dbcs: byte-mode `?` for SEARCHB. `?` matches only codepoints
// whose ja-JP DBCS cost is 1 (ASCII, half-width katakana); kanji/hiragana/
// full-width punctuation/emoji refuse to match. `*` and `~?`/`~*` retain
// their normal semantics.
// ---------------------------------------------------------------------------

TEST(WildcardFindDbcs, QuestionSkipsKanjiToAsciiOffset) {
  // "漢" is a 3-byte UTF-8 codepoint with DBCS cost 2; `?` refuses to match
  // it. The next codepoint is 'A' at byte offset 3.
  EXPECT_EQ(wildcard_find_dbcs("?", "漢ABC"), std::size_t{3});
}

TEST(WildcardFindDbcs, QuestionMatchesAsciiAtOffsetZero) {
  EXPECT_EQ(wildcard_find_dbcs("?", "abc"), std::size_t{0});
}

TEST(WildcardFindDbcs, QuestionFindsNothingInPureKanji) {
  // No SBCS codepoint anywhere; `?` matches nothing.
  EXPECT_EQ(wildcard_find_dbcs("?", "漢"), std::string_view::npos);
}

TEST(WildcardFindDbcs, QuestionWithLiteralTailAfterKanji) {
  // At start=0, "?" against 漢 fails. At start=3 the suffix is "abc" but
  // pattern "?abc" needs 4 chars and only "abc" remains — no match.
  EXPECT_EQ(wildcard_find_dbcs("?abc", "漢abc"), std::string_view::npos);
}

TEST(WildcardFindDbcs, QuestionInsideAsciiHaystack) {
  EXPECT_EQ(wildcard_find_dbcs("a?c", "abc"), std::size_t{0});
}

TEST(WildcardFindDbcs, StarSwallowsKanjiThenQuestionMatchesAscii) {
  // `*` consumes the leading 漢 (2 DBCS bytes / 3 UTF-8 bytes) and `?`
  // then matches 'A'.
  EXPECT_EQ(wildcard_find_dbcs("*?", "漢A"), std::size_t{0});
}

TEST(WildcardFindDbcs, EscapedQuestionIsLiteralRegardlessOfMode) {
  // `~?` is a literal '?' under both wildcard modes; the byte-mode flag
  // only affects unescaped `?`.
  EXPECT_EQ(wildcard_find_dbcs("~?", "x?y"), std::size_t{1});
}

TEST(WildcardFindDbcs, QuestionMatchesHalfWidthKatakana) {
  // ｱ (U+FF71) is half-width katakana, classified as DBCS=1 in ja-JP, so
  // byte-mode `?` matches it at offset 0.
  EXPECT_EQ(wildcard_find_dbcs("?", "ｱABC"), std::size_t{0});
}

}  // namespace
}  // namespace eval
}  // namespace formulon
