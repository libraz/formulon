// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the byte-oriented text built-ins: LENB, LEFTB,
// RIGHTB, MIDB, CHAR, CODE. Pins the Mac Excel 365 (ja-JP) DBCS byte-count
// rule across ASCII, half-width katakana, BMP CJK, and supplementary-plane
// codepoints, plus the "pad with ASCII space on a 1-byte overflow of a
// 2-byte character" rule shared by LEFTB / RIGHTB / MIDB.

#include <string_view>

#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry. Mirrors
// the helper used by `builtins_text_test.cpp`.
Value EvalSource(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena);
}

// ---------------------------------------------------------------------------
// LENB
// ---------------------------------------------------------------------------

TEST(TextLenb, AsciiEachByteIsOne) {
  const Value v = EvalSource("=LENB(\"hello\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(TextLenb, EmptyStringIsZero) {
  const Value v = EvalSource("=LENB(\"\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(TextLenb, HiraganaEachIsTwoBytes) {
  // "あいう" = 3 hiragana chars * 2 bytes.
  const Value v = EvalSource("=LENB(\"あいう\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(TextLenb, KanjiEachIsTwoBytes) {
  // "日本語" = 3 kanji chars * 2 bytes.
  const Value v = EvalSource("=LENB(\"日本語\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(TextLenb, MixedAsciiAndHiragana) {
  // "abcあ" = 3 ASCII (1 byte each) + 1 hiragana (2 bytes) = 5.
  const Value v = EvalSource("=LENB(\"abcあ\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(TextLenb, HalfWidthKatakanaEachIsOneByte) {
  // "ｱｲｳ" (half-width katakana) = 3 chars * 1 byte each.
  const Value v = EvalSource("=LENB(\"ｱｲｳ\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(TextLenb, FullWidthDigitsAreTwoBytes) {
  // Full-width digits U+FF10..U+FF19 are outside the half-width katakana
  // block so they should count as 2 bytes each.
  const Value v = EvalSource("=LENB(\"１２３\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(TextLenb, SupplementaryPlaneIsTwoBytes) {
  // "😀" U+1F600: Mac Excel ja-JP counts supplementary-plane codepoints as
  // 2 bytes (oracle-verified 2026-04-23), not 4 as a naive
  // surrogate-pair-times-two would suggest.
  const Value v = EvalSource("=LENB(\"😀\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(TextLenb, CoercesNumber) {
  // LENB(123) -> LENB("123") -> 3.
  const Value v = EvalSource("=LENB(123)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(TextLenb, ErrorPropagates) {
  const Value v = EvalSource("=LENB(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// LEFTB
// ---------------------------------------------------------------------------

TEST(TextLeftb, AsciiFirstThree) {
  const Value v = EvalSource("=LEFTB(\"hello\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hel");
}

TEST(TextLeftb, DefaultCountIsOne) {
  const Value v = EvalSource("=LEFTB(\"hello\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "h");
}

TEST(TextLeftb, HiraganaExactBoundary) {
  // "あいう" 4 bytes -> 2 full hiragana chars.
  const Value v = EvalSource("=LEFTB(\"あいう\", 4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "あい");
}

TEST(TextLeftb, HiraganaSplitPadsSpace) {
  // Budget = 3: one full char (2 bytes) + 1 byte remaining; the next char
  // costs 2 bytes, so we pad with a single ASCII space.
  const Value v = EvalSource("=LEFTB(\"あいう\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "あ ");
}

TEST(TextLeftb, ZeroCount) {
  const Value v = EvalSource("=LEFTB(\"hello\", 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextLeftb, NegativeCountIsValueError) {
  const Value v = EvalSource("=LEFTB(\"hello\", -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextLeftb, CountExceedsLength) {
  const Value v = EvalSource("=LEFTB(\"ab\", 100)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ab");
}

TEST(TextLeftb, MixedAsciiHiragana) {
  // "aあb" = 'a' (1) + 'あ' (2) + 'b' (1) = 4 bytes.
  // LEFTB(...,2) consumes 'a' (1) then tries 'あ' (2); 1 byte remaining,
  // so pad with a space. Expected: "a ".
  const Value v = EvalSource("=LEFTB(\"aあb\", 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a ");
}

// ---------------------------------------------------------------------------
// RIGHTB
// ---------------------------------------------------------------------------

TEST(TextRightb, AsciiLastThree) {
  const Value v = EvalSource("=RIGHTB(\"hello\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "llo");
}

TEST(TextRightb, DefaultCountIsOne) {
  const Value v = EvalSource("=RIGHTB(\"hello\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "o");
}

TEST(TextRightb, HiraganaExactBoundary) {
  // "あいう" -> rightmost 4 bytes -> "いう".
  const Value v = EvalSource("=RIGHTB(\"あいう\", 4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "いう");
}

TEST(TextRightb, HiraganaSplitPadsSpace) {
  // "あいう" with budget 3: rightmost full "う" (2 bytes) + 1 byte; next
  // char costs 2 -> pad with space on the left.
  const Value v = EvalSource("=RIGHTB(\"あいう\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), " う");
}

TEST(TextRightb, ZeroCount) {
  const Value v = EvalSource("=RIGHTB(\"hello\", 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextRightb, NegativeCountIsValueError) {
  const Value v = EvalSource("=RIGHTB(\"hello\", -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextRightb, CountExceedsLength) {
  const Value v = EvalSource("=RIGHTB(\"ab\", 100)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ab");
}

// ---------------------------------------------------------------------------
// MIDB
// ---------------------------------------------------------------------------

TEST(TextMidb, AsciiMiddleWindow) {
  const Value v = EvalSource("=MIDB(\"hello\", 2, 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ell");
}

TEST(TextMidb, HiraganaCleanBoundary) {
  // "あいうえお" bytes: [あ:1-2][い:3-4][う:5-6][え:7-8][お:9-10].
  // MIDB(...,3,4) = bytes 3..6 = "いう".
  const Value v = EvalSource("=MIDB(\"あいうえお\", 3, 4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "いう");
}

TEST(TextMidb, HiraganaStartSplitsCharPadsSpace) {
  // MIDB(...,2,4): start_byte=2 falls inside "あ" (bytes 1..2). Excel pads
  // the first byte with a space and continues from byte 3. Remaining budget
  // = 3 bytes: include "い" (2) then 1 byte overflow on "う" -> trailing
  // space. Expected: " い ".
  const Value v = EvalSource("=MIDB(\"あいうえお\", 2, 4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), " い ");
}

TEST(TextMidb, StartBeyondLengthYieldsEmpty) {
  const Value v = EvalSource("=MIDB(\"abc\", 10, 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextMidb, StartZeroIsValueError) {
  const Value v = EvalSource("=MIDB(\"abc\", 0, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextMidb, NegativeNumBytesIsValueError) {
  const Value v = EvalSource("=MIDB(\"abc\", 1, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextMidb, ZeroLengthYieldsEmpty) {
  const Value v = EvalSource("=MIDB(\"abc\", 2, 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

// ---------------------------------------------------------------------------
// CHAR
// ---------------------------------------------------------------------------

TEST(TextChar, AsciiLetterA) {
  const Value v = EvalSource("=CHAR(65)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "A");
}

TEST(TextChar, Space) {
  const Value v = EvalSource("=CHAR(32)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), " ");
}

TEST(TextChar, ZeroIsValueError) {
  const Value v = EvalSource("=CHAR(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextChar, TooLargeIsValueError) {
  const Value v = EvalSource("=CHAR(256)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextChar, NegativeIsValueError) {
  const Value v = EvalSource("=CHAR(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextChar, TruncatesFractional) {
  // 65.9 truncates to 65 -> "A".
  const Value v = EvalSource("=CHAR(65.9)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "A");
}

TEST(TextChar, HalfWidthKatakanaViaCp932) {
  // CP932 maps 0xA9 to half-width katakana small-u U+FF69.
  // 169 - 0xA1 = 8, so U+FF61 + 8 = U+FF69 = "ｩ".
  // UTF-8 encoding of U+FF69 = 0xEF 0xBD 0xA9.
  const Value v = EvalSource("=CHAR(169)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xEF\xBD\xA9");
}

TEST(TextChar, HalfWidthKatakanaFirstSlot) {
  // 0xA1 -> U+FF61 = "｡" (half-width ideographic full stop).
  const Value v = EvalSource("=CHAR(161)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xEF\xBD\xA1");
}

// ---------------------------------------------------------------------------
// CODE
// ---------------------------------------------------------------------------

TEST(TextCode, AsciiLetterA) {
  const Value v = EvalSource("=CODE(\"A\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 65.0);
}

TEST(TextCode, EmptyStringIsValueError) {
  const Value v = EvalSource("=CODE(\"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextCode, UsesFirstCharOnly) {
  const Value v = EvalSource("=CODE(\"Apple\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 65.0);
}

TEST(TextCode, HiraganaReturnsUnicode) {
  // "あ" = U+3042 = 12354 decimal. Mac Excel ja-JP historically returned
  // CP932 for non-ASCII; this test pins Formulon's current Unicode
  // behaviour. If the oracle diverges the result will be captured via
  // tests/divergence.yaml and this assertion stays the contract.
  const Value v = EvalSource("=CODE(\"あ\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 12354.0);
}

TEST(TextCode, ErrorPropagates) {
  const Value v = EvalSource("=CODE(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
