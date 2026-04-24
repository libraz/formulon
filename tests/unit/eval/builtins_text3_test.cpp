// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the third batch of text built-ins: REPLACE,
// REPLACEB, FINDB, SEARCHB, TEXTBEFORE, TEXTAFTER, FIXED, DOLLAR. Each
// test parses a formula source, evaluates the AST through the default
// registry, and asserts the resulting Value. REPLACEB / FINDB / SEARCHB
// are exercised against hiragana (2 DBCS bytes each) to pin the
// multi-byte-character path; DOLLAR pins the exact UTF-8 byte sequence
// including the yen sign prefix `¥` (0xC2 0xA5) so any oracle divergence
// is detectable.

#include <string>
#include <string_view>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

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
// Registry pin: all eight names are registered.
// ---------------------------------------------------------------------------

TEST(BuiltinsText3Registry, AllNamesRegistered) {
  const FunctionRegistry& r = default_registry();
  for (const char* name :
       {"REPLACE", "REPLACEB", "FINDB", "SEARCHB", "TEXTBEFORE", "TEXTAFTER", "FIXED", "DOLLAR", "HYPERLINK"}) {
    EXPECT_NE(r.lookup(name), nullptr) << "missing registration: " << name;
  }
}

// ---------------------------------------------------------------------------
// REPLACE
// ---------------------------------------------------------------------------

TEST(BuiltinsText3Replace, Middle) {
  const Value v = EvalSource("=REPLACE(\"abcdef\", 2, 3, \"XYZ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "aXYZef");
}

TEST(BuiltinsText3Replace, StartPastEndAppends) {
  const Value v = EvalSource("=REPLACE(\"abc\", 5, 2, \"XYZ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abcXYZ");
}

TEST(BuiltinsText3Replace, EmptyNewText) {
  const Value v = EvalSource("=REPLACE(\"abcdef\", 2, 3, \"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "aef");
}

TEST(BuiltinsText3Replace, StartOneReplacesPrefix) {
  const Value v = EvalSource("=REPLACE(\"abcdef\", 1, 2, \"X\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Xcdef");
}

TEST(BuiltinsText3Replace, StartZeroIsValueError) {
  const Value v = EvalSource("=REPLACE(\"abc\", 0, 1, \"X\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText3Replace, NegativeNumCharsIsValueError) {
  const Value v = EvalSource("=REPLACE(\"abc\", 1, -1, \"X\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// REPLACEB (each hiragana = 2 DBCS bytes)
// ---------------------------------------------------------------------------

TEST(BuiltinsText3ReplaceB, HiraganaMiddle) {
  // "あいう" = 3 hiragana (6 DBCS bytes). Replace 2 bytes starting at
  // position 3 (which is the start of "い") with "X": "あXう".
  const Value v = EvalSource("=REPLACEB(\"\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86\", 3, 2, \"X\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82X\xE3\x81\x86");
}

TEST(BuiltinsText3ReplaceB, HeadPadWhenStartMidChar) {
  // "あい" (DBCS: 1-2 = あ, 3-4 = い). start_byte=2 lands mid-"あ".
  // MIDB-head-pad: prefix becomes " " (replacing "あ"), then budget=2
  // deletes bytes from the rest of "あ" (1 byte) plus 1 byte into "い".
  // That straddles "い" -> tail pad ' '. Result: " " + "X" + " ".
  const Value v = EvalSource("=REPLACEB(\"\xE3\x81\x82\xE3\x81\x84\", 2, 2, \"X\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), " X ");
}

TEST(BuiltinsText3ReplaceB, AsciiRoundTripsAsReplace) {
  // ASCII-only input should behave identically to REPLACE.
  const Value v = EvalSource("=REPLACEB(\"abcdef\", 2, 3, \"XYZ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "aXYZef");
}

TEST(BuiltinsText3ReplaceB, StartZeroIsValueError) {
  const Value v = EvalSource("=REPLACEB(\"abc\", 0, 1, \"X\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText3ReplaceB, NegativeBytesIsValueError) {
  const Value v = EvalSource("=REPLACEB(\"abc\", 1, -1, \"X\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// FINDB
// ---------------------------------------------------------------------------

TEST(BuiltinsText3FindB, Hiragana) {
  // Each hiragana is 2 DBCS bytes: "あ"=1..2, "い"=3..4, "う"=5..6.
  const Value v = EvalSource("=FINDB(\"\xE3\x81\x86\", \"\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsText3FindB, Ascii) {
  const Value v = EvalSource("=FINDB(\"c\", \"abcdef\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsText3FindB, NotFoundIsValueError) {
  const Value v = EvalSource("=FINDB(\"z\", \"abcdef\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText3FindB, EmptyNeedleReturnsStart) {
  const Value v = EvalSource("=FINDB(\"\", \"abc\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsText3FindB, EmptyNeedleWithExplicitStart) {
  const Value v = EvalSource("=FINDB(\"\", \"abc\", 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsText3FindB, StartZeroIsValueError) {
  const Value v = EvalSource("=FINDB(\"b\", \"abc\", 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText3FindB, CaseSensitive) {
  const Value v = EvalSource("=FINDB(\"B\", \"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// SEARCHB
// ---------------------------------------------------------------------------

TEST(BuiltinsText3SearchB, HiraganaWildcard) {
  // "い*" matches "い" onward. "あいう": "い" begins at DBCS 3.
  const Value v = EvalSource("=SEARCHB(\"\xE3\x81\x84*\", \"\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsText3SearchB, CaseInsensitive) {
  const Value v = EvalSource("=SEARCHB(\"B\", \"abc\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsText3SearchB, QuestionMarkWildcard) {
  const Value v = EvalSource("=SEARCHB(\"a?c\", \"xabc\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsText3SearchB, NotFoundIsValueError) {
  const Value v = EvalSource("=SEARCHB(\"z\", \"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText3SearchB, EmptyNeedleReturnsStart) {
  const Value v = EvalSource("=SEARCHB(\"\", \"abc\", 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// TEXTBEFORE
// ---------------------------------------------------------------------------

TEST(BuiltinsText3TextBefore, FirstDelimiter) {
  const Value v = EvalSource("=TEXTBEFORE(\"abc-def-ghi\", \"-\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(BuiltinsText3TextBefore, SecondInstance) {
  const Value v = EvalSource("=TEXTBEFORE(\"abc-def-ghi\", \"-\", 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc-def");
}

TEST(BuiltinsText3TextBefore, NegativeInstanceLast) {
  const Value v = EvalSource("=TEXTBEFORE(\"abc-def-ghi\", \"-\", -1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc-def");
}

TEST(BuiltinsText3TextBefore, CaseInsensitive) {
  const Value v = EvalSource("=TEXTBEFORE(\"ABC\", \"b\", 1, 1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "A");
}

TEST(BuiltinsText3TextBefore, MatchEndPositiveInstanceUsesEndSentinel) {
  // match_end=1 with positive instance: only the end-of-text virtual match is
  // active. No actual "-" in "abc", so the Nth=1 match is the end sentinel
  // and TEXTBEFORE returns the full text.
  const Value v = EvalSource("=TEXTBEFORE(\"abc\", \"-\", 1, 0, 1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(BuiltinsText3TextBefore, MatchEndNegativeInstanceUsesStartSentinel) {
  // match_end=1 with negative instance: only the start-of-text virtual match
  // is active. The Nth=-1 match from the end is the start sentinel, so
  // TEXTBEFORE returns everything before position 0 = "".
  const Value v = EvalSource("=TEXTBEFORE(\"abc\", \"-\", -1, 0, 1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsText3TextBefore, NotFoundDefaultIsNA) {
  const Value v = EvalSource("=TEXTBEFORE(\"abc\", \"-\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsText3TextBefore, IfNotFoundCustom) {
  const Value v = EvalSource("=TEXTBEFORE(\"abc\", \"-\", 1, 0, 0, \"NA\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "NA");
}

TEST(BuiltinsText3TextBefore, InstanceZeroIsValueError) {
  const Value v = EvalSource("=TEXTBEFORE(\"abc-def\", \"-\", 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText3TextBefore, EmptyDelimiterIsValueError) {
  const Value v = EvalSource("=TEXTBEFORE(\"abc\", \"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// TEXTAFTER
// ---------------------------------------------------------------------------

TEST(BuiltinsText3TextAfter, FirstDelimiter) {
  const Value v = EvalSource("=TEXTAFTER(\"abc-def-ghi\", \"-\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "def-ghi");
}

TEST(BuiltinsText3TextAfter, NegativeInstanceLast) {
  const Value v = EvalSource("=TEXTAFTER(\"abc-def-ghi\", \"-\", -1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ghi");
}

TEST(BuiltinsText3TextAfter, MatchEndNegativeInstanceUsesStartSentinel) {
  // match_end=1 with negative instance: only the start-of-text sentinel is
  // active. TEXTAFTER of the start sentinel returns the full text "abc".
  const Value v = EvalSource("=TEXTAFTER(\"abc\", \"-\", -1, 0, 1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(BuiltinsText3TextAfter, NotFoundDefaultIsNA) {
  const Value v = EvalSource("=TEXTAFTER(\"abc\", \"-\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsText3TextAfter, IfNotFoundCustom) {
  const Value v = EvalSource("=TEXTAFTER(\"abc\", \"-\", 1, 0, 0, \"NA\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "NA");
}

// ---------------------------------------------------------------------------
// FIXED
// ---------------------------------------------------------------------------

TEST(BuiltinsText3Fixed, PositiveDefaultDecimals) {
  const Value v = EvalSource("=FIXED(1234567.891)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1,234,567.89");
}

TEST(BuiltinsText3Fixed, PositiveExplicitDecimals) {
  const Value v = EvalSource("=FIXED(1234567.891, 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1,234,567.89");
}

TEST(BuiltinsText3Fixed, ZeroDecimals) {
  const Value v = EvalSource("=FIXED(1234.56, 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1,235");
}

TEST(BuiltinsText3Fixed, NegativeDecimalsRoundsLeft) {
  const Value v = EvalSource("=FIXED(1234.56, -2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1,200");
}

TEST(BuiltinsText3Fixed, NoCommasTrue) {
  const Value v = EvalSource("=FIXED(1234567.891, -3, TRUE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1235000");
}

TEST(BuiltinsText3Fixed, NoCommasFalse) {
  const Value v = EvalSource("=FIXED(1234567.891, 2, FALSE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1,234,567.89");
}

TEST(BuiltinsText3Fixed, NegativeValue) {
  const Value v = EvalSource("=FIXED(-1234.5, 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "-1,234.50");
}

TEST(BuiltinsText3Fixed, DecimalsBeyondCapIsValueError) {
  const Value v = EvalSource("=FIXED(1.5, 200)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText3Fixed, ErrorPropagates) {
  const Value v = EvalSource("=FIXED(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// DOLLAR
// ---------------------------------------------------------------------------

TEST(BuiltinsText3Dollar, PositiveYenSign) {
  const Value v = EvalSource("=DOLLAR(1234.5)");
  ASSERT_TRUE(v.is_text());
  // UTF-8: "¥1,234.50" = 0xC2 0xA5 + "1,234.50"
  EXPECT_EQ(v.as_text(), std::string("\xC2\xA5"
                                     "1,234.50"));
}

TEST(BuiltinsText3Dollar, NegativeParens) {
  const Value v = EvalSource("=DOLLAR(-1234.5)");
  ASSERT_TRUE(v.is_text());
  // UTF-8: "(¥1,234.50)"
  EXPECT_EQ(v.as_text(), std::string("(\xC2\xA5"
                                     "1,234.50)"));
}

TEST(BuiltinsText3Dollar, Zero) {
  const Value v = EvalSource("=DOLLAR(0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string("\xC2\xA5"
                                     "0.00"));
}

TEST(BuiltinsText3Dollar, ZeroDecimals) {
  // Use 1234.7 to avoid platform-dependent tie-rounding (snprintf's %.*f on
  // macOS rounds half-to-even, giving 1234 for 1234.5).
  const Value v = EvalSource("=DOLLAR(1234.7, 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string("\xC2\xA5"
                                     "1,235"));
}

TEST(BuiltinsText3Dollar, NegativeDecimals) {
  // DOLLAR(1234.5, -2) rounds to 1200 first, then formats without fraction.
  const Value v = EvalSource("=DOLLAR(1234.5, -2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string("\xC2\xA5"
                                     "1,200"));
}

TEST(BuiltinsText3Dollar, DecimalsBeyondCapIsValueError) {
  const Value v = EvalSource("=DOLLAR(1.5, 200)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText3Dollar, ErrorPropagates) {
  const Value v = EvalSource("=DOLLAR(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// HYPERLINK
// ---------------------------------------------------------------------------

TEST(BuiltinsText3Hyperlink, OneArgReturnsLink) {
  const Value v = EvalSource("=HYPERLINK(\"http://example.com\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "http://example.com");
}

TEST(BuiltinsText3Hyperlink, TwoArgReturnsFriendlyName) {
  const Value v = EvalSource("=HYPERLINK(\"http://example.com\", \"Click here\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Click here");
}

TEST(BuiltinsText3Hyperlink, EmptyFriendlyNameWins) {
  // An explicit "" as friendly_name beats the link, matching Excel's
  // two-arg semantics (the caller asked for the friendly name branch).
  const Value v = EvalSource("=HYPERLINK(\"http://example.com\", \"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsText3Hyperlink, NumericLinkCoercesToText) {
  const Value v = EvalSource("=HYPERLINK(123)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "123");
}

TEST(BuiltinsText3Hyperlink, NumericFriendlyCoercesToText) {
  const Value v = EvalSource("=HYPERLINK(\"x\", 42)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "42");
}

TEST(BuiltinsText3Hyperlink, ErrorInLinkPropagates) {
  const Value v = EvalSource("=HYPERLINK(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
