// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the second batch of text built-ins: TEXTJOIN,
// UNICHAR, UNICODE, CLEAN, PROPER. Each test parses a formula source,
// evaluates the AST through the default registry, and asserts the
// resulting Value.
//
// Note: CHAR / CODE are intentionally NOT covered here - they require an
// Excel ja-JP codepage decision tracked separately.

#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

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

// Parses `src` and evaluates it against a bound workbook + current sheet.
// Used by tests that exercise range expansion through `expand_range`.
Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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
  EvalState state;
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// TEXTJOIN
// ---------------------------------------------------------------------------

TEST(BuiltinsText2TextJoin, BasicJoinIgnoreEmptyTrue) {
  const Value v = EvalSource("=TEXTJOIN(\",\", TRUE, \"a\", \"b\", \"c\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a,b,c");
}

TEST(BuiltinsText2TextJoin, IgnoreEmptyTrueSkipsEmptyArgs) {
  // The empty middle argument is skipped, so the result has no double comma.
  const Value v = EvalSource("=TEXTJOIN(\",\", TRUE, \"a\", \"\", \"b\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a,b");
}

TEST(BuiltinsText2TextJoin, IgnoreEmptyFalseKeepsEmptyArgs) {
  // With ignore_empty=FALSE, the empty argument participates and yields
  // adjacent delimiters.
  const Value v = EvalSource("=TEXTJOIN(\",\", FALSE, \"a\", \"\", \"b\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a,,b");
}

TEST(BuiltinsText2TextJoin, IgnoreEmptyFalseAllEmpty) {
  // Three empty args + ignore_empty=FALSE -> two delimiters between them.
  const Value v = EvalSource("=TEXTJOIN(\"-\", FALSE, \"\", \"\", \"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "--");
}

TEST(BuiltinsText2TextJoin, IgnoreEmptyTrueAllEmptyYieldsEmpty) {
  const Value v = EvalSource("=TEXTJOIN(\"-\", TRUE, \"\", \"\", \"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsText2TextJoin, MultiCharDelimiter) {
  const Value v = EvalSource("=TEXTJOIN(\" / \", TRUE, \"x\", \"y\", \"z\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "x / y / z");
}

TEST(BuiltinsText2TextJoin, EmptyDelimiterConcatenates) {
  const Value v = EvalSource("=TEXTJOIN(\"\", TRUE, \"a\", \"b\", \"c\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(BuiltinsText2TextJoin, NumberArgCoercesToText) {
  // 42 is rendered "42" and joined.
  const Value v = EvalSource("=TEXTJOIN(\":\", TRUE, \"x\", 42, \"y\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "x:42:y");
}

TEST(BuiltinsText2TextJoin, BoolArgCoercesToText) {
  const Value v = EvalSource("=TEXTJOIN(\",\", TRUE, TRUE, FALSE, \"a\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "TRUE,FALSE,a");
}

TEST(BuiltinsText2TextJoin, ErrorInTextArgPropagates) {
  const Value v = EvalSource("=TEXTJOIN(\",\", TRUE, \"a\", #REF!, \"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsText2TextJoin, ErrorInDelimiterPropagates) {
  const Value v = EvalSource("=TEXTJOIN(#REF!, TRUE, \"a\", \"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsText2TextJoin, IgnoreEmptyAsNumberOneIsTrue) {
  // Numeric 1 coerces to TRUE for ignore_empty.
  const Value v = EvalSource("=TEXTJOIN(\",\", 1, \"a\", \"\", \"b\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a,b");
}

TEST(BuiltinsText2TextJoin, IgnoreEmptyAsNumberZeroIsFalse) {
  const Value v = EvalSource("=TEXTJOIN(\",\", 0, \"a\", \"\", \"b\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a,,b");
}

TEST(BuiltinsText2TextJoin, ResultExceedingCapIsValueError) {
  // Each REPT yields 32,000 ASCII bytes (= 32,000 UTF-16 units). Two of
  // them plus a 1-byte delimiter sums to 64,001 units, well above the
  // 32,767-unit cap.
  const Value v = EvalSource("=TEXTJOIN(\",\", TRUE, REPT(\"a\", 32000), REPT(\"b\", 32000))");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText2TextJoin, JustAtCapIsAllowed) {
  // 32,766 bytes + "" + 1-byte delimiter? Actually we want exactly 32,767:
  // first arg 16,383 chars + delim 1 + second arg 16,383 chars = 32,767.
  const Value v = EvalSource("=TEXTJOIN(\",\", TRUE, REPT(\"a\", 16383), REPT(\"b\", 16383))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text().size(), 32767u);
}

TEST(BuiltinsText2TextJoin, TooFewArgsIsArityError) {
  // Two args (delimiter, ignore_empty) but no text values - registry rejects
  // because min_arity = 3.
  const Value v = EvalSource("=TEXTJOIN(\",\", TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText2TextJoin, ZeroArgsIsArityError) {
  const Value v = EvalSource("=TEXTJOIN()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText2TextJoin, BlankArgIsTreatedAsEmpty) {
  // The "" middle arg parses as a Text("") and is skipped under
  // ignore_empty=TRUE just like an empty literal would be.
  const Value v = EvalSource("=TEXTJOIN(\".\", TRUE, \"a\", \"\", \"b\", \"\", \"c\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a.b.c");
}

TEST(BuiltinsText2TextJoin, MultiByteDelimiterAndArgs) {
  // Japanese delimiter "、" (U+3001, 3 UTF-8 bytes) joining Japanese args.
  const Value v = EvalSource("=TEXTJOIN(\"\xE3\x80\x81\", TRUE, \"\xE3\x81\x82\", \"\xE3\x81\x84\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82\xE3\x80\x81\xE3\x81\x84");
}

TEST(BuiltinsText2TextJoin, RangeArgument) {
  // TEXTJOIN registers `accepts_ranges = true`, so A1:A3 should flatten to
  // three text values and join as "a,b,c".
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("c"));
  const Value v = EvalSourceIn("=TEXTJOIN(\",\", TRUE, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a,b,c");
}

TEST(BuiltinsText2TextJoin, RangeArgumentSkipsBlanksWhenIgnoreEmpty) {
  // A2 is blank: under ignore_empty=TRUE the range flattens as
  // [a, "", c] and the empty projection of Blank is skipped.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("a"));
  // row 1 left blank
  wb.sheet(0).set_cell_value(2, 0, Value::text("c"));
  const Value v = EvalSourceIn("=TEXTJOIN(\",\", TRUE, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a,c");
}

// ---------------------------------------------------------------------------
// UNICHAR
// ---------------------------------------------------------------------------

TEST(BuiltinsText2Unichar, AsciiUppercaseA) {
  const Value v = EvalSource("=UNICHAR(65)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "A");
}

TEST(BuiltinsText2Unichar, AsciiLowestPrintable) {
  // U+0020 SPACE.
  const Value v = EvalSource("=UNICHAR(32)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), " ");
}

TEST(BuiltinsText2Unichar, BmpHiraganaA) {
  // U+3042 "あ" -> 3-byte UTF-8.
  const Value v = EvalSource("=UNICHAR(12354)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82");
}

TEST(BuiltinsText2Unichar, SupplementaryEmoji) {
  // U+1F600 "😀" -> 4-byte UTF-8.
  const Value v = EvalSource("=UNICHAR(128512)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xF0\x9F\x98\x80");
}

TEST(BuiltinsText2Unichar, MaxValidCodepoint) {
  // U+10FFFF -> 4-byte UTF-8 0xF4 0x8F 0xBF 0xBF.
  const Value v = EvalSource("=UNICHAR(1114111)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xF4\x8F\xBF\xBF");
}

TEST(BuiltinsText2Unichar, ZeroIsValueError) {
  const Value v = EvalSource("=UNICHAR(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText2Unichar, NegativeIsValueError) {
  const Value v = EvalSource("=UNICHAR(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText2Unichar, AboveMaxIsValueError) {
  // 0x110000 is just past the max codepoint.
  const Value v = EvalSource("=UNICHAR(1114112)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText2Unichar, SurrogateLowEndIsValueError) {
  // 0xD800 - first UTF-16 surrogate half.
  const Value v = EvalSource("=UNICHAR(55296)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText2Unichar, SurrogateHighEndIsValueError) {
  // 0xDFFF - last UTF-16 surrogate half.
  const Value v = EvalSource("=UNICHAR(57343)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText2Unichar, FractionalArgTruncates) {
  // 65.9 -> truncate -> 65 -> "A".
  const Value v = EvalSource("=UNICHAR(65.9)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "A");
}

TEST(BuiltinsText2Unichar, BoolArgCoercesToOne) {
  // TRUE -> 1 -> U+0001 (a control character, but still a valid codepoint).
  const Value v = EvalSource("=UNICHAR(TRUE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string("\x01", 1));
}

TEST(BuiltinsText2Unichar, TextNumberCoerces) {
  const Value v = EvalSource("=UNICHAR(\"66\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "B");
}

TEST(BuiltinsText2Unichar, ErrorPropagates) {
  const Value v = EvalSource("=UNICHAR(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// UNICODE
// ---------------------------------------------------------------------------

TEST(BuiltinsText2Unicode, AsciiUppercaseA) {
  const Value v = EvalSource("=UNICODE(\"A\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 65.0);
}

TEST(BuiltinsText2Unicode, BmpHiraganaA) {
  const Value v = EvalSource("=UNICODE(\"\xE3\x81\x82\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12354.0);
}

TEST(BuiltinsText2Unicode, SupplementaryEmojiReturnsCodepointNotSurrogate) {
  // "😀" U+1F600 = 128512. Not the high surrogate value 0xD83D = 55357.
  const Value v = EvalSource("=UNICODE(\"\xF0\x9F\x98\x80\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 128512.0);
}

TEST(BuiltinsText2Unicode, EmptyStringIsValueError) {
  const Value v = EvalSource("=UNICODE(\"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsText2Unicode, MultiCharOnlyFirstReturned) {
  const Value v = EvalSource("=UNICODE(\"ABC\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 65.0);
}

TEST(BuiltinsText2Unicode, MultiCharJapaneseOnlyFirst) {
  // "あい" -> first char "あ" -> 12354.
  const Value v = EvalSource("=UNICODE(\"\xE3\x81\x82\xE3\x81\x84\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12354.0);
}

TEST(BuiltinsText2Unicode, NumericArgCoercedToText) {
  // 123 -> text "123" -> first char '1' -> 49.
  const Value v = EvalSource("=UNICODE(123)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 49.0);
}

TEST(BuiltinsText2Unicode, BoolArgCoercedToText) {
  // TRUE -> text "TRUE" -> first char 'T' -> 84.
  const Value v = EvalSource("=UNICODE(TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 84.0);
}

TEST(BuiltinsText2Unicode, RoundTripWithUnichar) {
  const Value v = EvalSource("=UNICODE(UNICHAR(42))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(BuiltinsText2Unicode, ErrorPropagates) {
  const Value v = EvalSource("=UNICODE(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// CLEAN
// ---------------------------------------------------------------------------

TEST(BuiltinsText2Clean, TabRemoved) {
  const Value v = EvalSource("=CLEAN(\"hello\tworld\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "helloworld");
}

TEST(BuiltinsText2Clean, NulRemoved) {
  // Embedded NUL is a real input byte (the parser allows it inside literals
  // here via UNICHAR(1)+UNICHAR(0)+...); we use CONCAT for safety.
  const Value v = EvalSource("=CLEAN(CONCAT(\"a\", UNICHAR(1), \"b\"))");
  ASSERT_TRUE(v.is_text());
  // U+0001 is < 0x20 so it is stripped.
  EXPECT_EQ(v.as_text(), "ab");
}

TEST(BuiltinsText2Clean, BellRemoved) {
  // U+0007 BEL via UNICHAR.
  const Value v = EvalSource("=CLEAN(CONCAT(\"a\", UNICHAR(7), \"b\"))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ab");
}

TEST(BuiltinsText2Clean, NewlineAndCRRemoved) {
  // 0x0A LF and 0x0D CR are both controls.
  const Value v = EvalSource("=CLEAN(CONCAT(\"a\", UNICHAR(10), \"b\", UNICHAR(13), \"c\"))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(BuiltinsText2Clean, AllControlBytesStripped) {
  // 0x1F is the highest control byte that should be stripped.
  const Value v = EvalSource("=CLEAN(CONCAT(\"x\", UNICHAR(31), \"y\"))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "xy");
}

TEST(BuiltinsText2Clean, DelPreserved) {
  // 0x7F DEL is NOT in the 0x00-0x1F range and is preserved.
  const Value v = EvalSource("=CLEAN(CONCAT(UNICHAR(127), \"hello\"))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\x7Fhello");
}

TEST(BuiltinsText2Clean, SpacePreserved) {
  // 0x20 SPACE is the boundary; preserved.
  const Value v = EvalSource("=CLEAN(\"a b c\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a b c");
}

TEST(BuiltinsText2Clean, MultiByteUtf8Preserved) {
  // Japanese bytes are all >= 0x80 so they survive cleaning.
  // Input: "あ\x07い" -> "あい".
  const Value v = EvalSource("=CLEAN(CONCAT(\"\xE3\x81\x82\", UNICHAR(7), \"\xE3\x81\x84\"))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82\xE3\x81\x84");
}

TEST(BuiltinsText2Clean, EmptyString) {
  const Value v = EvalSource("=CLEAN(\"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsText2Clean, NoControlsUnchanged) {
  const Value v = EvalSource("=CLEAN(\"hello world\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello world");
}

TEST(BuiltinsText2Clean, NumericArgCoerces) {
  // Number 123 renders as "123", which contains no controls.
  const Value v = EvalSource("=CLEAN(123)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "123");
}

TEST(BuiltinsText2Clean, ErrorPropagates) {
  const Value v = EvalSource("=CLEAN(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// PROPER
// ---------------------------------------------------------------------------

TEST(BuiltinsText2Proper, LowercaseToTitle) {
  const Value v = EvalSource("=PROPER(\"hello world\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Hello World");
}

TEST(BuiltinsText2Proper, UppercaseToTitle) {
  const Value v = EvalSource("=PROPER(\"HELLO WORLD\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Hello World");
}

TEST(BuiltinsText2Proper, MixedCase) {
  const Value v = EvalSource("=PROPER(\"hELLo wOrLD\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Hello World");
}

TEST(BuiltinsText2Proper, HyphenIsWordBoundary) {
  const Value v = EvalSource("=PROPER(\"hello-world\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Hello-World");
}

TEST(BuiltinsText2Proper, DigitsBreakAndCapNextLetter) {
  // "123abc def456ghi" -> "123Abc Def456Ghi". A digit is a non-letter, so
  // the letter that follows starts a new word.
  const Value v = EvalSource("=PROPER(\"123abc def456ghi\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "123Abc Def456Ghi");
}

TEST(BuiltinsText2Proper, EmptyString) {
  const Value v = EvalSource("=PROPER(\"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsText2Proper, SingleLetter) {
  const Value v = EvalSource("=PROPER(\"a\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "A");
}

TEST(BuiltinsText2Proper, AllPunctuation) {
  // No letters - input passes through untouched.
  const Value v = EvalSource("=PROPER(\"!@#$ %^&*\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "!@#$ %^&*");
}

TEST(BuiltinsText2Proper, AsciiAfterJapaneseStartsWord) {
  // "あbc" - the 'b' follows non-ASCII-letter bytes (UTF-8 continuation),
  // so it is treated as the start of a word.
  const Value v = EvalSource("=PROPER(\"\xE3\x81\x82\x62\x63\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82" "Bc");
}

TEST(BuiltinsText2Proper, MultiByteOnlyPassesThrough) {
  // All-Japanese input: every byte is non-letter, so nothing changes.
  const Value v = EvalSource("=PROPER(\"\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86");
}

TEST(BuiltinsText2Proper, ApostropheBoundary) {
  // Apostrophe is a non-letter, so the next letter is title-cased. Excel
  // treats "don't" -> "Don'T". This mirrors the documented behavior.
  const Value v = EvalSource("=PROPER(\"don't\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Don'T");
}

TEST(BuiltinsText2Proper, NumericArgCoerces) {
  // 12 renders as "12" - no letters, passes through.
  const Value v = EvalSource("=PROPER(12)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "12");
}

TEST(BuiltinsText2Proper, BoolArgCoerces) {
  // TRUE -> text "TRUE" -> "True".
  const Value v = EvalSource("=PROPER(TRUE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "True");
}

TEST(BuiltinsText2Proper, ErrorPropagates) {
  const Value v = EvalSource("=PROPER(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
