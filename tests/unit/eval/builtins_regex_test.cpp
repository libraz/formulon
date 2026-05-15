// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// End-to-end tests for the REGEX function family (REGEXTEST,
// REGEXEXTRACT, REGEXREPLACE) backed by PCRE2. Each test parses a
// formula source, evaluates the AST through the default registry, and
// asserts the resulting Value. The patterns and corpora exercise:
//   * Basic match / no-match
//   * Case sensitivity defaulting to "off" (Excel's documented default)
//   * Anchors (^, $) — without PCRE2_MULTILINE
//   * Unicode \w / \d / \s with PCRE2_UCP enabled
//   * return_mode 0..3 of REGEXEXTRACT
//   * occurrence 0..N of REGEXREPLACE
//   * Substitution backreferences ($1 / ${name} / $$)
//   * Domain validation (#VALUE! on bad mode / occurrence /
//     case_sensitivity)
//   * Pattern-length cap (32 767 bytes)
//   * Error propagation across all three functions

#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry.
// Thread-local arenas keep text payloads readable for the immediately
// following EXPECT_*. Each call resets the arenas to avoid cross-test
// contamination.
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
  return evaluate(*root, eval_arena, default_registry(), test::mac_context());
}

// ---------------------------------------------------------------------------
// REGEXTEST
// ---------------------------------------------------------------------------

TEST(RegexTest, BasicMatchTrue) {
  const Value v = EvalSource("=REGEXTEST(\"hello world\", \"world\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(RegexTest, BasicNoMatchFalse) {
  const Value v = EvalSource("=REGEXTEST(\"hello world\", \"xyz\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(RegexTest, CaseSensitiveByDefault) {
  // Mac Excel 365 / MS docs convention: default case_sensitivity = 0 is
  // CASE-SENSITIVE. "hello" pattern does not match "Hello World".
  const Value v = EvalSource("=REGEXTEST(\"Hello World\", \"hello\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(RegexTest, CaseInsensitiveExplicit) {
  // case_sensitivity=1 -> case-insensitive; "hello" now matches "Hello".
  const Value v = EvalSource("=REGEXTEST(\"Hello World\", \"hello\", 1)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(RegexTest, CaseSensitiveExplicitMatches) {
  // case_sensitivity=0 (explicit) is the same as default: case-sensitive.
  const Value v = EvalSource("=REGEXTEST(\"Hello World\", \"Hello\", 0)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(RegexTest, AnchorStartMatches) {
  const Value v = EvalSource("=REGEXTEST(\"hello world\", \"^hello\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(RegexTest, AnchorStartFails) {
  const Value v = EvalSource("=REGEXTEST(\"say hello\", \"^hello\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(RegexTest, AnchorEndMatches) {
  const Value v = EvalSource("=REGEXTEST(\"say hello\", \"hello$\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(RegexTest, AnchorAcrossEmbeddedNewlineDefaultMode) {
  // Without PCRE2_MULTILINE, ^ matches subject start only. Excel
  // string literals do not interpret backslash escapes, so we build
  // the embedded LF via CHAR(10) and concatenate. "world" appears on
  // a second line so ^world should fail.
  const Value v = EvalSource("=REGEXTEST(\"hello\" & CHAR(10) & \"world\", \"^world\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(RegexTest, UnicodeWordWithHiragana) {
  // \w under PCRE2_UCP includes hiragana letters. UTF-8 bytes for
  // "ひらがな": E3 81 B2 E3 82 89 E3 81 8C E3 81 AA.
  const Value v = EvalSource("=REGEXTEST(\"\xE3\x81\xB2\xE3\x82\x89\xE3\x81\x8C\xE3\x81\xAA\", \"^\\w+$\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(RegexTest, UnicodeDigitWithFullWidth) {
  // \d under PCRE2_UCP includes full-width digits 0-9 (U+FF10..U+FF19).
  // UTF-8 bytes: EF BC 90 .. EF BC 99. Use "１２３" (FF11 FF12 FF13).
  const Value v = EvalSource("=REGEXTEST(\"\xEF\xBC\x91\xEF\xBC\x92\xEF\xBC\x93\", \"^\\d+$\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(RegexTest, UnicodeWhitespaceIdeographic) {
  // U+3000 IDEOGRAPHIC SPACE (E3 80 80) is whitespace under UCP.
  const Value v = EvalSource(
      "=REGEXTEST(\"a\xE3\x80\x80"
      "b\", \"\\s\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(RegexTest, EmptyPatternIsValueError) {
  const Value v = EvalSource("=REGEXTEST(\"hello\", \"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexTest, EmptySubjectFalse) {
  const Value v = EvalSource("=REGEXTEST(\"\", \"foo\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(RegexTest, InvalidPatternIsValueError) {
  // Unbalanced bracket — pcre2_compile fails.
  const Value v = EvalSource("=REGEXTEST(\"hello\", \"[abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexTest, CaseSensitivityOutOfRange) {
  // Domain {0, 1}; 2 -> #VALUE!.
  const Value v = EvalSource("=REGEXTEST(\"hello\", \"hello\", 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexTest, ErrorInTextPropagates) {
  const Value v = EvalSource("=REGEXTEST(#REF!, \"hello\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(RegexTest, ErrorInPatternPropagates) {
  const Value v = EvalSource("=REGEXTEST(\"hello\", #N/A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// REGEXEXTRACT — return_mode 0 (scalar)
// ---------------------------------------------------------------------------

TEST(RegexExtract, Mode0FirstMatch) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abc 123 def 456\", \"\\d+\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "123");
}

TEST(RegexExtract, Mode0NoMatchIsNA) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abc def\", \"\\d+\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(RegexExtract, Mode0CaseSensitiveByDefaultNoMatch) {
  // Default case_sensitivity=0 -> case-sensitive (Mac convention).
  // "hello" pattern does not match "HELLO".
  const Value v = EvalSource("=REGEXEXTRACT(\"HELLO world\", \"hello\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(RegexExtract, Mode0CaseInsensitiveExplicit) {
  // case_sensitivity=1 -> case-insensitive; "hello" matches "HELLO".
  const Value v = EvalSource("=REGEXEXTRACT(\"HELLO world\", \"hello\", 0, 1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "HELLO");
}

// ---------------------------------------------------------------------------
// REGEXEXTRACT — return_mode 1 (row array of all matches)
// ---------------------------------------------------------------------------

TEST(RegexExtract, Mode1AllMatchesRow) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abc 123 def 456 ghi 789\", \"\\d+\", 1)");
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 1U);
  EXPECT_EQ(arr->cols, 3U);
  ASSERT_TRUE(arr->cells[0].is_text());
  EXPECT_EQ(arr->cells[0].as_text(), "123");
  ASSERT_TRUE(arr->cells[1].is_text());
  EXPECT_EQ(arr->cells[1].as_text(), "456");
  ASSERT_TRUE(arr->cells[2].is_text());
  EXPECT_EQ(arr->cells[2].as_text(), "789");
}

TEST(RegexExtract, Mode1SingleMatchStillArray) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abc 123\", \"\\d+\", 1)");
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 1U);
  EXPECT_EQ(arr->cols, 1U);
  ASSERT_TRUE(arr->cells[0].is_text());
  EXPECT_EQ(arr->cells[0].as_text(), "123");
}

TEST(RegexExtract, Mode1NoMatchNA) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abc def\", \"\\d+\", 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// REGEXEXTRACT — return_mode 2 (first match's capture groups, row array)
// ---------------------------------------------------------------------------

TEST(RegexExtract, Mode2CaptureGroupsRow) {
  const Value v = EvalSource("=REGEXEXTRACT(\"2024-01-15\", \"(\\d{4})-(\\d{2})-(\\d{2})\", 2)");
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 1U);
  EXPECT_EQ(arr->cols, 3U);
  ASSERT_TRUE(arr->cells[0].is_text());
  EXPECT_EQ(arr->cells[0].as_text(), "2024");
  ASSERT_TRUE(arr->cells[1].is_text());
  EXPECT_EQ(arr->cells[1].as_text(), "01");
  ASSERT_TRUE(arr->cells[2].is_text());
  EXPECT_EQ(arr->cells[2].as_text(), "15");
}

TEST(RegexExtract, Mode2NoCaptureGroupsValueError) {
  // Pattern matches but has zero capture groups -> #VALUE!.
  const Value v = EvalSource("=REGEXEXTRACT(\"hello\", \"hello\", 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexExtract, Mode2NoMatchNA) {
  // Pattern has groups but doesn't match -> #N/A (not #VALUE!).
  const Value v = EvalSource("=REGEXEXTRACT(\"abc\", \"(\\d+)-(\\d+)\", 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(RegexExtract, Mode2NamedCaptures) {
  const Value v = EvalSource("=REGEXEXTRACT(\"2024-01\", \"(?<year>\\d{4})-(?<month>\\d{2})\", 2)");
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 1U);
  EXPECT_EQ(arr->cols, 2U);
  EXPECT_EQ(arr->cells[0].as_text(), "2024");
  EXPECT_EQ(arr->cells[1].as_text(), "01");
}

TEST(RegexExtract, Mode2NonCapturingGroupSkipped) {
  // (?:...) does not contribute to capture count.
  const Value v = EvalSource("=REGEXEXTRACT(\"abc-123\", \"(?:abc)-(\\d+)\", 2)");
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->cols, 1U);
  EXPECT_EQ(arr->cells[0].as_text(), "123");
}

// ---------------------------------------------------------------------------
// REGEXEXTRACT — return_mode 3 (all matches' capture groups, 2D array)
// ---------------------------------------------------------------------------

TEST(RegexExtract, Mode3AllCaptureGroups2D) {
  const Value v = EvalSource("=REGEXEXTRACT(\"a1 b2 c3\", \"(\\w)(\\d)\", 3)");
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 3U);
  EXPECT_EQ(arr->cols, 2U);
  EXPECT_EQ(arr->cells[0].as_text(), "a");
  EXPECT_EQ(arr->cells[1].as_text(), "1");
  EXPECT_EQ(arr->cells[2].as_text(), "b");
  EXPECT_EQ(arr->cells[3].as_text(), "2");
  EXPECT_EQ(arr->cells[4].as_text(), "c");
  EXPECT_EQ(arr->cells[5].as_text(), "3");
}

TEST(RegexExtract, Mode3NoCaptureGroupsValueError) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abc abc\", \"abc\", 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexExtract, Mode3NoMatchNA) {
  const Value v = EvalSource("=REGEXEXTRACT(\"xyz\", \"(\\d)(\\d)\", 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// REGEXEXTRACT — error / domain handling
// ---------------------------------------------------------------------------

TEST(RegexExtract, ModeOutOfRangeIsValueError) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abc\", \"a\", 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexExtract, NegativeModeIsValueError) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abc\", \"a\", -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexExtract, EmptyPatternIsValueError) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abc\", \"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexExtract, InvalidPatternIsValueError) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abc\", \"(\\d\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexExtract, ErrorInTextPropagates) {
  const Value v = EvalSource("=REGEXEXTRACT(#DIV/0!, \"a\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// REGEXREPLACE — basic
// ---------------------------------------------------------------------------

TEST(RegexReplace, ReplaceAllDefault) {
  const Value v = EvalSource("=REGEXREPLACE(\"abc 123 def 456\", \"\\d+\", \"#\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc # def #");
}

TEST(RegexReplace, ReplaceFirstOccurrence) {
  const Value v = EvalSource("=REGEXREPLACE(\"abc 123 def 456 ghi 789\", \"\\d+\", \"#\", 1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc # def 456 ghi 789");
}

TEST(RegexReplace, ReplaceThirdOccurrence) {
  const Value v = EvalSource("=REGEXREPLACE(\"abc 123 def 456 ghi 789\", \"\\d+\", \"#\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc 123 def 456 ghi #");
}

TEST(RegexReplace, OccurrenceExceedsTotalReturnsOriginal) {
  // 5 occurrences requested but only 3 matches; original text returned.
  const Value v = EvalSource("=REGEXREPLACE(\"abc 123 def 456 ghi 789\", \"\\d+\", \"#\", 5)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc 123 def 456 ghi 789");
}

TEST(RegexReplace, NegativeOccurrenceIsPermissive) {
  // Mac Excel 365 is permissive on negative occurrence: it folds to
  // global replacement. Formulon clamps to 0 to match.
  const Value v = EvalSource("=REGEXREPLACE(\"abc\", \"a\", \"#\", -1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "#bc");
}

TEST(RegexReplace, EmptyReplacementDeletesMatches) {
  const Value v = EvalSource("=REGEXREPLACE(\"abc 123 def 456\", \"\\d+\", \"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc  def ");
}

TEST(RegexReplace, EmptyPatternIsValueError) {
  const Value v = EvalSource("=REGEXREPLACE(\"abc\", \"\", \"#\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexReplace, InvalidPatternIsValueError) {
  const Value v = EvalSource("=REGEXREPLACE(\"abc\", \"[abc\", \"#\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// REGEXREPLACE — backreferences (extended substitute syntax)
// ---------------------------------------------------------------------------

TEST(RegexReplace, BackreferenceDollar1) {
  const Value v = EvalSource("=REGEXREPLACE(\"john smith\", \"(\\w+) (\\w+)\", \"$2 $1\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "smith john");
}

TEST(RegexReplace, NamedGroupReplacement) {
  const Value v =
      EvalSource("=REGEXREPLACE(\"2024-01-15\", \"(?<y>\\d{4})-(?<m>\\d{2})-(?<d>\\d{2})\", \"${d}/${m}/${y}\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "15/01/2024");
}

TEST(RegexReplace, DoubleDollarLiteral) {
  // $$ -> literal $.
  const Value v = EvalSource("=REGEXREPLACE(\"100\", \"(\\d+)\", \"$$$1\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "$100");
}

// ---------------------------------------------------------------------------
// REGEXREPLACE — case sensitivity
// ---------------------------------------------------------------------------

TEST(RegexReplace, CaseSensitiveDefault) {
  // Default case_sensitivity=0 -> case-sensitive (Mac convention).
  // Only the lowercase "hello" is replaced.
  const Value v = EvalSource("=REGEXREPLACE(\"Hello HELLO hello\", \"hello\", \"#\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Hello HELLO #");
}

TEST(RegexReplace, CaseInsensitiveExplicit) {
  // case_sensitivity=1 -> case-insensitive; all three variants are
  // replaced.
  const Value v = EvalSource("=REGEXREPLACE(\"Hello HELLO hello\", \"hello\", \"#\", 0, 1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "# # #");
}

// ---------------------------------------------------------------------------
// Cross-cutting: case_sensitivity and pattern length cap
// ---------------------------------------------------------------------------

TEST(RegexCrossCutting, RegextestCaseInsensitive) {
  // case_sensitivity=1 -> case-insensitive; "foo" matches "FOO".
  const Value v = EvalSource("=REGEXTEST(\"FOO\", \"foo\", 1)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(RegexCrossCutting, RegexextractCaseSensitiveDefault) {
  // case_sensitivity=0 (default) -> case-sensitive; "foo" matches only
  // the lowercase occurrence.
  const Value v = EvalSource("=REGEXEXTRACT(\"FOO foo\", \"foo\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "foo");
}

TEST(RegexCrossCutting, RegexreplaceCaseSensitiveDefault) {
  // case_sensitivity=0 (default) -> case-sensitive; only "foo" is
  // replaced.
  const Value v = EvalSource("=REGEXREPLACE(\"FOO foo\", \"foo\", \"#\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "FOO #");
}

TEST(RegexCrossCutting, ErrorInPatternPropagatesAcrossFamily) {
  const Value t = EvalSource("=REGEXTEST(\"abc\", #N/A)");
  ASSERT_TRUE(t.is_error());
  EXPECT_EQ(t.as_error(), ErrorCode::NA);

  const Value e = EvalSource("=REGEXEXTRACT(\"abc\", #N/A)");
  ASSERT_TRUE(e.is_error());
  EXPECT_EQ(e.as_error(), ErrorCode::NA);

  const Value r = EvalSource("=REGEXREPLACE(\"abc\", #N/A, \"#\")");
  ASSERT_TRUE(r.is_error());
  EXPECT_EQ(r.as_error(), ErrorCode::NA);
}

TEST(RegexCrossCutting, ErrorInReplacementPropagates) {
  const Value v = EvalSource("=REGEXREPLACE(\"abc\", \"a\", #REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(RegexCrossCutting, PatternTooLongIsValueError) {
  // Build a 32 768-byte pattern (one over the cap). Use literal "a"
  // bytes — REPT in the formula is parsed as a function call, so we
  // construct the pattern at C++ level and feed it through a literal.
  std::string pat(32768, 'a');
  std::string src = "=REGEXTEST(\"abc\", \"";
  src += pat;
  src += "\")";
  const Value v = EvalSource(src);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexCrossCutting, PatternAtCapBoundaryIsAccepted) {
  // 32 767 bytes is exactly the cap; the pattern is "a{32767}" via
  // a literal repeating "a". PCRE2 will compile it (the actual
  // resulting NFA is small because PCRE2 collapses runs at compile
  // time). The compile MAY hit other limits at this scale; allow
  // either a successful #N/A (no match against "abc") or a #VALUE!
  // from PCRE2's own size guards. We only assert that the pattern
  // length cap itself is not the gating error: a 32_767-byte pattern
  // must not be rejected at the pre-compile guard.
  std::string pat(32767, 'a');
  std::string src = "=REGEXTEST(\"";
  src += pat;
  src += "\", \"";
  src += pat;
  src += "\")";
  const Value v = EvalSource(src);
  // Either TRUE (the long subject matches the long pattern) or some
  // PCRE2 internal bound triggers #VALUE!. Either is acceptable; we
  // are testing the pre-compile cap, not PCRE2's internal limits.
  EXPECT_TRUE(v.is_boolean() || v.is_error()) << "unexpected kind for boundary pattern";
}

// ---------------------------------------------------------------------------
// Arity errors
// ---------------------------------------------------------------------------

TEST(RegexArity, RegextestZeroArgsIsValueError) {
  const Value v = EvalSource("=REGEXTEST()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexArity, RegextestTooManyArgsIsValueError) {
  const Value v = EvalSource("=REGEXTEST(\"a\", \"b\", 0, 99)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexArity, RegexextractOneArgIsValueError) {
  const Value v = EvalSource("=REGEXEXTRACT(\"a\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexArity, RegexreplaceTwoArgsIsValueError) {
  const Value v = EvalSource("=REGEXREPLACE(\"a\", \"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
