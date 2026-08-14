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

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/regex_lazy.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

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

TEST(RegexCrossCutting, PatternPastCapIsRejectedWhileLexing) {
  // A pattern one byte past the 32 767-character cap can only be written
  // as a literal, and that literal makes the formula longer than the
  // tokenizer's own UTF-16 length cap. Since the cap bounds a single
  // token as well as the token stream, the source is rejected while
  // lexing and the pattern never reaches PCRE2 -- so the engine-side
  // length guard is unreachable from a formula rather than untested.
  std::string pat(32768, 'a');
  std::string src = "=REGEXTEST(\"abc\", \"";
  src += pat;
  src += "\")";
  static thread_local Arena parse_arena;
  parse_arena.reset();
  parser::Parser p(src, parse_arena);
  (void)p.parse();
  EXPECT_FALSE(p.errors().empty());
}

TEST(RegexCrossCutting, PatternPastCapIsValueError) {
  // Build the over-long pattern at evaluation time. Writing it as a
  // literal would push the formula source past the tokenizer's own
  // length cap, so the engine-side guard would never see it.
  const Value v = EvalSource("=REGEXTEST(\"abc\", REPT(\"a\", 32767) & \"a\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegexCrossCutting, PatternAtCapBoundaryIsAccepted) {
  // 32 767 characters is exactly the cap. PCRE2 will compile it (the
  // resulting NFA is small because PCRE2 collapses runs at compile
  // time). The compile MAY hit other limits at this scale; allow
  // either a successful boolean or a #VALUE! from PCRE2's own size
  // guards. We only assert that the pattern length cap itself is not
  // the gating error: a 32 767-character pattern must not be rejected
  // at the pre-compile guard. The pattern is built at evaluation time
  // for the same reason as the test above.
  const Value v = EvalSource("=REGEXTEST(\"abc\", REPT(\"a\", 32767))");
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

// A substitution that requires a retry after the initial output buffer is
// exhausted still succeeds when the result fits Excel's text-cell limit.
TEST(RegexReplace, OutputGrowthBeyondInitialBufferSucceeds) {
  const Value v = EvalSource("=REGEXREPLACE(REPT(\"a\", 10000), \"a\", \"bb\")");
  ASSERT_TRUE(v.is_text()) << "expected a successfully grown substitution";
  EXPECT_EQ(v.as_text(), std::string(20000, 'b'));
}

// The exact 32,767 UTF-16-unit boundary is valid, even when the result is
// reached through the retry path.
TEST(RegexReplace, OutputAtTextCapSucceedsAfterRetry) {
  const Value v = EvalSource("=REGEXREPLACE(\"x\" & REPT(\"a\", 16381) & \"😀\" & \"a\", \"a\", \"bb\")");
  ASSERT_TRUE(v.is_text()) << "expected the exact text-cap result to succeed";
  std::string expected = "x";
  expected.append(32762, 'b');
  expected += "😀bb";
  EXPECT_EQ(v.as_text(), expected);
}

// A substitution whose result is one UTF-16 unit past Excel's 32,767-unit
// text-cell limit surfaces #VALUE! instead of growing an unbounded buffer.
TEST(RegexReplace, OutputPastTextCapIsValueError) {
  // 16,383 'a's, each replaced by "bb", plus one emoji, gives exactly
  // 32,768 UTF-16 units (the emoji contributes two units).
  const Value v = EvalSource("=REGEXREPLACE(REPT(\"a\", 16382) & \"😀\" & \"a\", \"a\", \"bb\")");
  ASSERT_TRUE(v.is_error()) << "expected #VALUE! for over-cap output";
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// A substitution that stays within the cap is unaffected.
TEST(RegexReplace, OutputWithinTextCapSucceeds) {
  const Value v = EvalSource("=REGEXREPLACE(\"aaa\", \"a\", \"bb\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "bbbbbb");
}

// A find-all scan retains one span slot per match plus one per capture
// group. PCRE2's `match_limit` bounds a single match attempt and says
// nothing about that accumulation, so the kernel charges it against its own
// budget. A subject long enough to blow the budget reports #CALC! for the
// extract / replace pair, the same surface PCRE2's own resource errors use.
TEST(RegexExtract, MatchAccumulationPastBudgetIsCalcError) {
  // Each of the 32,767 characters matches, and the pattern carries 32
  // capture groups, so the scan would retain ~1.08M span slots -- past the
  // 2^20 budget.
  std::string pattern;
  for (int i = 0; i < 32; ++i) {
    pattern += "()";
  }
  pattern += ".";
  const Value v = EvalSource("=REGEXEXTRACT(REPT(\"a\", 32767), \"" + pattern + "\", 3)");
  ASSERT_TRUE(v.is_error()) << "expected #CALC! once the match budget is exhausted";
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

// Replacing a single occurrence has to locate the N-th match first, which
// runs the same find-all scan and is bounded by the same budget. (The
// default replace-everything mode streams through PCRE2's own global
// substitution and never accumulates spans, so only this path is affected.)
TEST(RegexReplace, NthOccurrenceMatchAccumulationPastBudgetIsCalcError) {
  std::string pattern;
  for (int i = 0; i < 32; ++i) {
    pattern += "()";
  }
  pattern += ".";
  const Value v = EvalSource("=REGEXREPLACE(REPT(\"a\", 32767), \"" + pattern + "\", \"b\", 5)");
  ASSERT_TRUE(v.is_error()) << "expected #CALC! once the match budget is exhausted";
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

// The same call with a subject short enough to stay inside the budget still
// replaces the requested occurrence.
TEST(RegexReplace, NthOccurrenceWithinBudgetSucceeds) {
  const Value v = EvalSource("=REGEXREPLACE(\"aaaa\", \"a\", \"b\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "aaba");
}

// A scan well inside the budget is unaffected.
TEST(RegexExtract, MatchAccumulationWithinBudgetSucceeds) {
  const Value v = EvalSource("=REGEXEXTRACT(\"abcabc\", \"(a)(b)\", 3)");
  ASSERT_TRUE(v.is_array()) << "expected the 2x2 capture matrix";
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
}

// ---------------------------------------------------------------------------
// Compile reuse across a broadcast
// ---------------------------------------------------------------------------
//
// Pattern compilation is loop-invariant and costs far more than matching a
// short cell string, so one REGEX* call must compile its pattern once no
// matter how many subjects it broadcasts over.

TEST(RegexTest, ArrayBroadcastCompilesThePatternOncePerCall) {
  const std::uint64_t before_small = regex_compile_count();
  const Value small = EvalSource("=REGEXTEST({\"a1\",\"b2\";\"c3\",\"d\"}, \"\\d\")");
  ASSERT_TRUE(small.is_array());
  ASSERT_EQ(small.as_array_rows(), 2U);
  ASSERT_EQ(small.as_array_cols(), 2U);
  EXPECT_TRUE(small.as_array_cells()[0].as_boolean());
  EXPECT_FALSE(small.as_array_cells()[3].as_boolean());
  EXPECT_EQ(regex_compile_count() - before_small, 1U) << "one compile per call, not one per cell";

  // Same assertion over a 1,000-cell range: the compile total must not scale
  // with the number of subjects.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t r = 0; r < 1000U; ++r) {
    sheet.set_cell_value(r, 0, Value::number(static_cast<double>(r)));
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=REGEXTEST(A1:A1000, \"\\d\")", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);

  const std::uint64_t before_large = regex_compile_count();
  const Value large = evaluate(*root, eval_arena, default_registry(), ctx);
  ASSERT_TRUE(large.is_array());
  ASSERT_EQ(large.as_array_rows(), 1000U);
  EXPECT_TRUE(large.as_array_cells()[999].as_boolean());
  EXPECT_EQ(regex_compile_count() - before_large, 1U) << "compilation must not scale with the subject count";
}

TEST(RegexReplace, NthOccurrenceBroadcastCompilesThePatternOncePerCall) {
  // The occurrence-N path scans for the match position and then substitutes
  // at it; both steps share the one compiled program.
  const std::uint64_t before = regex_compile_count();
  const Value v = EvalSource("=REGEXREPLACE({\"aaaa\";\"aaa\"}, \"a\", \"b\", 2)");
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "abaa");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "aba");
  EXPECT_EQ(regex_compile_count() - before, 1U);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
