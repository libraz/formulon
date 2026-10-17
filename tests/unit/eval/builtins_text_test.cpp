//
// End-to-end tests for the text built-in functions: UPPER, LOWER, TRIM,
// LEFT, RIGHT, MID, REPT, SUBSTITUTE, FIND, SEARCH, EXACT. Each test
// parses a formula source, evaluates the AST through the default registry,
// and asserts the resulting Value. TEXT / VALUE / NUMBERVALUE live in
// `builtins_value_numbervalue_test.cpp` next to their shared format engine.

#include <array>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;

struct TextProfileCase {
  ExcelProfile profile;
  const char* id;
};

constexpr std::array<TextProfileCase, 4> kTextProfiles = {{
    {mac_365_ja_jp_profile(), "mac-365-ja_JP"},
    {win_365_ja_jp_profile(), "win-365-ja_JP"},
    {mac_365_en_us_profile(), "mac-365-en_US"},
    {win_365_en_us_profile(), "win-365-en_US"},
}};

Value EvalWithProfile(std::string_view src, ExcelProfile profile) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(profile);
  return EvalSourceIn(src, wb, wb.sheet(0));
}

// ---------------------------------------------------------------------------
// UPPER
// ---------------------------------------------------------------------------

TEST(TextUpper, AsciiBasic) {
  const Value v = EvalSource("=UPPER(\"hello\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "HELLO");
}

TEST(TextUpper, MixedCase) {
  const Value v = EvalSource("=UPPER(\"Hello World\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "HELLO WORLD");
}

TEST(TextUpper, EmptyString) {
  const Value v = EvalSource("=UPPER(\"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextUpper, MultiBytePassesThrough) {
  // Characters without a case mapping pass through: "caf\xE3\x81\x82" -> "CAF\xE3\x81\x82".
  const Value v = EvalSource("=UPPER(\"caf\xE3\x81\x82\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "CAF\xE3\x81\x82");
}

TEST(TextUpper, AccentedLetterIsMapped) {
  const Value v = EvalSource("=UPPER(\"caf\xC3\xA9\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "CAF\xC3\x89");
}

TEST(TextUpper, ErrorPropagates) {
  const Value v = EvalSource("=UPPER(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// LOWER
// ---------------------------------------------------------------------------

TEST(TextLower, AsciiBasic) {
  const Value v = EvalSource("=LOWER(\"HELLO\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello");
}

TEST(TextLower, MixedCase) {
  const Value v = EvalSource("=LOWER(\"Hello World\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello world");
}

TEST(TextLower, MultiBytePassesThrough) {
  const Value v = EvalSource("=LOWER(\"CAF\xC3\x89\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "caf\xC3\xA9");
}

// ---------------------------------------------------------------------------
// TRIM
// ---------------------------------------------------------------------------

TEST(TextTrim, LeadingSpaces) {
  const Value v = EvalSource("=TRIM(\"   abc\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(TextTrim, TrailingSpaces) {
  const Value v = EvalSource("=TRIM(\"abc   \")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(TextTrim, InternalRunsCollapse) {
  const Value v = EvalSource("=TRIM(\"a   b\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a b");
}

TEST(TextTrim, AllSpacesBecomeEmpty) {
  const Value v = EvalSource("=TRIM(\"     \")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextTrim, NoSpacesUnchanged) {
  const Value v = EvalSource("=TRIM(\"abc\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(TextTrim, SingleSpaceUnchanged) {
  const Value v = EvalSource("=TRIM(\"a b\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a b");
}

TEST(TextTrim, MultiBytePreserved) {
  // "  あい  " -> "あい". The multi-byte bytes pass through untouched.
  const Value v = EvalSource("=TRIM(\"  \xE3\x81\x82\xE3\x81\x84  \")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82\xE3\x81\x84");
}

TEST(TextTrim, StripsIdeographicSpace) {
  // Mac Excel 365 ja-JP: TRIM("　a　") -> "a". U+3000 (UTF-8 E3 80 80)
  // is treated as a trimmable whitespace at edges.
  const Value v = EvalWithProfile(
      "=TRIM(\"\xE3\x80\x80"
      "a\xE3\x80\x80\")",
      mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a");
}

TEST(TextTrim, CollapsesIdeographicWithAscii) {
  // A mixed run collapses to its FIRST member, not to an ASCII space. The
  // run here opens with U+3000, so that is what survives. Excel keeps the
  // character rather than normalising the run, which matters because a
  // full-width space between two words is typography, not padding.
  const Value v = EvalWithProfile(
      "=TRIM(\"a\xE3\x80\x80 \xE3\x80\x80"
      "b\")",
      mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a\xE3\x80\x80\x62");
}

TEST(TextTrim, MixedRunOpeningWithAsciiKeepsTheAsciiSpace) {
  // The mirror of the case above: the same two space kinds in the other
  // order, so the ASCII space is the survivor. The pair is what shows the
  // rule is "first of the run" rather than "prefer one kind".
  const Value v = EvalWithProfile(
      "=TRIM(\"a \xE3\x80\x80"
      "b\")",
      mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a b");
}

// The interior rule for U+3000 on its own. A lone ideographic space
// between two non-space characters is a run of one, so it survives as
// itself and the text is unchanged. This is the ja-JP everyday case:
// rewriting it to U+0020 would silently reflow Japanese text that was
// spaced deliberately.
TEST(TextTrim, SingleInteriorIdeographicSpaceIsKept) {
  // "あ　い" is already trimmed.
  const Value v = EvalWithProfile("=TRIM(\"\xE3\x81\x82\xE3\x80\x80\xE3\x81\x84\")", mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82\xE3\x80\x80\xE3\x81\x84");

  const Value len = EvalWithProfile("=LEN(TRIM(\"\xE3\x81\x82\xE3\x80\x80\xE3\x81\x84\"))", mac_365_ja_jp_profile());
  ASSERT_TRUE(len.is_number());
  EXPECT_DOUBLE_EQ(len.as_number(), 3.0);
}

TEST(TextTrim, RepeatedInteriorIdeographicSpacesCollapseToOneIdeographic) {
  // "あ　　い" -> "あ　い": the run shortens to one, and the one it keeps
  // is still U+3000.
  const Value v =
      EvalWithProfile("=TRIM(\"\xE3\x81\x82\xE3\x80\x80\xE3\x80\x80\xE3\x81\x84\")", mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82\xE3\x80\x80\xE3\x81\x84");
}

TEST(TextTrim, EdgeAndInteriorIdeographicSpacesTogether) {
  // "　　あ　　い　　" -> "あ　い": the edge runs go away entirely and the
  // interior one shortens to a single U+3000.
  const Value v = EvalWithProfile(
      "=TRIM(\"\xE3\x80\x80\xE3\x80\x80\xE3\x81\x82\xE3\x80\x80\xE3\x80\x80\xE3\x81\x84\xE3\x80\x80\xE3\x80\x80\")",
      mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82\xE3\x80\x80\xE3\x81\x84");
}

// A byte that merely looks space-adjacent is not one. Tabs and newlines
// survive verbatim, and a trimmable space next to one is still collapsed
// rather than deleted, because the run ends at the tab.
TEST(TextTrim, TabAndNewlineArePreservedVerbatim) {
  const Value tabbed = EvalSource("=TRIM(\"a\tb\")");
  ASSERT_TRUE(tabbed.is_text());
  EXPECT_EQ(tabbed.as_text(), "a\tb");

  const Value newlined = EvalSource("=TRIM(\"a\nb\")");
  ASSERT_TRUE(newlined.is_text());
  EXPECT_EQ(newlined.as_text(), "a\nb");

  const Value leading_tab = EvalSource("=TRIM(\"\ta\")");
  ASSERT_TRUE(leading_tab.is_text());
  EXPECT_EQ(leading_tab.as_text(), "\ta");
}

TEST(TextTrim, IdeographicSpaceAdjacentToTabSurvivesAsItself) {
  // "a　\t　b" is unchanged. The tab is not trimmable and ends the run, so
  // each ideographic space is a run of one and keeps its own character
  // rather than being deleted or rewritten.
  const Value v = EvalWithProfile(
      "=TRIM(\"a\xE3\x80\x80\t\xE3\x80\x80"
      "b\")",
      mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a\xE3\x80\x80\t\xE3\x80\x80\x62");
}

TEST(TextTrim, DoesNotStripNbsp) {
  // NBSP (U+00A0, CHAR(160)) is NOT trimmable in Excel. The output must
  // retain both leading and trailing NBSPs verbatim.
  const Value v = EvalSource("=CHAR(160) & \"x\" & CHAR(160)");
  ASSERT_TRUE(v.is_text());
  const std::string raw{v.as_text()};
  const Value trimmed = EvalSource("=TRIM(CHAR(160) & \"x\" & CHAR(160))");
  ASSERT_TRUE(trimmed.is_text());
  EXPECT_EQ(trimmed.as_text(), raw);
}

TEST(TextProfileWidth, EnglishWidthFunctionsAreNoOps) {
  for (const TextProfileCase& test_case : kTextProfiles) {
    SCOPED_TRACE(test_case.id);
    const bool japanese = test_case.profile.locale == ExcelLocale::kJaJP;

    const Value asc = EvalWithProfile("=ASC(\"ＡＢＣ　１２３\")", test_case.profile);
    ASSERT_TRUE(asc.is_text());
    EXPECT_EQ(asc.as_text(), japanese ? "ABC 123" : "ＡＢＣ　１２３");

    const Value dbcs = EvalWithProfile("=DBCS(\"ABC 123\")", test_case.profile);
    ASSERT_TRUE(dbcs.is_text());
    EXPECT_EQ(dbcs.as_text(), japanese ? "ＡＢＣ　１２３" : "ABC 123");
  }
}

TEST(TextProfileTrim, IdeographicSpaceFollowsDbcsProfile) {
  for (const TextProfileCase& test_case : kTextProfiles) {
    SCOPED_TRACE(test_case.id);
    const Value trimmed = EvalWithProfile("=TRIM(\"　a　\")", test_case.profile);
    ASSERT_TRUE(trimmed.is_text());
    EXPECT_EQ(trimmed.as_text(), test_case.profile.locale == ExcelLocale::kJaJP ? "a" : "　a　");
  }
}

TEST(TextProfileLazyDispatch, TextBackslashRewriteUsesTheSbcCodepage) {
  for (const TextProfileCase& test_case : kTextProfiles) {
    SCOPED_TRACE(test_case.id);
    const Value formatted = EvalWithProfile("=TEXT(10, \"\\$0\")", test_case.profile);
    ASSERT_TRUE(formatted.is_text());
    const bool win_ja = test_case.profile.host == ExcelHost::kWin365 && test_case.profile.locale == ExcelLocale::kJaJP;
    EXPECT_EQ(formatted.as_text(), win_ja ? "\xC2\xA5$10" : "$10");
  }
}

TEST(TextProfileLazyDispatch, TextBackslashRewriteBroadcastsOverFormatArrays) {
  const Value joined = EvalWithProfile("=TEXTJOIN(\"|\",FALSE,IFERROR(TEXT(1234.5,{\"0\",\"0\\:0\",\"0 \\3\"}),\"#\"))",
                                       win_365_ja_jp_profile());
  ASSERT_TRUE(joined.is_text());
  EXPECT_EQ(joined.as_text(),
            "1235|#|1235 \xC2\xA5"
            "3");
  const Value mac = EvalWithProfile("=TEXTJOIN(\"|\",FALSE,TEXT(1234.5,{\"0\",\"0\\:0\"}))", mac_365_ja_jp_profile());
  ASSERT_TRUE(mac.is_text());
  EXPECT_EQ(mac.as_text(), "1235|123:5");
}

// ---------------------------------------------------------------------------
// LEFT
// ---------------------------------------------------------------------------

TEST(TextLeft, DefaultNIsOne) {
  const Value v = EvalSource("=LEFT(\"hello\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "h");
}

TEST(TextLeft, ExplicitN) {
  const Value v = EvalSource("=LEFT(\"hello\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hel");
}

TEST(TextLeft, ZeroIsEmpty) {
  const Value v = EvalSource("=LEFT(\"hello\", 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextLeft, NGreaterThanLengthClamps) {
  const Value v = EvalSource("=LEFT(\"hi\", 10)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hi");
}

TEST(TextLeft, NegativeNIsValueError) {
  const Value v = EvalSource("=LEFT(\"hello\", -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextLeft, HugeNBeyondIntRangeReturnsWholeText) {
  // 1E+15 overflows int32; the same clamp-before-cast must apply on every
  // target, so the result is the whole string regardless of host.
  const Value v = EvalSource("=LEFT(\"text\", 1E+15)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "text");
}

TEST(TextLeft, HugeNegativeNBeyondIntRangeIsValueError) {
  // -1E+15 saturates to INT_MIN, which the negative-n guard still rejects
  // the same way it rejects any negative n.
  const Value v = EvalSource("=LEFT(\"text\", -1E+15)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextLeft, MultiByte) {
  // "あいう" (3 BMP units) -> first 2 units = "あい".
  const Value v = EvalSource("=LEFT(\"\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86\", 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82\xE3\x81\x84");
}

TEST(TextLeft, EmojiTwoUnitsKeepsOneCodepoint) {
  // "🎉🎊" — each emoji is 2 units. LEFT(.., 2) returns the first emoji.
  const Value v = EvalSource("=LEFT(\"\xF0\x9F\x8E\x89\xF0\x9F\x8E\x8A\", 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xF0\x9F\x8E\x89");
}

TEST(TextLeft, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=LEFT()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// RIGHT
// ---------------------------------------------------------------------------

TEST(TextRight, DefaultNIsOne) {
  const Value v = EvalSource("=RIGHT(\"hello\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "o");
}

TEST(TextRight, ExplicitN) {
  const Value v = EvalSource("=RIGHT(\"hello\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "llo");
}

TEST(TextRight, ZeroIsEmpty) {
  const Value v = EvalSource("=RIGHT(\"hello\", 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextRight, NGreaterThanLengthClamps) {
  const Value v = EvalSource("=RIGHT(\"hi\", 10)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hi");
}

TEST(TextRight, NegativeNIsValueError) {
  const Value v = EvalSource("=RIGHT(\"hello\", -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextRight, MultiByte) {
  // "あいう" (3 BMP units) -> last 2 units = "いう".
  const Value v = EvalSource("=RIGHT(\"\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86\", 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x84\xE3\x81\x86");
}

// A supplementary-plane codepoint is 2 UTF-16 units, so asking for the last
// 1 unit lands in the middle of a surrogate pair. The slice widens to the
// whole codepoint the same way LEFT's end bound does; collapsing to the
// empty string instead would silently break `=IF(RIGHT(A1,1)="。",...)`.
TEST(TextRight, SupplementaryPlaneTailReturnsWholeCodepoint) {
  const Value v = EvalSource("=RIGHT(\"abc\xF0\x9F\x99\x82\", 1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xF0\x9F\x99\x82");
}

TEST(TextRight, SupplementaryPlaneOnlyTextIsNeverEmpty) {
  const Value v = EvalSource("=RIGHT(\"\xF0\x9F\x99\x82\", 1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xF0\x9F\x99\x82");
}

// LEFT and RIGHT round a surrogate-splitting bound to the same side of the
// codepoint: whichever end is asked for, the whole codepoint survives.
TEST(TextRight, SurrogateRoundingAgreesWithLeft) {
  const Value left = EvalSource("=LEFT(\"\xF0\x9F\x99\x82\x61\x62\", 1)");
  ASSERT_TRUE(left.is_text());
  EXPECT_EQ(left.as_text(), "\xF0\x9F\x99\x82");

  const Value right = EvalSource("=RIGHT(\"\x61\x62\xF0\x9F\x99\x82\", 1)");
  ASSERT_TRUE(right.is_text());
  EXPECT_EQ(right.as_text(), "\xF0\x9F\x99\x82");
}

// ---------------------------------------------------------------------------
// MID
// ---------------------------------------------------------------------------

TEST(TextMid, OneBasedIndexing) {
  // Pin: MID is 1-based; MID("abcde", 2, 3) returns "bcd", not "abc".
  const Value v = EvalSource("=MID(\"abcde\", 2, 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "bcd");
}

TEST(TextMid, StartAtOne) {
  const Value v = EvalSource("=MID(\"abcde\", 1, 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ab");
}

TEST(TextMid, StartBeyondEndReturnsEmpty) {
  const Value v = EvalSource("=MID(\"abc\", 10, 5)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextMid, StartLessThanOneIsValueError) {
  const Value v = EvalSource("=MID(\"abc\", 0, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextMid, NumCharsNegativeIsValueError) {
  const Value v = EvalSource("=MID(\"abc\", 1, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextMid, NumCharsZeroReturnsEmpty) {
  const Value v = EvalSource("=MID(\"abc\", 1, 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextMid, MultiByteSlice) {
  // "あいうえお" (5 BMP units) -> MID(.., 2, 3) = "いうえ".
  const Value v = EvalSource("=MID(\"\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86\xE3\x81\x88\xE3\x81\x8A\", 2, 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x84\xE3\x81\x86\xE3\x81\x88");
}

TEST(TextMid, HugeNumCharsBeyondIntRangeReturnsRestOfText) {
  // num_chars = 1E+15 overflows int32 in the same way LEFT's n does; the
  // result must be identical across x86-64, arm64, and WASM.
  const Value v = EvalSource("=MID(\"text\", 2, 1E+15)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ext");
}

// ---------------------------------------------------------------------------
// REPT
// ---------------------------------------------------------------------------

TEST(TextRept, BasicRepeat) {
  const Value v = EvalSource("=REPT(\"ab\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ababab");
}

TEST(TextRept, ZeroCountIsEmpty) {
  const Value v = EvalSource("=REPT(\"ab\", 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextRept, NegativeCountIsValueError) {
  const Value v = EvalSource("=REPT(\"ab\", -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextRept, ExceedsExcelCapIsValueError) {
  // 100,000 single-character repeats exceeds Excel's 32,767-unit cap.
  const Value v = EvalSource("=REPT(\"a\", 100000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextRept, EmptyTextIsEmpty) {
  const Value v = EvalSource("=REPT(\"\", 5)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

// ---------------------------------------------------------------------------
// SUBSTITUTE
// ---------------------------------------------------------------------------

TEST(TextSubstitute, AllOccurrences) {
  const Value v = EvalSource("=SUBSTITUTE(\"abcabc\", \"a\", \"X\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "XbcXbc");
}

TEST(TextSubstitute, NthInstance) {
  // Replace only the second occurrence.
  const Value v = EvalSource("=SUBSTITUTE(\"abcabc\", \"a\", \"X\", 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abcXbc");
}

TEST(TextSubstitute, InstanceBeyondOccurrencesReturnsInput) {
  const Value v = EvalSource("=SUBSTITUTE(\"abcabc\", \"a\", \"X\", 5)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abcabc");
}

TEST(TextSubstitute, EmptyOldTextReturnsInput) {
  const Value v = EvalSource("=SUBSTITUTE(\"abc\", \"\", \"X\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(TextSubstitute, EmptyNewTextRemoves) {
  const Value v = EvalSource("=SUBSTITUTE(\"abcabc\", \"a\", \"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "bcbc");
}

TEST(TextSubstitute, InstanceLessThanOneIsValueError) {
  const Value v = EvalSource("=SUBSTITUTE(\"abc\", \"a\", \"X\", 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextSubstitute, MultiCharNeedle) {
  const Value v = EvalSource("=SUBSTITUTE(\"foo bar foo\", \"foo\", \"baz\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "baz bar baz");
}

TEST(TextSubstitute, OutputBeyondExcelTextLimitIsValueError) {
  const Value v = EvalSource("=SUBSTITUTE(REPT(\"a\",32767),\"a\",\"aa\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// FIND
// ---------------------------------------------------------------------------

TEST(TextFind, BasicMatch) {
  const Value v = EvalSource("=FIND(\"b\", \"abc\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(TextFind, CaseSensitive) {
  // "a" not found in "Abc" because FIND is case-sensitive.
  const Value v = EvalSource("=FIND(\"a\", \"Abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextFind, WithStartNum) {
  // First "a" at 1 is skipped because start_num=2; next match at 4.
  const Value v = EvalSource("=FIND(\"a\", \"abca\", 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(TextFind, NotFound) {
  const Value v = EvalSource("=FIND(\"z\", \"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextFind, StartNumLessThanOne) {
  const Value v = EvalSource("=FIND(\"a\", \"abc\", 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextFind, StartNumPastEnd) {
  // LEN("abc")+1 = 4 is the last valid start; 5 is out of range.
  const Value v = EvalSource("=FIND(\"a\", \"abc\", 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextFind, EmptyFindTextReturnsStart) {
  const Value v = EvalSource("=FIND(\"\", \"abc\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(TextFind, EmptyFindTextWithStart) {
  const Value v = EvalSource("=FIND(\"\", \"abc\", 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(TextFind, MultiByteResultIsUtf16Position) {
  // "あいうえお" -> position of "う" is 3 (1-based UTF-16 units).
  const Value v =
      EvalSource("=FIND(\"\xE3\x81\x86\", \"\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86\xE3\x81\x88\xE3\x81\x8A\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// SEARCH
// ---------------------------------------------------------------------------

TEST(TextSearch, CaseInsensitive) {
  const Value v = EvalSource("=SEARCH(\"A\", \"abc\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(TextSearch, MixedCaseMatch) {
  const Value v = EvalSource("=SEARCH(\"WORLD\", \"hello world\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(TextSearch, WithStartNum) {
  const Value v = EvalSource("=SEARCH(\"A\", \"abca\", 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

Value EvalMacEnUs(std::string_view formula) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(formula, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << formula;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena, default_registry(), EvalContext().with_excel_profile(mac_365_en_us_profile()));
}

// Mac Excel 365 folds case, width and kana one code point at a time; nothing
// composes, expands or ignores accents (measured, en-US).
TEST(TextSearch, FoldsBeyondAscii) {
  for (const char* formula : {"=SEARCH(\"ä\",\"Äpfel\")", "=SEARCH(\"σ\",\"ΣΟΦΙΑ\")", "=SEARCH(\"п\",\"ПРИВЕТ\")",
                              "=SEARCH(\"ア\",\"あいう\")", "=SEARCH(\"ア\",\"ｱｲｳ\")", "=SEARCH(\"ａ\",\"ABC\")",
                              "=SEARCH(\"A\",\"ａｂｃ\")"}) {
    const Value v = EvalMacEnUs(formula);
    ASSERT_TRUE(v.is_number()) << formula << " -> " << v.debug_to_string();
    EXPECT_EQ(v.as_number(), 1.0) << formula;
  }
  for (const char* formula : {"=SEARCH(\"e\",\"Été\")", "=SEARCH(\"ガ\",\"ｶﾞ\")", "=SEARCH(\"ss\",\"ß\")",
                              "=SEARCH(\"ab\",\"a-b\")", "=SEARCH(\"i\",\"İ\")"}) {
    const Value v = EvalMacEnUs(formula);
    ASSERT_TRUE(v.is_error()) << formula << " -> " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }
}

// A fold that changes the byte length must not shift the reported position.
TEST(TextSearch, FoldedPositionCountsSourceUnits) {
  const Value v = EvalMacEnUs("=SEARCH(\"c\",\"ＡＢＣ\")");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 3.0);
  const Value from = EvalMacEnUs("=SEARCH(\"ä\",\"ÄxÄ\",2)");
  ASSERT_TRUE(from.is_number()) << from.debug_to_string();
  EXPECT_EQ(from.as_number(), 3.0);
}

TEST(TextSearch, NotFound) {
  const Value v = EvalSource("=SEARCH(\"z\", \"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextSearch, EmptyFindTextReturnsStart) {
  const Value v = EvalSource("=SEARCH(\"\", \"abc\", 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

// SEARCH honours Excel's DOS-style wildcards: `?` matches any single
// unit, `*` matches zero-or-more units, `~?` / `~*` match the literal
// metacharacters.
TEST(TextSearch, WildcardQuestionMark) {
  const Value v = EvalSource("=SEARCH(\"a?c\", \"zabcabc\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(TextSearch, WildcardStarGreedy) {
  const Value v = EvalSource("=SEARCH(\"a*c\", \"zabbbbc\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(TextSearch, WildcardStarZeroWidth) {
  // `a*c` with `*` matching the empty run: "ac" starts at UTF-16 position 1.
  const Value v = EvalSource("=SEARCH(\"a*c\", \"ac\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(TextSearch, WildcardEscapedQuestionLiteral) {
  // `~?` should match a literal `?`, so the pattern only matches at pos 4.
  const Value v = EvalSource("=SEARCH(\"a~?c\", \"abcaa?c\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(TextSearch, WildcardEscapedStarLiteral) {
  const Value v = EvalSource("=SEARCH(\"a~*c\", \"abca*c\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(TextSearch, WildcardNoMatchReturnsValueError) {
  const Value v = EvalSource("=SEARCH(\"a?d\", \"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// FIND is strictly literal — confirm it does NOT gain wildcard semantics.
TEST(TextFind, StarIsLiteralNotWildcard) {
  const Value v = EvalSource("=FIND(\"a*c\", \"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// VALUE tests live in `builtins_value_numbervalue_test.cpp` alongside
// the format-string converters (TEXT / VALUE / NUMBERVALUE).

// ---------------------------------------------------------------------------
// EXACT
// ---------------------------------------------------------------------------

TEST(TextExact, IdenticalIsTrue) {
  const Value v = EvalSource("=EXACT(\"abc\", \"abc\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(TextExact, CaseSensitive) {
  const Value v = EvalSource("=EXACT(\"abc\", \"ABC\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(TextExact, DifferingLengthsAreFalse) {
  const Value v = EvalSource("=EXACT(\"abc\", \"abcd\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(TextExact, BothEmptyAreTrue) {
  const Value v = EvalSource("=EXACT(\"\", \"\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(TextExact, ErrorPropagates) {
  const Value v = EvalSource("=EXACT(#REF!, \"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// ARRAYTOTEXT
// ---------------------------------------------------------------------------

TEST(TextArrayToText, ScalarConciseMirrorsValueToText) {
  const Value v = EvalSource("=ARRAYTOTEXT(123.45)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "123.45");
}

TEST(TextArrayToText, ScalarStrictQuotesText) {
  const Value v = EvalSource("=ARRAYTOTEXT(\"a\"\"b\",1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "{\"a\"\"b\"}");
}

TEST(TextArrayToText, RowVectorConciseKeepsAllElements) {
  const Value v = EvalSource("=ARRAYTOTEXT({1,2,3})");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1, 2, 3");
}

TEST(TextArrayToText, ColumnVectorConciseKeepsAllElements) {
  const Value v = EvalSource("=ARRAYTOTEXT({1;2;3})");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1, 2, 3");
}

TEST(TextArrayToText, MatrixStrictUsesArrayLiteralDelimiters) {
  const Value v = EvalSource("=ARRAYTOTEXT({1,\"x\";TRUE,#N/A},1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "{1,\"x\";TRUE,#N/A}");
}

TEST(TextArrayToText, ScalarErrorArgPropagates) {
  // A lone error argument propagates like any other function; only errors
  // *inside* a multi-cell array render as their display text.
  const Value v = EvalSource("=ARRAYTOTEXT(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(TextArrayToText, InvalidFormatReturnsValueError) {
  const Value v = EvalSource("=ARRAYTOTEXT({1,2},2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextArrayToText, RangeConcisePreservesRectangleCells) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(2));
  sheet.set_cell_value(1, 0, Value::number(3));
  sheet.set_cell_value(1, 1, Value::number(4));

  const Value v = EvalSourceIn("=ARRAYTOTEXT(A1:B2)", wb, sheet);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1, 2, 3, 4");
}

TEST(TextArrayToText, RangeStrictUsesArrayLiteralShape) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::text("x"));
  sheet.set_cell_value(1, 0, Value::boolean(true));
  sheet.set_cell_value(1, 1, Value::error(ErrorCode::NA));

  const Value v = EvalSourceIn("=ARRAYTOTEXT(A1:B2,1)", wb, sheet);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "{1,\"x\";TRUE,#N/A}");
}

TEST(TextArrayToText, SpillRangeStrictUsesCommittedSpillShape) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells = {Value::number(10), Value::number(20), Value::number(30)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  const Value v = EvalSourceIn("=ARRAYTOTEXT(A1#,1)", wb, sheet);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "{10;20;30}");
}

// ---------------------------------------------------------------------------
// Near-integer snap of integer arguments
// ---------------------------------------------------------------------------

TEST(BuiltinsTextSnap, LeftRightCountSnaps) {
  const Value left = EvalSource("=LEN(LEFT(\"abcdef\",2.9999999))");
  ASSERT_TRUE(left.is_number());
  EXPECT_EQ(left.as_number(), 3.0);
  const Value right = EvalSource("=LEN(RIGHT(\"abcdef\",2.9999999))");
  ASSERT_TRUE(right.is_number());
  EXPECT_EQ(right.as_number(), 3.0);
}

TEST(BuiltinsTextSnap, LeftNegativeRawCountIsValue) {
  const Value v = EvalSource("=LEFT(\"abc\",-0.0000001)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsTextSnap, FindStartSnaps) {
  const Value v = EvalSource("=FIND(\"b\",\"abcb\",2.9999999)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsTextSnap, ReplaceStartSnaps) {
  const Value v = EvalSource("=REPLACE(\"abcdef\",1.9999999,1,\"x\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "axcdef");
}

TEST(BuiltinsTextSnap, SubstituteInstanceSnaps) {
  const Value v = EvalSource("=SUBSTITUTE(\"aaa\",\"a\",\"b\",1.9999999)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "aba");
}

TEST(BuiltinsTextSnap, MidLengthDoesNotSnap) {
  const Value v = EvalSource("=LEN(MID(\"abcdef\",1,2.9999999))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsTextSnap, SearchStartSnaps) {
  const Value v = EvalSource("=SEARCH(\"b\",\"abcb\",2.9999999)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsTextSnap, SearchbStartSnaps) {
  const Value v = EvalSource("=SEARCHB(\"b\",\"abcb\",2.9999999)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsTextSnap, FindbStartSnaps) {
  const Value v = EvalSource("=FINDB(\"b\",\"abcb\",2.9999999)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsTextSnap, LeftbCountSnaps) {
  const Value v = EvalSource("=LENB(LEFTB(\"abcdef\",2.9999999))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsTextSnap, RightbCountSnaps) {
  const Value v = EvalSource("=LENB(RIGHTB(\"abcdef\",2.9999999))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsTextSnap, ReplacebStartSnaps) {
  const Value v = EvalSource("=REPLACEB(\"abcdef\",1.9999999,1,\"x\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "axcdef");
}

TEST(BuiltinsTextSnap, TextBeforeInstanceSnaps) {
  const Value v = EvalSource("=TEXTBEFORE(\"a-b-c\",\"-\",1.9999999)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a-b");
}

TEST(BuiltinsTextSnap, TextAfterInstanceSnaps) {
  const Value v = EvalSource("=TEXTAFTER(\"a-b-c\",\"-\",1.9999999)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "c");
}

TEST(ExactRemainderAndAngles, PercentKeepsEveryDigitAfterACarry) {
  const Value v = EvalSource("=TEXT(VALUE(\"1E+21\"),\"0%\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "100000000000000000000000%");
  const Value w = EvalSource("=TEXT(VALUE(\"1E+22\"),\"0.0%\")");
  ASSERT_TRUE(w.is_text());
  EXPECT_EQ(w.as_text(), "1000000000000000000000000.0%");
}

TEST(TextLengthLimits, FormattedNumberIsAtMost255Characters) {
  const Value ok = EvalSource("=LEN(TEXT(VALUE(\"1E+254\"),\"0\"))");
  ASSERT_TRUE(ok.is_number());
  EXPECT_EQ(ok.as_number(), 255.0);
  for (const char* formula :
       {"=TEXT(VALUE(\"-1E+254\"),\"0\")", "=TEXT(VALUE(\"1E+255\"),\"0\")", "=TEXT(VALUE(\"1E+250\"),\"#,##0\")",
        "=FIXED(VALUE(\"1E+250\"),0)", "=FIXED(VALUE(\"-1E+254\"),0,TRUE)", "=DOLLAR(VALUE(\"1E+250\"),0)"}) {
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }
  const Value fixed = EvalSource("=LEN(FIXED(VALUE(\"1E+254\"),0,TRUE))");
  ASSERT_TRUE(fixed.is_number());
  EXPECT_EQ(fixed.as_number(), 255.0);
  const Value rept = EvalSource("=LEN(REPT(\"x\",309)&\"\")");
  ASSERT_TRUE(rept.is_number());
  EXPECT_EQ(rept.as_number(), 309.0);
}

TEST(TextLengthLimits, FormatWithMoreThan127DecimalPlacesIsValue) {
  const Value ok = EvalSource("=LEN(TEXT(1,\"0.\"&REPT(\"0\",127)))");
  ASSERT_TRUE(ok.is_number());
  EXPECT_EQ(ok.as_number(), 129.0);
  for (const char* formula : {"=TEXT(1,\"0.\"&REPT(\"0\",128))", "=TEXT(1,\"0.\"&REPT(\"#\",128))", "=FIXED(1,128)"}) {
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
