//
// End-to-end tests for the text built-in functions: UPPER, LOWER, TRIM,
// LEFT, RIGHT, MID, REPT, SUBSTITUTE, FIND, SEARCH, EXACT. Each test
// parses a formula source, evaluates the AST through the default registry,
// and asserts the resulting Value. TEXT / VALUE / NUMBERVALUE live in
// `builtins_value_numbervalue_test.cpp` next to their shared format engine.

#include <string_view>
#include <utility>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
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
  // ASCII-only fold: "café" (UTF-8 \xC3\xA9) -> "CAFé". Only ASCII bytes
  // change. Excel uses Unicode case folding here, but the MVP intentionally
  // covers ASCII only.
  const Value v = EvalSource("=UPPER(\"caf\xC3\xA9\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "CAF\xC3\xA9");
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
  const Value v = EvalSource("=LOWER(\"CAF\xC3\xA9\")");
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
  const Value v = EvalSource(
      "=TRIM(\"\xE3\x80\x80"
      "a\xE3\x80\x80\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a");
}

TEST(TextTrim, CollapsesIdeographicWithAscii) {
  // Mixed ideographic + ASCII spaces between non-space chars collapse to a
  // single ASCII space, mirroring Excel's behaviour of normalising any run
  // of trimmable spaces to one U+0020.
  const Value v = EvalSource(
      "=TRIM(\"a\xE3\x80\x80 \xE3\x80\x80"
      "b\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a b");
}

// The interior rule for U+3000 on its own, without an ASCII space in the
// run to carry the result: a run of ideographic spaces between two
// non-space characters becomes exactly one U+0020, so the trimmed text is
// shorter in both bytes and UTF-16 units than the input. This is the
// ja-JP everyday case, and it is what makes a later FIND / MID position
// land where the caller expects.
TEST(TextTrim, SingleInteriorIdeographicSpaceBecomesAsciiSpace) {
  // "あ　い" -> "あ い"
  const Value v = EvalSource("=TRIM(\"\xE3\x81\x82\xE3\x80\x80\xE3\x81\x84\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82 \xE3\x81\x84");

  const Value len = EvalSource("=LEN(TRIM(\"\xE3\x81\x82\xE3\x80\x80\xE3\x81\x84\"))");
  ASSERT_TRUE(len.is_number());
  EXPECT_DOUBLE_EQ(len.as_number(), 3.0);
}

TEST(TextTrim, RepeatedInteriorIdeographicSpacesCollapseToOne) {
  // "あ　　い" -> "あ い"
  const Value v = EvalSource("=TRIM(\"\xE3\x81\x82\xE3\x80\x80\xE3\x80\x80\xE3\x81\x84\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82 \xE3\x81\x84");
}

TEST(TextTrim, EdgeAndInteriorIdeographicSpacesTogether) {
  // "　　あ　　い　　" -> "あ い": the edge runs go away entirely and the
  // interior one becomes a single ASCII space.
  const Value v = EvalSource(
      "=TRIM(\"\xE3\x80\x80\xE3\x80\x80\xE3\x81\x82\xE3\x80\x80\xE3\x80\x80\xE3\x81\x84\xE3\x80\x80\xE3\x80\x80\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x81\x82 \xE3\x81\x84");
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

TEST(TextTrim, IdeographicSpaceAdjacentToTabIsStillCollapsed) {
  // "a　\t　b" -> "a \t b". The tab ends the run, so the ideographic space
  // on either side of it is interior and collapses to one ASCII space each
  // rather than being deleted.
  const Value v = EvalSource(
      "=TRIM(\"a\xE3\x80\x80\t\xE3\x80\x80"
      "b\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a \t b");
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

}  // namespace
}  // namespace eval
}  // namespace formulon
