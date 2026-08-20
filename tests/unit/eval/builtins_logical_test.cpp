//
// End-to-end tests for the logical built-in functions: TRUE, FALSE, NOT,
// AND, OR (eager) and IFERROR, IFNA (lazy short-circuit). Each test parses
// a formula source, evaluates the AST through the default registry, and
// asserts the resulting Value.

#include <array>
#include <cstdint>
#include <string_view>

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
// TRUE
// ---------------------------------------------------------------------------

TEST(BuiltinsTrue, NoArgsReturnsTrue) {
  const Value v = EvalSource("=TRUE()");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsTrue, KeywordTokenIsAlwaysBoolLiteral) {
  // The tokenizer maps `TRUE` to a `Bool` token unconditionally, so the
  // call dispatcher is never reached for `=TRUE(1)` or `=TRUE`: both
  // parse as the bool literal `TRUE` (the trailing `(1)` is consumed as
  // an adjacent atom that does not influence the result here).
  const Value bare = EvalSource("=TRUE");
  ASSERT_TRUE(bare.is_boolean());
  EXPECT_TRUE(bare.as_boolean());

  const Value parened = EvalSource("=TRUE(1)");
  ASSERT_TRUE(parened.is_boolean());
  EXPECT_TRUE(parened.as_boolean());
}

// ---------------------------------------------------------------------------
// FALSE
// ---------------------------------------------------------------------------

TEST(BuiltinsFalse, NoArgsReturnsFalse) {
  const Value v = EvalSource("=FALSE()");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsFalse, KeywordTokenIsAlwaysBoolLiteral) {
  // Same parser behaviour as TRUE: `FALSE` always tokenizes as the bool
  // literal `FALSE`, so the call path is never reached.
  const Value bare = EvalSource("=FALSE");
  ASSERT_TRUE(bare.is_boolean());
  EXPECT_FALSE(bare.as_boolean());

  const Value parened = EvalSource("=FALSE(1)");
  ASSERT_TRUE(parened.is_boolean());
  EXPECT_FALSE(parened.as_boolean());
}

// ---------------------------------------------------------------------------
// NOT
// ---------------------------------------------------------------------------

TEST(BuiltinsNot, NotTrueIsFalse) {
  const Value v = EvalSource("=NOT(TRUE())");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsNot, NotFalseIsTrue) {
  const Value v = EvalSource("=NOT(FALSE())");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsNot, NonZeroNumberCoercesToTrue) {
  const Value v = EvalSource("=NOT(1)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsNot, ZeroCoercesToFalse) {
  const Value v = EvalSource("=NOT(0)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsNot, NumericTextRejected) {
  // Mac Excel 365 (text_bool_not_one_string golden): NOT("1") -> #VALUE!.
  // Numeric-text never coerces to a Bool through this path.
  const Value v = EvalSource("=NOT(\"1\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsNot, NonNumericTextYieldsValue) {
  const Value v = EvalSource("=NOT(\"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsNot, ErrorPropagates) {
  const Value v = EvalSource("=NOT(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsNot, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=NOT()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsNot, TwoArgsIsArityViolation) {
  const Value v = EvalSource("=NOT(1,2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// AND
// ---------------------------------------------------------------------------

TEST(BuiltinsAnd, SingleTrueIsTrue) {
  const Value v = EvalSource("=AND(TRUE())");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsAnd, AllTrueIsTrue) {
  const Value v = EvalSource("=AND(TRUE(), TRUE(), TRUE())");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsAnd, AnyFalseIsFalse) {
  const Value v = EvalSource("=AND(TRUE(), FALSE())");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsAnd, NonZeroNumbersAreTrue) {
  const Value v = EvalSource("=AND(1, 2, 3)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsAnd, ZeroMakesFalse) {
  const Value v = EvalSource("=AND(1, 0)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsAnd, NumericTextIsValue) {
  // AND / OR / XOR use a stricter logical-coercion rule than NOT: a
  // numeric text like "1" is NOT accepted. Only the literal strings
  // "TRUE" / "FALSE" (case-insensitive, trimmed) map to a bool. Matches
  // Mac Excel 365's IronCalc-oracle behaviour.
  const Value v = EvalSource("=AND(\"1\", 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsAnd, BoolLiteralTextCoerces) {
  const Value v = EvalSource("=AND(\"TRUE\", 2)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsAnd, EmptyStringIsSkipped) {
  // An empty text is treated as "no value here"; since the remaining
  // TRUE is the only real input the result is TRUE.
  const Value v = EvalSource("=AND(\"\", TRUE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsAnd, AllEmptyIsValue) {
  // Every argument is skipped -> no logical value produced -> #VALUE!.
  const Value v = EvalSource("=AND(\"\", \"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsAnd, NonNumericTextYieldsValue) {
  const Value v = EvalSource("=AND(\"abc\", 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsAnd, LeftMostErrorWins) {
  const Value v = EvalSource("=AND(1, #DIV/0!, #REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsAnd, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=AND()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// AND / OR — array / spill / dynamic-array / LET provenance. These arrive
// through the unified lazy range-argument gate. An array constant and a
// dynamic-array producer contribute their Bool / Number cells (Text / Blank
// cells are skipped, not coerced), and a LET-bound range keeps range shape.
// ---------------------------------------------------------------------------

TEST(BuiltinsAnd, ArrayConstantAllTrue) {
  const Value v = EvalSource("=AND({TRUE,TRUE})");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsAnd, ArrayConstantWithFalse) {
  // Previously `AND({TRUE,FALSE})` surfaced #VALUE! because the array
  // literal never reached the range gate; it now evaluates to FALSE.
  const Value v = EvalSource("=AND({TRUE,FALSE})");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsAnd, ArrayConstantSkipsText) {
  // Text inside an array is skipped (range provenance), so only the bool
  // contributes: TRUE.
  const Value v = EvalSource("=AND({TRUE,\"x\"})");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsAnd, LetBoundRangeEvaluatesEveryCell) {
  // A LET binding to a range must evaluate every cell, not just the anchor:
  // A1:A3 = [TRUE, FALSE, TRUE] -> AND is FALSE.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(1, 0, Value::boolean(false));
  wb.sheet(0).set_cell_value(2, 0, Value::boolean(true));
  const Value v = EvalSourceIn("=LET(r, A1:A3, AND(r))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsOr, DynamicArrayFromFilter) {
  // OR over a FILTER result (a dynamic array) must expand the array. A1:A3
  // are values, B1:B3 the include flags; FILTER keeps rows where B is TRUE,
  // yielding {FALSE;TRUE} -> OR is TRUE.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(false));
  wb.sheet(0).set_cell_value(1, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(2, 0, Value::boolean(false));
  wb.sheet(0).set_cell_value(0, 1, Value::boolean(true));
  wb.sheet(0).set_cell_value(1, 1, Value::boolean(true));
  wb.sheet(0).set_cell_value(2, 1, Value::boolean(false));
  const Value v = EvalSourceIn("=OR(FILTER(A1:A3, B1:B3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsOr, ArrayConstantAllFalse) {
  const Value v = EvalSource("=OR({FALSE,FALSE})");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// OR
// ---------------------------------------------------------------------------

TEST(BuiltinsOr, SingleFalseIsFalse) {
  const Value v = EvalSource("=OR(FALSE())");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsOr, AnyTrueIsTrue) {
  const Value v = EvalSource("=OR(FALSE(), FALSE(), TRUE())");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsOr, AllZeroIsFalse) {
  const Value v = EvalSource("=OR(0, 0, 0)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsOr, AnyNonZeroIsTrue) {
  const Value v = EvalSource("=OR(0, 0, 1)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsOr, NumericTextIsValue) {
  // Same strict-coercion rule as AND: numeric text "0"/"1" does not map
  // to a bool, so the left-most failure surfaces #VALUE!.
  const Value v = EvalSource("=OR(\"0\", \"0\", \"1\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsOr, EmptyStringIsSkipped) {
  const Value v = EvalSource("=OR(\"\", TRUE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsOr, AllEmptyIsValue) {
  const Value v = EvalSource("=OR(\"\", \"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsOr, ErrorPropagates) {
  const Value v = EvalSource("=OR(0, #REF!, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsOr, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=OR()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// IFERROR (lazy short-circuit; ANY error triggers fallback)
// ---------------------------------------------------------------------------

TEST(BuiltinsIferror, NumericPrimaryReturnsAsIs) {
  const Value v = EvalSource("=IFERROR(1, \"fallback\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsIferror, TextPrimaryReturnsAsIs) {
  const Value v = EvalSource("=IFERROR(\"ok\", \"fallback\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ok");
}

TEST(BuiltinsIferror, Div0TriggersFallback) {
  const Value v = EvalSource("=IFERROR(1/0, \"fallback\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "fallback");
}

TEST(BuiltinsIferror, RefTriggersFallback) {
  const Value v = EvalSource("=IFERROR(#REF!, \"fallback\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "fallback");
}

TEST(BuiltinsIferror, NaTriggersFallback) {
  const Value v = EvalSource("=IFERROR(#N/A, \"fallback\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "fallback");
}

TEST(BuiltinsIferror, FallbackErrorPropagates) {
  const Value v = EvalSource("=IFERROR(1/0, 1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// Critical short-circuit assertion: when the primary is non-error, the
// fallback subtree must NOT be evaluated. If it were, `1/0` would surface
// as #DIV/0! and clobber the expected `1`.
TEST(BuiltinsIferror, FallbackIsNotEvaluatedWhenPrimaryClean) {
  const Value v = EvalSource("=IFERROR(1, 1/0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsIferror, OneArgIsArityViolation) {
  const Value v = EvalSource("=IFERROR(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIferror, ThreeArgsIsArityViolation) {
  const Value v = EvalSource("=IFERROR(1,2,3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// IFNA (lazy short-circuit; ONLY #N/A triggers fallback)
// ---------------------------------------------------------------------------

TEST(BuiltinsIfna, NaTriggersFallback) {
  const Value v = EvalSource("=IFNA(#N/A, \"fallback\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "fallback");
}

TEST(BuiltinsIfna, NumericPrimaryReturnsAsIs) {
  const Value v = EvalSource("=IFNA(1, \"fallback\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

// Critical filter: only `#N/A` triggers IFNA. `#DIV/0!` must propagate
// unchanged, NOT be replaced by the fallback.
TEST(BuiltinsIfna, Div0PropagatesNotReplaced) {
  const Value v = EvalSource("=IFNA(#DIV/0!, \"fallback\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsIfna, RefPropagatesNotReplaced) {
  const Value v = EvalSource("=IFNA(#REF!, \"fallback\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIfna, ValuePropagatesNotReplaced) {
  const Value v = EvalSource("=IFNA(#VALUE!, \"fallback\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIfna, TextPrimaryReturnsAsIs) {
  const Value v = EvalSource("=IFNA(\"text\", \"fallback\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "text");
}

// Critical short-circuit assertion: the fallback subtree must not be
// evaluated when the primary is non-#N/A.
TEST(BuiltinsIfna, FallbackIsNotEvaluatedWhenPrimaryClean) {
  const Value v = EvalSource("=IFNA(1, 1/0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsIfna, OneArgIsArityViolation) {
  const Value v = EvalSource("=IFNA(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIfna, BlankPrimaryStaysBlank) {
  // IFNA passes a non-#N/A primary through unchanged, preserving Blank as
  // Blank exactly like IF / IFERROR do. ISBLANK observes the value before
  // the grid-level Blank->0 promotion, so it reports TRUE.
  const Value v = EvalSource("=ISBLANK(IFNA(,\"text\"))");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIfna, BlankFallbackStaysBlank) {
  // When #N/A triggers the fallback and the fallback evaluates to Blank,
  // the Blank is returned as Blank (consistent with IFERROR's fallback).
  const Value v = EvalSource("=ISBLANK(IFNA(#N/A,))");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

// Blank handling must be identical across IF / IFERROR / IFNA: a blank
// result is returned as Blank, so ISBLANK(...) is TRUE for all three.
TEST(BuiltinsBlankSymmetry, IfIferrorIfnaAllPreserveBlank) {
  const Value if_blank = EvalSource("=ISBLANK(IF(TRUE,))");
  ASSERT_TRUE(if_blank.is_boolean());
  EXPECT_TRUE(if_blank.as_boolean());

  const Value iferror_blank = EvalSource("=ISBLANK(IFERROR(#N/A,))");
  ASSERT_TRUE(iferror_blank.is_boolean());
  EXPECT_TRUE(iferror_blank.as_boolean());

  const Value ifna_blank = EvalSource("=ISBLANK(IFNA(#N/A,))");
  ASSERT_TRUE(ifna_blank.is_boolean());
  EXPECT_TRUE(ifna_blank.as_boolean());
}

TEST(BuiltinsIfna, ThreeArgsIsArityViolation) {
  const Value v = EvalSource("=IFNA(1,2,3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Bool coercion follow-up: Mac Excel 365 (ja-JP) probes (`tests/oracle/
// golden/text_to_bool_probes.golden.json`) settle the rule for
// `coerce_to_bool` (the path taken by IF / NOT / IFERROR / BETA.DIST's
// cumulative argument and friends): only the EXACT (no leading or
// trailing whitespace) ASCII case-insensitive literals "TRUE" / "FALSE"
// coerce to a Bool. Numeric-text such as "0" / "1" / "0.5", padded
// forms such as "  TRUE  ", localised truth-words ("VRAI", "WAHR",
// "真"), and arbitrary text ("yes") all surface `#VALUE!`.
// ---------------------------------------------------------------------------

TEST(BuiltinsLogicalCoerce, IfRejectsNumericText) {
  // Mac Excel 365: `=IF("1", "y", "n")` -> #VALUE! (text_num_if_one_string
  // golden). The previous Formulon behaviour (numeric-text falls through
  // to the numeric path) was wrong.
  const Value v = EvalSource("=IF(\"1\", \"y\", \"n\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsLogicalCoerce, IfAcceptsLiteralTrueText) {
  // Mac Excel: `=IF("TRUE", "y", "n")` -> "y".
  const Value v = EvalSource("=IF(\"TRUE\", \"y\", \"n\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "y");
}

TEST(BuiltinsLogicalCoerce, IfAcceptsLowercaseFalseText) {
  // Case-insensitive: lowercase "false" coerces to FALSE.
  const Value v = EvalSource("=IF(\"false\", \"y\", \"n\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "n");
}

TEST(BuiltinsLogicalCoerce, IfAcceptsMixedCaseTrueText) {
  // Mixed-case "True" also coerces to TRUE.
  const Value v = EvalSource("=IF(\"True\", \"y\", \"n\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "y");
}

TEST(BuiltinsLogicalCoerce, IfRejectsTrueWithWhitespace) {
  // Mac Excel does NOT trim whitespace around "TRUE" before bool coercion
  // (text_bool_if_true_with_whitespace golden).
  const Value v = EvalSource("=IF(\"  TRUE  \", \"y\", \"n\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsLogicalCoerce, IfRejectsArbitraryText) {
  // Only "TRUE" / "FALSE" are recognised; "yes" still surfaces #VALUE!.
  const Value v = EvalSource("=IF(\"yes\", \"y\", \"n\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// IFS — stricter logical coercion than IF: numeric-text conditions surface
// `#VALUE!` (matching Mac Excel 365's AND / OR / XOR rule), while the
// literal strings "TRUE" / "FALSE" (case-insensitive, trimmed) are
// accepted. See `src/eval/logical_coerce.h`.
// ---------------------------------------------------------------------------

TEST(BuiltinsIfs, NumericTextConditionIsValue) {
  // IF and IFS now share the same strict text rule: numeric-text such as
  // "1" surfaces #VALUE! in either context (Mac Excel 365 ja-JP).
  const Value v = EvalSource("=IFS(\"1\", 34)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIfs, NumericTextConditionSecondPositionIsValue) {
  // Second condition is numeric-text: the first branch is FALSE so we walk
  // into the second condition, and the strict rule surfaces #VALUE! there.
  const Value v = EvalSource("=IFS(FALSE, \"first\", \"1\", 7)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIfs, LiteralTrueTextConditionAccepted) {
  // The case-insensitive literal "TRUE" is the allowed text form.
  const Value v = EvalSource("=IFS(\"TRUE\", 42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

TEST(BuiltinsIfs, LiteralFalseTextConditionSkipped) {
  // Lower-case "false" is a valid FALSE, so the first branch is skipped
  // and the trailing `TRUE, "catchall"` pair wins.
  const Value v = EvalSource("=IFS(\"false\", 1, TRUE, \"catchall\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "catchall");
}

TEST(BuiltinsIfs, WhitespaceWrappedTrueTextAccepted) {
  // IFS now shares AND / OR's host-aware coercion, which trims surrounding
  // ASCII whitespace before matching "TRUE" / "FALSE": `IFS("  TRUE  ", 7)`
  // takes the first branch, mirroring `AND("  TRUE  ", TRUE)`.
  const Value v = EvalSource("=IFS(\"  TRUE  \", 7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

// ---------------------------------------------------------------------------
// Blank-vs-Bool comparison: a blank cell chameleons to the boolean FALSE, so
// it equals FALSE (and 0 and "") but not TRUE, and orders as FALSE.
// ---------------------------------------------------------------------------

TEST(OperatorBlankBool, BlankEqualsFalse) {
  Workbook wb = Workbook::create();  // A1 left blank.
  const Value v = EvalSourceIn("=(A1=FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(OperatorBlankBool, BlankNotEqualToTrue) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=(A1=TRUE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(OperatorBlankBool, BlankLessThanTrue) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=(A1<TRUE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(OperatorBlankBool, BlankNotLessThanFalse) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=(A1<FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(OperatorBlankBool, BlankGreaterOrEqualFalse) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=(A1>=FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// SWITCH equality: type-strict + case-insensitive text, with the single
// special case that a blank subject matches numeric 0 (but not "" or FALSE).
// ---------------------------------------------------------------------------

TEST(BuiltinsSwitch, BlankSubjectMatchesZero) {
  Workbook wb = Workbook::create();  // A1 blank.
  const Value v = EvalSourceIn("=SWITCH(A1, 0, \"zero\", \"other\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "zero");
}

TEST(BuiltinsSwitch, BlankSubjectDoesNotMatchEmptyText) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=SWITCH(A1, \"\", \"empty\", \"other\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "other");
}

TEST(BuiltinsSwitch, BlankSubjectDoesNotMatchFalse) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=SWITCH(A1, FALSE, \"f\", \"other\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "other");
}

TEST(BuiltinsSwitch, NoCrossTypeCoercionNumberVsText) {
  const Value v = EvalSource("=SWITCH(23, \"23\", \"x\", \"other\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "other");
}

TEST(BuiltinsSwitch, NoCrossTypeCoercionTextVsNumber) {
  const Value v = EvalSource("=SWITCH(\"23\", 23, \"x\", \"other\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "other");
}

TEST(BuiltinsSwitch, TextMatchIsCaseInsensitive) {
  const Value v = EvalSource("=SWITCH(\"a\", \"A\", \"casehit\", \"other\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "casehit");
}

// ---------------------------------------------------------------------------
// SWITCH over an array subject selects per cell. There is no branch to
// short-circuit onto when different cells take different arms, so every arm
// is evaluated once and broadcast — the scalar rule applied cellwise,
// including the trailing default and the `#N/A` a subject matching nothing
// receives.
// ---------------------------------------------------------------------------

namespace {

/// A workbook whose A1:A5 holds 1..5, the subject the array cases select on.
Workbook AscendingColumn() {
  Workbook wb = Workbook::create();
  for (std::uint32_t row = 0; row < 5U; ++row) {
    wb.sheet(0).set_cell_value(row, 0, Value::number(static_cast<double>(row + 1U)));
  }
  return wb;
}

/// Asserts `v` is a 5x1 numeric column equal to `expected`.
void ExpectColumn(const Value& v, const std::array<double, 5>& expected) {
  ASSERT_TRUE(v.is_array()) << "kind=" << static_cast<int>(v.kind());
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  for (std::uint32_t i = 0; i < 5U; ++i) {
    const Value& cell = v.as_array()->cells[i];
    ASSERT_TRUE(cell.is_number()) << "cell " << i << " kind=" << static_cast<int>(cell.kind());
    EXPECT_EQ(cell.as_number(), expected[i]) << "cell " << i;
  }
}

}  // namespace

TEST(BuiltinsSwitch, ArraySubjectSelectsPerCell) {
  Workbook wb = AscendingColumn();
  ExpectColumn(EvalSourceIn("=SWITCH(A1:A5,1,10,0)", wb, wb.sheet(0)), {10.0, 0.0, 0.0, 0.0, 0.0});
  ExpectColumn(EvalSourceIn("=SWITCH(A1:A5,1,10,2,20,0)", wb, wb.sheet(0)), {10.0, 20.0, 0.0, 0.0, 0.0});
}

TEST(BuiltinsSwitch, ArraySubjectBroadcastsAnArrayArm) {
  // A value arm may itself be a rectangle, in which case the picked cell
  // comes from the matching position rather than from the arm's first cell.
  Workbook wb = AscendingColumn();
  ExpectColumn(EvalSourceIn("=SWITCH(A1:A5,1,A1:A5,0)", wb, wb.sheet(0)), {1.0, 0.0, 0.0, 0.0, 0.0});
}

TEST(BuiltinsSwitch, ArraySubjectWithoutDefaultIsPerCellNotAvailable) {
  // The scalar form answers `#N/A` when nothing matches and no default is
  // given. Cellwise, that lands only in the cells that matched nothing.
  Workbook wb = AscendingColumn();
  const Value v = EvalSourceIn("=SWITCH(A1:A5,1,10)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << "kind=" << static_cast<int>(v.kind());
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_TRUE(v.as_array()->cells[0].is_number());
  EXPECT_EQ(v.as_array()->cells[0].as_number(), 10.0);
  for (std::uint32_t i = 1; i < 5U; ++i) {
    const Value& cell = v.as_array()->cells[i];
    ASSERT_TRUE(cell.is_error()) << "cell " << i << " kind=" << static_cast<int>(cell.kind());
    EXPECT_EQ(cell.as_error(), ErrorCode::NA) << "cell " << i;
  }
}

TEST(BuiltinsSwitch, ArraySubjectAggregatesWhatTheBareFormSpills) {
  // SWITCH is not a range-shaped call, so an aggregator consumes exactly
  // the array the bare form produces. Deriving the expectation from that
  // array rather than writing it out keeps the two from drifting.
  Workbook wb = AscendingColumn();

  const Value spilled = EvalSourceIn("=SWITCH(A1:A5,1,10,0)", wb, wb.sheet(0));
  ASSERT_TRUE(spilled.is_array()) << "kind=" << static_cast<int>(spilled.kind());
  double total = 0.0;
  double numeric_cells = 0.0;
  for (std::uint32_t i = 0; i < spilled.as_array_rows(); ++i) {
    const Value& cell = spilled.as_array()->cells[i];
    ASSERT_TRUE(cell.is_number()) << "cell " << i;
    total += cell.as_number();
    numeric_cells += 1.0;
  }

  const Value summed = EvalSourceIn("=SUM(SWITCH(A1:A5,1,10,0))", wb, wb.sheet(0));
  ASSERT_TRUE(summed.is_number()) << "kind=" << static_cast<int>(summed.kind());
  EXPECT_EQ(summed.as_number(), total);

  const Value counted = EvalSourceIn("=COUNT(SWITCH(A1:A5,1,10,0))", wb, wb.sheet(0));
  ASSERT_TRUE(counted.is_number()) << "kind=" << static_cast<int>(counted.kind());
  EXPECT_EQ(counted.as_number(), numeric_cells);
}

TEST(BuiltinsSwitch, ScalarSubjectStillShortCircuits) {
  // Admitting the array path must not cost the scalar path its
  // short-circuit. SWITCH has more untaken arms than IF does — every
  // unmatched value plus the default — so a division by zero is parked in
  // each position in turn.
  const Value unmatched_default = EvalSource("=SWITCH(1,1,10,1/0)");
  ASSERT_TRUE(unmatched_default.is_number()) << "kind=" << static_cast<int>(unmatched_default.kind());
  EXPECT_EQ(unmatched_default.as_number(), 10.0) << "the default is not evaluated once a case matches";

  const Value unmatched_value = EvalSource("=SWITCH(2,1,1/0,0)");
  ASSERT_TRUE(unmatched_value.is_number()) << "kind=" << static_cast<int>(unmatched_value.kind());
  EXPECT_EQ(unmatched_value.as_number(), 0.0) << "an unmatched case's value is not evaluated";

  const Value later_case = EvalSource("=SWITCH(1,1,10,2,1/0,0)");
  ASSERT_TRUE(later_case.is_number()) << "kind=" << static_cast<int>(later_case.kind());
  EXPECT_EQ(later_case.as_number(), 10.0) << "arms past the match are not evaluated";
}

// ---------------------------------------------------------------------------
// IFS over an array condition decides per cell. Conditions are still scanned
// lazily, so a scalar condition that wins outright never causes a later arm
// to be evaluated; the cellwise walk begins only at the first array
// condition actually reached, and every arm from there is evaluated once.
// ---------------------------------------------------------------------------

TEST(BuiltinsIfs, ArrayConditionSelectsPerCell) {
  Workbook wb = AscendingColumn();
  ExpectColumn(EvalSourceIn("=IFS(A1:A5<=3,A1:A5,TRUE,0)", wb, wb.sheet(0)), {1.0, 2.0, 3.0, 0.0, 0.0});
  // Several array conditions in a row: each cell takes the first that holds.
  ExpectColumn(EvalSourceIn("=IFS(A1:A5<=2,10,A1:A5<=4,20,TRUE,30)", wb, wb.sheet(0)), {10.0, 10.0, 20.0, 20.0, 30.0});
}

TEST(BuiltinsIfs, ScalarConditionsBeforeAnArrayConditionAreStillSkipped) {
  // A leading scalar FALSE contributes nothing and must not disturb the
  // rectangle the later array condition establishes.
  Workbook wb = AscendingColumn();
  ExpectColumn(EvalSourceIn("=IFS(FALSE,0,A1:A5<=3,A1:A5,TRUE,99)", wb, wb.sheet(0)), {1.0, 2.0, 3.0, 99.0, 99.0});
}

TEST(BuiltinsIfs, ArrayConditionWithoutCatchallIsPerCellNotAvailable) {
  // The scalar rule "no condition held -> #N/A" applied cellwise: only the
  // cells that matched nothing carry it. A trailing unpaired condition has
  // no value to return, so it cannot win a cell either.
  Workbook wb = AscendingColumn();
  for (std::string_view src : {"=IFS(A1:A5<=3,A1:A5)", "=IFS(A1:A5<=3,A1:A5,TRUE)"}) {
    const Value v = EvalSourceIn(src, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_array()) << src << " kind=" << static_cast<int>(v.kind());
    ASSERT_EQ(v.as_array_rows(), 5U) << src;
    for (std::uint32_t i = 0; i < 3U; ++i) {
      ASSERT_TRUE(v.as_array()->cells[i].is_number()) << src << " cell " << i;
      EXPECT_EQ(v.as_array()->cells[i].as_number(), static_cast<double>(i + 1U)) << src << " cell " << i;
    }
    for (std::uint32_t i = 3; i < 5U; ++i) {
      const Value& cell = v.as_array()->cells[i];
      ASSERT_TRUE(cell.is_error()) << src << " cell " << i << " kind=" << static_cast<int>(cell.kind());
      EXPECT_EQ(cell.as_error(), ErrorCode::NA) << src << " cell " << i;
    }
  }
}

TEST(BuiltinsIfs, ArrayConditionAggregatesWhatTheBareFormSpills) {
  // IFS is not a range-shaped call, so an aggregator consumes exactly the
  // array the bare form produced; deriving the expectation from that array
  // keeps the two from drifting.
  Workbook wb = AscendingColumn();

  const Value spilled = EvalSourceIn("=IFS(A1:A5<=3,A1:A5,TRUE,0)", wb, wb.sheet(0));
  ASSERT_TRUE(spilled.is_array()) << "kind=" << static_cast<int>(spilled.kind());
  double total = 0.0;
  double numeric_cells = 0.0;
  for (std::uint32_t i = 0; i < spilled.as_array_rows(); ++i) {
    const Value& cell = spilled.as_array()->cells[i];
    ASSERT_TRUE(cell.is_number()) << "cell " << i;
    total += cell.as_number();
    numeric_cells += 1.0;
  }

  const Value summed = EvalSourceIn("=SUM(IFS(A1:A5<=3,A1:A5,TRUE,0))", wb, wb.sheet(0));
  ASSERT_TRUE(summed.is_number()) << "kind=" << static_cast<int>(summed.kind());
  EXPECT_EQ(summed.as_number(), total);

  const Value counted = EvalSourceIn("=COUNT(IFS(A1:A5<=3,A1:A5,TRUE,0))", wb, wb.sheet(0));
  ASSERT_TRUE(counted.is_number()) << "kind=" << static_cast<int>(counted.kind());
  EXPECT_EQ(counted.as_number(), numeric_cells);
}

TEST(BuiltinsIfs, ScalarConditionStillShortCircuits) {
  // The untaken arms are still never evaluated when a scalar condition
  // decides the call.
  const Value first = EvalSource("=IFS(TRUE,1,FALSE,1/0)");
  ASSERT_TRUE(first.is_number()) << "kind=" << static_cast<int>(first.kind());
  EXPECT_EQ(first.as_number(), 1.0);

  const Value second = EvalSource("=IFS(FALSE,1/0,TRUE,2)");
  ASSERT_TRUE(second.is_number()) << "kind=" << static_cast<int>(second.kind());
  EXPECT_EQ(second.as_number(), 2.0) << "an unmatched condition's value is not evaluated";
}

TEST(BuiltinsIfs, ScalarWinBeforeAnArrayConditionKeepsTheScalarResult) {
  // The load-bearing case for the lazy scan: a scalar condition that wins
  // must decide the call before any later array condition is even reached,
  // so the arms past it stay unevaluated and the result stays scalar. If
  // the scan stopped being lazy this would surface the `1/0`.
  Workbook wb = AscendingColumn();
  const Value v = EvalSourceIn("=IFS(TRUE,1,A1:A5<=3,1/0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsIfs, ArrayConditionLandsArmErrorsPerCell) {
  // Once the cellwise walk starts there is no branch to short-circuit onto,
  // so every arm is evaluated and an erroring one reaches only the cells
  // that select it.
  Workbook wb = AscendingColumn();
  const Value v = EvalSourceIn("=IFS(A1:A5<=3,1/0,TRUE,0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << "kind=" << static_cast<int>(v.kind());
  ASSERT_EQ(v.as_array_rows(), 5U);
  for (std::uint32_t i = 0; i < 3U; ++i) {
    const Value& cell = v.as_array()->cells[i];
    ASSERT_TRUE(cell.is_error()) << "cell " << i << " kind=" << static_cast<int>(cell.kind());
    EXPECT_EQ(cell.as_error(), ErrorCode::Div0) << "cell " << i;
  }
  for (std::uint32_t i = 3; i < 5U; ++i) {
    const Value& cell = v.as_array()->cells[i];
    ASSERT_TRUE(cell.is_number()) << "cell " << i << " kind=" << static_cast<int>(cell.kind());
    EXPECT_EQ(cell.as_number(), 0.0) << "cell " << i;
  }
}

TEST(BuiltinsSwitch, ArraySubjectLandsArmErrorsPerCell) {
  // With no branch to short-circuit onto, every arm is evaluated — so an
  // erroring arm reaches only the cells that select it.
  Workbook wb = AscendingColumn();
  const Value v = EvalSourceIn("=SWITCH(A1:A5,1,1/0,0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << "kind=" << static_cast<int>(v.kind());
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_TRUE(v.as_array()->cells[0].is_error());
  EXPECT_EQ(v.as_array()->cells[0].as_error(), ErrorCode::Div0);
  for (std::uint32_t i = 1; i < 5U; ++i) {
    const Value& cell = v.as_array()->cells[i];
    ASSERT_TRUE(cell.is_number()) << "cell " << i << " kind=" << static_cast<int>(cell.kind());
    EXPECT_EQ(cell.as_number(), 0.0) << "cell " << i;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
