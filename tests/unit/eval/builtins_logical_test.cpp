//
// End-to-end tests for the logical built-in functions: TRUE, FALSE, NOT,
// AND, OR (eager) and IFERROR, IFNA (lazy short-circuit). Each test parses
// a formula source, evaluates the AST through the default registry, and
// asserts the resulting Value.

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

}  // namespace
}  // namespace eval
}  // namespace formulon
