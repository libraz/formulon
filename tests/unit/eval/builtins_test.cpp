//
// End-to-end tests for the registered built-in functions and the special-
// cased `IF` short-circuit. Each test parses a formula source, evaluates
// the AST through the default registry, and asserts the resulting Value.

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
// SUM
// ---------------------------------------------------------------------------

TEST(BuiltinsSum, SingleArgument) {
  const Value v = EvalSource("=SUM(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsSum, ThreeIntegers) {
  const Value v = EvalSource("=SUM(1,2,3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsSum, FractionalArguments) {
  const Value v = EvalSource("=SUM(1.5, 2.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsSum, BoolsCoerceToNumbers) {
  const Value v = EvalSource("=SUM(TRUE, FALSE, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsSum, NumericTextCoerces) {
  const Value v = EvalSource("=SUM(\"2\", 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsSum, NonNumericTextYieldsValue) {
  const Value v = EvalSource("=SUM(\"abc\", 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSum, ErrorPropagates) {
  const Value v = EvalSource("=SUM(1, #REF!, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsSum, LeftMostErrorWins) {
  const Value v = EvalSource("=SUM(1, #DIV/0!, #REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsSum, OverflowYieldsNum) {
  const Value v = EvalSource("=SUM(1e308, 1e308)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsSum, EmptyArgListIsArityViolation) {
  const Value v = EvalSource("=SUM()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// CONCAT / CONCATENATE
// ---------------------------------------------------------------------------

TEST(BuiltinsConcat, SingleString) {
  const Value v = EvalSource("=CONCAT(\"a\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a");
}

TEST(BuiltinsConcat, ThreeStrings) {
  const Value v = EvalSource("=CONCAT(\"a\",\"b\",\"c\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(BuiltinsConcat, NumbersStringify) {
  const Value v = EvalSource("=CONCAT(1,2,3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "123");
}

TEST(BuiltinsConcat, BoolsStringify) {
  const Value v = EvalSource("=CONCAT(TRUE,\"-\",FALSE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "TRUE-FALSE");
}

TEST(BuiltinsConcat, ErrorPropagates) {
  const Value v = EvalSource("=CONCAT(\"x\", #REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsConcatenate, AliasMatchesConcat) {
  const Value v = EvalSource("=CONCATENATE(\"a\",\"b\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ab");
}

TEST(BuiltinsConcatenate, EmptyArgListIsArityViolation) {
  const Value v = EvalSource("=CONCATENATE()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// LEN
// ---------------------------------------------------------------------------

TEST(BuiltinsLen, EmptyString) {
  const Value v = EvalSource("=LEN(\"\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsLen, AsciiString) {
  const Value v = EvalSource("=LEN(\"hello\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsLen, NumberCoercedToString) {
  const Value v = EvalSource("=LEN(123)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsLen, BoolCoercedToString) {
  const Value v = EvalSource("=LEN(TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsLen, BmpCharactersCountAsOneEach) {
  // "あいう" = 3 BMP codepoints, 3 UTF-16 units.
  const Value v = EvalSource("=LEN(\"\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsLen, SupplementaryPlaneCountsAsTwo) {
  // "🎉" U+1F389 -> surrogate pair -> 2 UTF-16 units.
  const Value v = EvalSource("=LEN(\"\xF0\x9F\x8E\x89\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsLen, ErrorPropagates) {
  const Value v = EvalSource("=LEN(#N/A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsLen, TooManyArgsIsArityViolation) {
  const Value v = EvalSource("=LEN(\"a\",\"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// IF (short-circuit, special-cased in the evaluator)
// ---------------------------------------------------------------------------

TEST(BuiltinsIf, TrueBranchSelected) {
  const Value v = EvalSource("=IF(TRUE, \"yes\", \"no\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "yes");
}

TEST(BuiltinsIf, FalseBranchSelected) {
  const Value v = EvalSource("=IF(FALSE, \"yes\", \"no\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "no");
}

TEST(BuiltinsIf, TruthyNumberSelectsTrueBranch) {
  const Value v = EvalSource("=IF(1, \"yes\", \"no\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "yes");
}

TEST(BuiltinsIf, ZeroSelectsFalseBranch) {
  const Value v = EvalSource("=IF(0, \"yes\", \"no\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "no");
}

TEST(BuiltinsIf, OmittedFalseBranchTrueCase) {
  const Value v = EvalSource("=IF(TRUE, \"yes\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "yes");
}

TEST(BuiltinsIf, OmittedFalseBranchFalseCaseReturnsBooleanFalse) {
  const Value v = EvalSource("=IF(FALSE, \"yes\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIf, ConditionErrorPropagates) {
  const Value v = EvalSource("=IF(#REF!, 1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// The two short-circuit tests below are the load-bearing ones: if either
// fails, the IF branch is no longer being skipped at evaluation time.
TEST(BuiltinsIf, TrueBranchShortCircuitsDivByZero) {
  const Value v = EvalSource("=IF(TRUE, 1, 1/0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsIf, FalseBranchShortCircuitsDivByZero) {
  const Value v = EvalSource("=IF(FALSE, 1/0, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// IF over an array condition inside a range-aware aggregator.
//
// `SUM(IF(cond, a, b))` is the legacy CSE array idiom, and in Excel 365 it
// needs no Ctrl+Shift+Enter. An array condition has no short-circuit: both
// branches are evaluated and picked per cell, so the aggregator's argument
// expander must reach the same broadcast the bare `IF` already used instead
// of trying to coerce a rectangle to one boolean.
// ---------------------------------------------------------------------------

namespace {

/// A workbook whose A1:A5 holds 1..5, the fixture the masking cases read.
Workbook AscendingColumn() {
  Workbook wb = Workbook::create();
  for (std::uint32_t row = 0; row < 5U; ++row) {
    wb.sheet(0).set_cell_value(row, 0, Value::number(static_cast<double>(row + 1U)));
  }
  return wb;
}

}  // namespace

TEST(BuiltinsIf, ArrayConditionMasksTheRangeUnderAnAggregator) {
  Workbook wb = AscendingColumn();

  const Value masked = EvalSourceIn("=SUM(IF(A1:A5<=3,A1:A5,0))", wb, wb.sheet(0));
  ASSERT_TRUE(masked.is_number()) << "kind=" << static_cast<int>(masked.kind());
  EXPECT_EQ(masked.as_number(), 6.0) << "1+2+3, the cells the condition kept";

  const Value counted = EvalSourceIn("=SUM(IF(A1:A5<=3,1,0))", wb, wb.sheet(0));
  ASSERT_TRUE(counted.is_number()) << "kind=" << static_cast<int>(counted.kind());
  EXPECT_EQ(counted.as_number(), 3.0) << "the branches broadcast against the condition";
}

TEST(BuiltinsIf, ArrayConditionIsNotSpecificToSum) {
  Workbook wb = AscendingColumn();

  const Value largest = EvalSourceIn("=MAX(IF(A1:A5<=3,A1:A5,0))", wb, wb.sheet(0));
  ASSERT_TRUE(largest.is_number()) << "kind=" << static_cast<int>(largest.kind());
  EXPECT_EQ(largest.as_number(), 3.0);

  const Value smallest = EvalSourceIn("=MIN(IF(A1:A5<=3,A1:A5,99))", wb, wb.sheet(0));
  ASSERT_TRUE(smallest.is_number()) << "kind=" << static_cast<int>(smallest.kind());
  EXPECT_EQ(smallest.as_number(), 1.0);

  // The masked array is five cells, two of them the zero the else branch
  // supplies, so the mean is over five and not over the three kept cells.
  const Value mean = EvalSourceIn("=AVERAGE(IF(A1:A5<=3,A1:A5,0))", wb, wb.sheet(0));
  ASSERT_TRUE(mean.is_number()) << "kind=" << static_cast<int>(mean.kind());
  EXPECT_DOUBLE_EQ(mean.as_number(), 1.2);
}

TEST(BuiltinsIf, ArrayConditionAggregatesWhatTheBareFormSpills) {
  // A range-shaped argument is resolved by one of two seams depending on
  // whether the consuming function is eager (SUM) or lazy (COUNT). Neither
  // may disagree with the bare form about one formula, so the expected
  // numbers are derived from what the bare `IF` spills rather than written
  // out twice.
  Workbook wb = AscendingColumn();

  const Value spilled = EvalSourceIn("=IF(A1:A5<=3,A1:A5,0)", wb, wb.sheet(0));
  ASSERT_TRUE(spilled.is_array()) << "kind=" << static_cast<int>(spilled.kind());
  ASSERT_EQ(spilled.as_array_rows(), 5U);
  ASSERT_EQ(spilled.as_array_cols(), 1U);
  double total = 0.0;
  double numeric_cells = 0.0;
  for (std::uint32_t i = 0; i < 5U; ++i) {
    const Value& cell = spilled.as_array()->cells[i];
    ASSERT_TRUE(cell.is_number()) << "cell " << i;
    total += cell.as_number();
    numeric_cells += 1.0;
  }

  const Value eager = EvalSourceIn("=SUM(IF(A1:A5<=3,A1:A5,0))", wb, wb.sheet(0));
  ASSERT_TRUE(eager.is_number()) << "kind=" << static_cast<int>(eager.kind());
  EXPECT_EQ(eager.as_number(), total);

  const Value lazy = EvalSourceIn("=COUNT(IF(A1:A5<=3,A1:A5,0))", wb, wb.sheet(0));
  ASSERT_TRUE(lazy.is_number()) << "kind=" << static_cast<int>(lazy.kind());
  EXPECT_EQ(lazy.as_number(), numeric_cells);
}

TEST(BuiltinsIf, ArrayConditionReachesLazyRangeConsumers) {
  // `COUNT` is the one that mattered most here. It registers
  // `propagate_errors = false` so it can inspect an error argument, which
  // also means a failed argument expansion arrives as an error in its value
  // list and is dropped as "not a number" — a plausible zero rather than a
  // visible failure. `SUMPRODUCT` and `INDEX` fail loudly instead.
  Workbook wb = AscendingColumn();

  const Value counted = EvalSourceIn("=COUNT(IF(A1:A5<=3,1,0))", wb, wb.sheet(0));
  ASSERT_TRUE(counted.is_number()) << "kind=" << static_cast<int>(counted.kind());
  EXPECT_EQ(counted.as_number(), 5.0) << "every cell of the masked array is a number";

  const Value paired = EvalSourceIn("=SUMPRODUCT(IF(A1:A5<=3,1,0),A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(paired.is_number()) << "kind=" << static_cast<int>(paired.kind());
  EXPECT_EQ(paired.as_number(), 6.0) << "1*1 + 1*2 + 1*3, the rest multiplied by the zero mask";

  const Value indexed = EvalSourceIn("=INDEX(IF(A1:A5<=3,A1:A5,0),2)", wb, wb.sheet(0));
  ASSERT_TRUE(indexed.is_number()) << "kind=" << static_cast<int>(indexed.kind());
  EXPECT_EQ(indexed.as_number(), 2.0);

  // The masked array is a range argument like any other, so the criteria
  // and lookup families read it too.
  const Value matched = EvalSourceIn("=MATCH(3,IF(A1:A5<=3,A1:A5,0),0)", wb, wb.sheet(0));
  ASSERT_TRUE(matched.is_number()) << "kind=" << static_cast<int>(matched.kind());
  EXPECT_EQ(matched.as_number(), 3.0);
}

TEST(BuiltinsIf, ArrayConditionUnderALazyConsumerKeepsBooleanElseUncounted) {
  // A two-arity `IF` supplies boolean FALSE, which COUNT does not count,
  // so the lazy seam reports only the cells the condition kept.
  Workbook wb = AscendingColumn();
  const Value v = EvalSourceIn("=COUNT(IF(A1:A5<=3,A1:A5))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsIf, ArrayConditionOmittedElseBranchStaysBoolean) {
  // A two-arity `IF` supplies boolean FALSE where the condition is false,
  // and SUM ignores booleans, so the omitted branch contributes nothing
  // rather than erroring or counting as zero-valued cells.
  Workbook wb = AscendingColumn();
  const Value v = EvalSourceIn("=SUM(IF(A1:A5<=3,A1:A5))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsIf, NestedArrayConditionsCompose) {
  // The inner `IF` is itself an array-condition call in a branch position.
  Workbook wb = AscendingColumn();
  const Value v = EvalSourceIn("=SUM(IF(A1:A5<=3,IF(A1:A5>1,A1:A5,0),0))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), 5.0) << "2+3: row 1 fails the inner test, rows 4-5 the outer";
}

TEST(BuiltinsIf, ArrayConditionLiteralSelectsPerCell) {
  Workbook wb = AscendingColumn();
  const Value v = EvalSourceIn("=SUM(IF({TRUE;FALSE;TRUE;FALSE;TRUE},A1:A5,0))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), 9.0) << "1+3+5";
}

TEST(BuiltinsIf, ScalarConditionUnderAnAggregatorStillShortCircuits) {
  // Admitting the array condition must not cost the scalar path its
  // short-circuit: the untaken branch is still never evaluated, so a
  // division by zero parked there does not surface.
  Workbook wb = AscendingColumn();

  const Value taken = EvalSourceIn("=SUM(IF(TRUE,A1:A5,1/0))", wb, wb.sheet(0));
  ASSERT_TRUE(taken.is_number()) << "kind=" << static_cast<int>(taken.kind());
  EXPECT_EQ(taken.as_number(), 15.0);

  const Value other = EvalSourceIn("=SUM(IF(FALSE,1/0,A1:A5))", wb, wb.sheet(0));
  ASSERT_TRUE(other.is_number()) << "kind=" << static_cast<int>(other.kind());
  EXPECT_EQ(other.as_number(), 15.0);
}

TEST(BuiltinsIf, ScalarConditionUnderALazyConsumerStillShortCircuits) {
  // The same obligation on the second argument seam: the untaken branch is
  // still never evaluated, so the division by zero parked in it does not
  // surface. `COUNT` would report the failure as a zero rather than an
  // error, so this is the seam where losing the short-circuit would be
  // quietest.
  Workbook wb = AscendingColumn();

  const Value taken = EvalSourceIn("=COUNT(IF(TRUE,A1:A5,1/0))", wb, wb.sheet(0));
  ASSERT_TRUE(taken.is_number()) << "kind=" << static_cast<int>(taken.kind());
  EXPECT_EQ(taken.as_number(), 5.0);

  const Value other = EvalSourceIn("=COUNT(IF(FALSE,1/0,A1:A5))", wb, wb.sheet(0));
  ASSERT_TRUE(other.is_number()) << "kind=" << static_cast<int>(other.kind());
  EXPECT_EQ(other.as_number(), 5.0);

  const Value product = EvalSourceIn("=SUMPRODUCT(IF(TRUE,A1:A5,1/0))", wb, wb.sheet(0));
  ASSERT_TRUE(product.is_number()) << "kind=" << static_cast<int>(product.kind());
  EXPECT_EQ(product.as_number(), 15.0);
}

TEST(BuiltinsIf, ArrayConditionUnderAnAggregatorLandsErrorsPerCell) {
  // An error in one condition cell belongs to that cell. SUM propagates the
  // first one it meets rather than the whole call failing to coerce.
  Workbook wb = AscendingColumn();
  const Value v = EvalSourceIn("=SUM(IF({TRUE;\"x\";TRUE;TRUE;TRUE},A1:A5,0))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIf, OneArgIsArityViolation) {
  const Value v = EvalSource("=IF(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIf, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=IF()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Unknown function
// ---------------------------------------------------------------------------

TEST(BuiltinsDispatch, UnknownNameYieldsName) {
  const Value v = EvalSource("=FOOBAR(1,2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
