// Conditional-format evaluator tests: expression, context-aware operands, and time periods.

#include <cstdint>

#include "cell.h"
#include "cf/cf_evaluator.h"
#include "cf/cf_types.h"
#include "cf_evaluator_fixtures.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"
namespace formulon::cf {
namespace {

using test::At;
using test::CFEvalHarness;
using test::MakeRange;
using test::MakeRule;
using test::PopulateLinearPopulation;
using test::RGB;
using test::SqrefContext;
using test::TwoStopMinMax;

// The context-aware overload routes Expression rules through
// `helpers::parse_shift_evaluate`, which must canonicalise Excel's
// `_xlfn.` / `_xlpm.` storage prefixes before parsing (see cf_helpers.cpp)
// -- the CF-rule ingestion path does not go through the defined-names
// reader's own stripping, so it needs its own.
TEST(CFEvaluator, ExpressionRuleAcceptsXlfnPrefixedFunction) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("Sheet1"));
  eval::EvalState state;
  eval::EvalContext eval_ctx(wb, s, state);
  Arena arena;
  const eval::FunctionRegistry& registry = eval::default_registry();

  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "_xlfn.LET(_xlpm.x,TRUE,_xlpm.x)";

  CFEvalContext ctx;
  ctx.anchor = CellAddress{0U, 0U};
  ctx.target = CellAddress{0U, 0U};
  ctx.arena = &arena;
  ctx.registry = &registry;
  ctx.eval_ctx = &eval_ctx;

  EXPECT_TRUE(match_rule(r, Value::blank(), ctx));
}
// ---------------------------------------------------------------------------
// Context-aware overload: Expression rules and CellIs with non-literal
// formula1 / formula2.
// ---------------------------------------------------------------------------

TEST(CFEvaluator, ExpressionRuleLiteralTrueMatches) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "TRUE";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleLiteralFalseDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "FALSE";
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleNonZeroNumberMatches) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "1";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleZeroNumberDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "0";
  EXPECT_FALSE(match_rule(r, Value::number(1.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleArithmeticComparison) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(15.0));  // A1 = 15
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "A1>10";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleRelativeReferenceShiftsToTargetRow) {
  // Anchor at A1, target at A4. Formula `=A1>10` shifts to `=A4>10`.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(5.0));   // A1
  harness.sheet.set_cell_value(3, 0, Value::number(15.0));  // A4
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "A1>10";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(3, 0))));
}

TEST(CFEvaluator, ExpressionRuleRelativeReferenceShiftsToTargetColumn) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(5.0));   // A1
  harness.sheet.set_cell_value(0, 2, Value::number(20.0));  // C1
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "A1>10";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 2))));
}

TEST(CFEvaluator, ExpressionRuleAbsoluteReferenceLocksAtAnchor) {
  // `$A$1>10` always reads A1, regardless of target.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(15.0));  // A1
  harness.sheet.set_cell_value(0, 5, Value::number(0.0));   // F1
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "$A$1>10";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 5))));
}

TEST(CFEvaluator, ExpressionRuleErrorResultDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "1/0";  // produces #DIV/0!
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleTextResultDoesNotMatch) {
  // Excel's expression-rule truthiness rejects text — only bool / number trigger.
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "\"yes\"";
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleParseFailureDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "(((";
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleMissingFormulaDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  // formula1 unset
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleOutOfBoundsShiftDoesNotMatch) {
  // Anchor at A1, target at A1 with delta = (-1, 0) is impossible
  // (target above anchor by one row would land on row 0 = okay).
  // Pick a formula referencing A1 with anchor at B5 and target at A1
  // → delta = (-4, -1). Shift A1 by (-4, -1) goes out of bounds.
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "A1>10";
  // The shifted ref becomes #REF!; comparison `(#REF!)>10` propagates
  // an error, which is not truthy.
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(4, 1), At(0, 0))));
}

// ---------------------------------------------------------------------------
// CellIs with non-literal operand
// ---------------------------------------------------------------------------

TEST(CFEvaluator, CellIsEvaluatesReferenceOperand) {
  // CellIs Equal `=A1`, with A1 = 10. Cell value = 10 → matches.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(10.0));  // A1
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "A1";
  EXPECT_TRUE(match_rule(r, Value::number(10.0), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(r, Value::number(11.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, CellIsEvaluatesArithmeticOperand) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::LessThan;
  r.formula1 = "10*2";  // = 20
  EXPECT_TRUE(match_rule(r, Value::number(15.0), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(r, Value::number(25.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, CellIsBetweenWithReferenceOperands) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(5.0));   // A1
  harness.sheet.set_cell_value(0, 1, Value::number(10.0));  // B1
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Between;
  r.formula1 = "A1";
  r.formula2 = "B1";
  EXPECT_TRUE(match_rule(r, Value::number(7.5), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(r, Value::number(11.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, CellIsLiteralPathStillWorksThroughEvaluatorOverload) {
  // Verify the context-aware overload doesn't regress the literal path.
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "42";
  EXPECT_TRUE(match_rule(r, Value::number(42.0), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(r, Value::number(43.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, CellIsReferenceShiftsWithTarget) {
  // CellIs with `A1` operand, anchor at A1, target at A2 → operand is A2.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(1, 0, Value::number(50.0));  // A2
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "A1";
  EXPECT_TRUE(match_rule(r, Value::number(50.0), harness.context(At(0, 0), At(1, 0))));
}

TEST(CFEvaluator, ContextAwareOverloadDelegatesValueOnlyKinds) {
  // Verify the context-aware overload routes value-only rules to the
  // simple overload. This is a regression guard: if the dispatcher
  // ever forgets to delegate, the rule would silently return false.
  CFEvalHarness harness;
  CFRule blanks = MakeRule(RuleType::ContainsBlanks);
  EXPECT_TRUE(match_rule(blanks, Value::blank(), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(blanks, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));

  CFRule errors = MakeRule(RuleType::ContainsErrors);
  EXPECT_TRUE(match_rule(errors, Value::error(ErrorCode::Div0), harness.context(At(0, 0), At(0, 0))));

  CFRule contains = MakeRule(RuleType::ContainsText);
  contains.text = "foo";
  EXPECT_TRUE(match_rule(contains, Value::text("foobar"), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(contains, Value::text("bar"), harness.context(At(0, 0), At(0, 0))));
}

// ---------------------------------------------------------------------------
// TimePeriod
//
// Anchored at Wednesday 2024-03-13 throughout. Excel weekday semantics
// place that day at WEEKDAY=4 (Sun=1). The week therefore spans
// 2024-03-10 (Sun) through 2024-03-16 (Sat).
// ---------------------------------------------------------------------------

CFEvalContext PinnedContext(CFEvalHarness& harness, double today_serial) {
  CFEvalContext ctx = harness.context(At(0, 0), At(0, 0));
  ctx.today_serial = today_serial;
  return ctx;
}

double Serial(int year, unsigned month, unsigned day) {
  return date_time::serial_from_ymd(year, month, day);
}

TEST(CFEvaluator, TimePeriodWithoutTodaySerialDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Today;
  // ctx.today_serial intentionally left as nullopt.
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 13)), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, TimePeriodMissingBucketDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  // r.time_period not set
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 13)), PinnedContext(harness, Serial(2024, 3, 13))));
}

TEST(CFEvaluator, TimePeriodNonNumericCellDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Today;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_FALSE(match_rule(r, Value::text("2024-03-13"), ctx));
  EXPECT_FALSE(match_rule(r, Value::boolean(true), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
}

TEST(CFEvaluator, TimePeriodTodayMatchesSameDay) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Today;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 12)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 14)), ctx));
}

TEST(CFEvaluator, TimePeriodTodayDropsTimeOfDayFraction) {
  // Cell carrying date + 0.75 (= 18:00) still matches Today.
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Today;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 13) + 0.75), ctx));
}

TEST(CFEvaluator, TimePeriodYesterdayMatchesPreviousDay) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Yesterday;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 12)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 11)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));
}

TEST(CFEvaluator, TimePeriodTomorrowMatchesNextDay) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Tomorrow;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 14)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 15)), ctx));
}

TEST(CFEvaluator, TimePeriodLast7DaysSpansSixDaysBackThroughToday) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Last7Days;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 7)), ctx));   // today - 6
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 6)), ctx));  // today - 7
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 14)), ctx));
}

TEST(CFEvaluator, TimePeriodThisWeekIsSundayThroughSaturday) {
  // 2024-03-13 is Wed. Week = 2024-03-10 (Sun) .. 2024-03-16 (Sat).
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::ThisWeek;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 10)), ctx));   // Sun
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));   // Wed
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 16)), ctx));   // Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 9)), ctx));   // prior Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 17)), ctx));  // next Sun
}

TEST(CFEvaluator, TimePeriodLastWeekIsPriorSundayThroughSaturday) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::LastWeek;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 3)), ctx));    // Sun
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 9)), ctx));    // Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 2)), ctx));   // 2 weeks back
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 10)), ctx));  // this Sun
}

TEST(CFEvaluator, TimePeriodNextWeekIsFollowingSundayThroughSaturday) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::NextWeek;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 17)), ctx));   // Sun
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 23)), ctx));   // Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 16)), ctx));  // this Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 24)), ctx));  // 2 weeks ahead
}

TEST(CFEvaluator, TimePeriodThisMonthMatchesSameYearAndMonth) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::ThisMonth;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 1)), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 31)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 2, 29)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 4, 1)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2023, 3, 13)), ctx));  // same month, prior year
}

TEST(CFEvaluator, TimePeriodLastMonthHandlesYearBoundary) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::LastMonth;
  // today = 2024-01-15 → last month = 2023-12.
  const auto ctx = PinnedContext(harness, Serial(2024, 1, 15));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2023, 12, 1)), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2023, 12, 31)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 1, 1)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2023, 11, 30)), ctx));
}

TEST(CFEvaluator, TimePeriodNextMonthHandlesYearBoundary) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::NextMonth;
  // today = 2024-12-15 → next month = 2025-01.
  const auto ctx = PinnedContext(harness, Serial(2024, 12, 15));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2025, 1, 1)), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2025, 1, 31)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 12, 31)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2025, 2, 1)), ctx));
}

double Date1904Serial(int year, unsigned month, unsigned day) {
  return date_time::serial_from_ymd(year, month, day, /*date1904=*/true);
}

// Same anchor and boundary assertions as `TimePeriodThisWeekIsSundayThroughSaturday`,
// but with every serial encoded under the 1904 system and the context's
// `EvalContext` carrying that epoch. A decode that silently assumed 1900
// would land on the wrong civil date and miss these boundaries.
TEST(CFEvaluator, TimePeriodThisWeekHonorsDate1904Epoch) {
  CFEvalHarness harness;
  eval::EvalContext date1904_ctx = harness.eval_ctx.with_date1904(true);
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::ThisWeek;
  CFEvalContext ctx = PinnedContext(harness, Date1904Serial(2024, 3, 13));
  ctx.eval_ctx = &date1904_ctx;
  EXPECT_TRUE(match_rule(r, Value::number(Date1904Serial(2024, 3, 10)), ctx));   // Sun
  EXPECT_TRUE(match_rule(r, Value::number(Date1904Serial(2024, 3, 16)), ctx));   // Sat
  EXPECT_FALSE(match_rule(r, Value::number(Date1904Serial(2024, 3, 9)), ctx));   // prior Sat
  EXPECT_FALSE(match_rule(r, Value::number(Date1904Serial(2024, 3, 17)), ctx));  // next Sun
}

// Same as `TimePeriodThisMonthMatchesSameYearAndMonth`, under the 1904
// epoch: the month boundary at 2024-02-29/2024-03-01 only lands correctly
// if the decode reads the epoch the caller encoded with.
TEST(CFEvaluator, TimePeriodThisMonthHonorsDate1904Epoch) {
  CFEvalHarness harness;
  eval::EvalContext date1904_ctx = harness.eval_ctx.with_date1904(true);
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::ThisMonth;
  CFEvalContext ctx = PinnedContext(harness, Date1904Serial(2024, 3, 13));
  ctx.eval_ctx = &date1904_ctx;
  EXPECT_TRUE(match_rule(r, Value::number(Date1904Serial(2024, 3, 1)), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(Date1904Serial(2024, 3, 31)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Date1904Serial(2024, 2, 29)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Date1904Serial(2024, 4, 1)), ctx));
}

TEST(CFEvaluator, TimePeriodValueOnlyOverloadStillReturnsFalse) {
  // The value-only overload has no today reference, so TimePeriod
  // continues to return false there. Pin the contract.
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Today;
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 13))));
}

// ---------------------------------------------------------------------------
// DuplicateValues / UniqueValues
// ---------------------------------------------------------------------------

}  // namespace
}  // namespace formulon::cf
