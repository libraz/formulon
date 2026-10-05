// PIVOTBY tests grouped by output layout and argument surface.

#include <cstdint>
#include <string>
#include <string_view>

#include "builtins_pivotby_test_helpers.h"
#include "util/test_log_recorder.h"

namespace formulon {
namespace eval {
namespace {
using namespace pivotby_test_helpers;

TEST(PivotBy, BareSumName) {
  // No headers (field_headers=0), no totals (row_total_depth=0,
  // col_total_depth=0). The col-axis label row (X/Y) is still emitted --
  // it identifies the pivot's columns independent of field_headers --
  // followed by the A/B body rows.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  // Col-axis row: [blank corner, "X", "Y"].
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 2.0);
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
}

TEST(PivotBy, BareSumFiltersRangeSourcedNonNumbers) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"X\";\"X\";\"X\"},"
      "         {1;TRUE;\"text\";2}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 2.0);
}

TEST(PivotBy, BareSumInvokedWhenRangeFilterKeepsNothing) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"A\"}, {\"X\";\"X\"}, {TRUE;\"text\"}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 0.0);
}

TEST(PivotBy, BareAverageName) {
  // (A, X) = avg(10, 30) = 20; (A, Y) = 40; (B, X) = 20; (B, Y) = avg(60, 80) = 70.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"A\";\"B\";\"B\";\"B\"}, {\"X\";\"X\";\"Y\";\"X\";\"Y\";\"Y\"},"
      "         {10;30;40;20;60;80}, AVERAGE, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);  // (A, X)
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 40.0);  // (A, Y)
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 20.0);  // (B, X)
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 70.0);  // (B, Y)
}

TEST(PivotBy, BareCountAName) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, COUNTA, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 1.0);
}

TEST(PivotBy, UnknownBareNameYieldsNameError) {
  // `NOPE_NAME` resolves neither in NameEnv nor in the registry; the raw
  // NameRef evaluation surfaces #NAME?.
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, NOPE_NAME, 0, 0,, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(PivotBy, NameBoundLambdaViaLet) {
  const Value v = EvalSrc(
      "=LET(agg, LAMBDA(v, SUM(v)),"
      "     PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "             {1;2;3;4}, agg, 0, 0,, 0))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
}

TEST(PivotBy, NonLambdaNonNameAggregatorYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, 42, 0,, 0,, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, LambdaWrongArityYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, LAMBDA(a, b, a+b), 0, 0,, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, LambdaWithTrailingOptionalIsAccepted) {
  // Same fixture as `BasicTwoByTwoLambdaAggregatorSum`, with an optional
  // param added to the aggregator.
  const Value v = EvalSrc(
      "=PIVOTBY({\"R\";\"A\";\"A\";\"B\";\"B\"}, {\"C\";\"X\";\"Y\";\"X\";\"Y\"},"
      "         {\"V\";1;2;3;4}, LAMBDA(v, [u], SUM(v)))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
}

TEST(PivotBy, OptionalParamLambdaMatchesPlainLambda) {
  // `EvalSrc` resets the shared arenas per call, so the first result is
  // copied out as plain doubles before the second formula runs.
  double plain[4] = {0.0, 0.0, 0.0, 0.0};
  {
    const Value v = EvalSrc(
        "=PIVOTBY({\"R\";\"A\";\"A\";\"B\";\"B\"}, {\"C\";\"X\";\"Y\";\"X\";\"Y\"},"
        "         {\"V\";1;2;3;4}, LAMBDA(v, SUM(v)))");
    ASSERT_TRUE(v.is_array()) << v.debug_to_string();
    ASSERT_EQ(v.as_array_rows(), 4U);
    plain[0] = Cell(v, 1, 1).as_number();
    plain[1] = Cell(v, 1, 2).as_number();
    plain[2] = Cell(v, 2, 1).as_number();
    plain[3] = Cell(v, 2, 2).as_number();
  }
  const Value with_opt = EvalSrc(
      "=PIVOTBY({\"R\";\"A\";\"A\";\"B\";\"B\"}, {\"C\";\"X\";\"Y\";\"X\";\"Y\"},"
      "         {\"V\";1;2;3;4}, LAMBDA(v, [u], SUM(v)))");
  ASSERT_TRUE(with_opt.is_array()) << with_opt.debug_to_string();
  ASSERT_EQ(with_opt.as_array_rows(), 4U);
  EXPECT_DOUBLE_EQ(plain[0], Cell(with_opt, 1, 1).as_number());
  EXPECT_DOUBLE_EQ(plain[1], Cell(with_opt, 1, 2).as_number());
  EXPECT_DOUBLE_EQ(plain[2], Cell(with_opt, 2, 1).as_number());
  EXPECT_DOUBLE_EQ(plain[3], Cell(with_opt, 2, 2).as_number());
}

TEST(PivotBy, OmittedParamIsBoundForIsomitted) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"R\";\"A\";\"A\";\"B\";\"B\"}, {\"C\";\"X\";\"Y\";\"X\";\"Y\"},"
      "         {\"V\";1;2;3;4}, LAMBDA(v, [u], IF(ISOMITTED(u), SUM(v), -1)))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
}

TEST(PivotBy, LetBoundLambdaWithOptionalIsAccepted) {
  // Form B resolves the aggregator through the name environment, a
  // separate branch from the inline literal.
  const Value v = EvalSrc(
      "=LET(f, LAMBDA(v, [u], SUM(v)),"
      "     PIVOTBY({\"R\";\"A\";\"A\";\"B\";\"B\"}, {\"C\";\"X\";\"Y\";\"X\";\"Y\"},"
      "             {\"V\";1;2;3;4}, f))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
}

TEST(PivotBy, ZeroParamLambdaStillRejected) {
  // The other edge: PIVOTBY supplies one argument, so a lambda declaring
  // no params at all cannot take it.
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, LAMBDA(42), 0, 0,, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, AggregatorReturningArrayYieldsCalcInThatCell) {
  // SEQUENCE returns multi-cell; each body cell becomes #CALC!.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\"}, {\"X\";\"Y\";\"X\"},"
      "         {1;2;3}, LAMBDA(v, SEQUENCE(1, 3)), 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_TRUE(Cell(v, 1, 1).is_error());
  EXPECT_EQ(Cell(v, 1, 1).as_error(), ErrorCode::Calc);
  ASSERT_TRUE(Cell(v, 1, 2).is_error());
  EXPECT_EQ(Cell(v, 1, 2).as_error(), ErrorCode::Calc);
  // (B, X) is non-empty so the aggregator runs and surfaces #CALC! too.
  ASSERT_TRUE(Cell(v, 2, 1).is_error());
  EXPECT_EQ(Cell(v, 2, 1).as_error(), ErrorCode::Calc);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
