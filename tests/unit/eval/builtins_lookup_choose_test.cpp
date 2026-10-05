// Lookup builtin tests grouped by function.

#include <cstdint>
#include <string>
#include <string_view>

#include "builtins_lookup_test_helpers.h"
#include "eval/builtins.h"
#include "util/test_eval_helpers.h"

namespace formulon {
namespace eval {
namespace {
using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;
using namespace lookup_test_helpers;

TEST(BuiltinsChoose, BasicSelect) {
  const Value v = EvalSource("=CHOOSE(2, \"a\", \"b\", \"c\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsChoose, FractionalIndexTruncates) {
  // 2.9 -> floor to 2, selects "b" (not rounding to 3).
  const Value v = EvalSource("=CHOOSE(2.9, \"a\", \"b\", \"c\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsChoose, OutOfRangeZeroIsValueError) {
  const Value v = EvalSource("=CHOOSE(0, \"a\", \"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChoose, OutOfRangeTooLargeIsValueError) {
  const Value v = EvalSource("=CHOOSE(4, \"a\", \"b\", \"c\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChoose, IndexErrorPropagates) {
  const Value v = EvalSource("=CHOOSE(#DIV/0!, \"a\", \"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsChoose, ChosenArgErrorPropagates) {
  const Value v = EvalSource("=CHOOSE(2, 1, #N/A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsChoose, ChosenArgIsCellRef) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(11.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(22.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(33.0));
  const Value v = EvalSourceIn("=CHOOSE(3, A1, A2, A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 33.0);
}

TEST(BuiltinsChoose, LargeArity) {
  // Ten values; pick the 7th.
  const Value v = EvalSource("=CHOOSE(7, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsChoose, UnselectedArgNotEvaluated) {
  // Non-chosen arg contains a #DIV/0! literal that would propagate if
  // evaluated. CHOOSE must short-circuit so the result is "a".
  const Value v = EvalSource("=CHOOSE(1, \"a\", #DIV/0!)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a");
}

TEST(BuiltinsChoose, ZeroArityIsValueError) {
  EXPECT_EQ(EvalSource("=CHOOSE()").as_error(), ErrorCode::Value);
  // Just an index without any values is also invalid.
  EXPECT_EQ(EvalSource("=CHOOSE(1)").as_error(), ErrorCode::Value);
}

TEST(BuiltinsChoose, ArrayIndexSelectsScalarBranches) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  const Value v = EvalSourceIn("=CHOOSE(A1:B2,10,20)", wb, wb.sheet(0));
  ExpectArrayShape(v, 2U, 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 10.0);
}

TEST(BuiltinsChoose, HorizontalIndexTilesVerticalBranches) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(11.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(12.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 4, Value::number(22.0));
  const Value v = EvalSourceIn("=CHOOSE(A1:B1,D1:D2,E1:E2)", wb, wb.sheet(0));
  ExpectArrayShape(v, 2U, 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 21.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 12.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 22.0);
}

TEST(BuiltinsChoose, VerticalIndexTilesHorizontalBranches) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(12.0));
  wb.sheet(0).set_cell_value(0, 5, Value::number(21.0));
  wb.sheet(0).set_cell_value(0, 6, Value::number(22.0));
  const Value v = EvalSourceIn("=CHOOSE(A1:A2,D1:E1,F1:G1)", wb, wb.sheet(0));
  ExpectArrayShape(v, 2U, 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 12.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 21.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 22.0);
}

TEST(BuiltinsChoose, ArrayIndexUsesSelectedBranchShape) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(22.0));
  const Value v = EvalSourceIn("=CHOOSE(A1:B1,9,D1:D2)", wb, wb.sheet(0));
  ExpectArrayShape(v, 2U, 2U);
  for (std::uint32_t i = 0; i < 4U; ++i) {
    EXPECT_DOUBLE_EQ(v.as_array_cells()[i].as_number(), 9.0);
  }
}

TEST(BuiltinsChoose, UnequalBranchShapesBroadcastShortAxes) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(11.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(12.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(21.0));
  wb.sheet(0).set_cell_value(0, 5, Value::number(22.0));
  const Value v = EvalSourceIn("=CHOOSE(A1:B1,D1:D2,E1:F1)", wb, wb.sheet(0));
  ExpectArrayShape(v, 2U, 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 22.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 12.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 22.0);
}

TEST(BuiltinsChoose, ArrayIndexSelectedBranchErrorIsLaneLocal) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  const Value v = EvalSourceIn("=CHOOSE(A1:B1,10,1/0)", wb, wb.sheet(0));
  ExpectArrayShape(v, 1U, 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 10.0);
  ASSERT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(v.as_array_cells()[1].as_error(), ErrorCode::Div0);
}

TEST(BuiltinsChoose, ArraySelectorReferenceBlankCountsInCountA) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=COUNTA(CHOOSE(SEQUENCE(1),A1:A3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);

  const Value scalar = EvalSourceIn("=COUNTA(CHOOSE(1,A1:A3))", wb, wb.sheet(0));
  ASSERT_TRUE(scalar.is_number());
  EXPECT_DOUBLE_EQ(scalar.as_number(), 2.0);
}

TEST(BuiltinsChoose, ArraySelectorReferenceBlankCountsAndNestedIsBlank) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(1.0));

  const Value count = EvalSourceIn("=COUNT(CHOOSE(SEQUENCE(1),A1:A3))", wb, wb.sheet(0));
  ASSERT_TRUE(count.is_number());
  EXPECT_DOUBLE_EQ(count.as_number(), 2.0);

  const Value blank = EvalSourceIn("=ISBLANK(INDEX(CHOOSE(SEQUENCE(1),A1:A3),2,1))", wb, wb.sheet(0));
  ASSERT_TRUE(blank.is_boolean());
  EXPECT_TRUE(blank.as_boolean());
}

TEST(BuiltinsChoose, ArraySelectorRangeExpanderEvaluatesIndexOnce) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(1.0));

  FunctionRegistry registry;
  register_builtins(registry);
  ASSERT_TRUE(registry.register_function(FunctionDef{"TICKARRAY", 0U, 0U, &TickArrayImpl}));

  std::uint32_t calls = 0U;
  g_tick_array_calls = &calls;
  const Value value = EvalSourceInWithRegistry("=COUNTA(CHOOSE(TICKARRAY(),A1:A3))", wb, wb.sheet(0), registry);
  g_tick_array_calls = nullptr;

  ASSERT_TRUE(value.is_number());
  EXPECT_DOUBLE_EQ(value.as_number(), 3.0);
  EXPECT_EQ(calls, 1U);
}

TEST(BuiltinsChoose, ChooseAsRangeProducerInSumPicksMiddleRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(40.0));
  // Picks A1:A3 -> 10 + 20 + 30 = 60.
  const Value v = EvalSourceIn("=SUM(CHOOSE(2, A1:A2, A1:A3, A1:A4))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 60.0);
}

TEST(BuiltinsChoose, ChooseAsRangeProducerInAverage) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  // CHOOSE(1, A1:A2, ...) -> AVERAGE(10, 20) = 15.
  const Value v = EvalSourceIn("=AVERAGE(CHOOSE(1, A1:A2, A1:A3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsChoose, ChooseAsRangeProducerSkipsTextInSum) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("Ranges"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(23.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(45.0));
  // CHOOSE(2, A1:A2, A1:A3, A1:A4) -> SUM over A1:A3, text in A1 is
  // skipped per range-provenance filter.
  const Value v = EvalSourceIn("=SUM(CHOOSE(2, A1:A2, A1:A3, A1:A4))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 68.0);
}

TEST(BuiltinsChoose, ChooseScalarChildIsTreatedAsOneCellRange) {
  // CHOOSE's scalar child is treated as a 1-cell range by the aggregator
  // (faf447f). Verified against Mac Excel 365 via xlwings: scalar
  // SUM(CHOOSE(1, 7, A1:A2)) = 7, range SUM(CHOOSE(2, 7, A1:A2)) = 3.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value scalar = EvalSourceIn("=SUM(CHOOSE(1, 7, A1:A2))", wb, wb.sheet(0));
  ASSERT_TRUE(scalar.is_number());
  EXPECT_DOUBLE_EQ(scalar.as_number(), 7.0);
  const Value range = EvalSourceIn("=SUM(CHOOSE(2, 7, A1:A2))", wb, wb.sheet(0));
  ASSERT_TRUE(range.is_number());
  EXPECT_DOUBLE_EQ(range.as_number(), 3.0);
}

TEST(BuiltinsChoose, ChooseOutOfRangeIndexIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=SUM(CHOOSE(5, A1:A2, A1:A3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChoose, ChooseErrorIndexPropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=SUM(CHOOSE(NA(), A1:A2))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsChoose, ChooseNestedRecurses) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(40.0));
  // Outer CHOOSE picks slot 2 -> inner CHOOSE(1, A1:A3, A1:A4) -> A1:A3.
  // SUM(A1:A3) = 60.
  const Value v = EvalSourceIn("=SUM(CHOOSE(2, A1:A2, CHOOSE(1, A1:A3, A1:A4)))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 60.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
