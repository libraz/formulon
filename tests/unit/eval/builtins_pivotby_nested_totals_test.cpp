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

TEST(PivotBy, RowTotalDepthTwoAddsASubtotalRowPerOuterRowGroup) {
  // Row keys are (outer, inner); the outer level is "A" (two rows) and "B"
  // (one row). Each outer group gets a subtotal row carrying its per-column
  // aggregates and its row total.
  const Value v =
      EvalSrc("=PIVOTBY({\"A\",\"x\";\"A\",\"y\";\"B\",\"x\"}, {\"X\";\"Y\";\"X\"}, {10;20;30}, SUM, 0, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Col-axis label row + 3 body rows + 2 subtotals + grand total = 7.
  ASSERT_EQ(v.as_array_rows(), 7U);
  // 2 row-key cols + 2 col groups + grand-total col = 5.
  ASSERT_EQ(v.as_array_cols(), 5U);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "x");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 10.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 20.0);
  // Subtotal for outer group A: 10 under X, 20 under Y, 30 in total.
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "A");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 1))) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 10.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 20.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 30.0);
  // Outer group B has no Y data, so that cell stays blank.
  EXPECT_EQ(std::string(Cell(v, 5, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 5, 2).as_number(), 30.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 5, 3))) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 5, 4).as_number(), 30.0);
  // Grand total closes the block, promoted to 総計 now that subtotal rows
  // share the same column.
  EXPECT_EQ(std::string(Cell(v, 6, 0).as_text()), "総計");
  EXPECT_DOUBLE_EQ(Cell(v, 6, 4).as_number(), 60.0);
}

TEST(PivotBy, RowTotalDepthNegativeTwoPutsEverySubtotalAboveItsGroup) {
  const Value v =
      EvalSrc("=PIVOTBY({\"A\",\"x\";\"A\",\"y\";\"B\",\"x\"}, {\"X\";\"Y\";\"X\"}, {10;20;30}, SUM, 0, -2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 7U);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "総計");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 3, 1).as_text()), "x");
  EXPECT_EQ(std::string(Cell(v, 4, 1).as_text()), "y");
  EXPECT_EQ(std::string(Cell(v, 5, 0).as_text()), "B");
}

TEST(PivotBy, RowTotalDepthTwoWithOneRowKeyColumnKeepsTheGrandTotalOnlyLayout) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"A\";\"B\"}, {\"X\";\"Y\";\"X\"}, {10;20;30}, SUM, 0, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Col-axis row + 2 body rows + grand total; a single row-key column
  // has no outer level to roll up, so ±2 degrades to the ordinary ±1
  // grand-total-only layout.
  ASSERT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "合計");
}

TEST(PivotBy, RowTotalDepthTwoEmitsNoDiagnostic) {
  // Row subtotals are implemented, so the degraded-layout warning must not
  // fire on this path. Asserted against a log sink rather than captured
  // stderr: logging ships off, so a stderr assertion would hold even if the
  // code did warn.
  ::formulon::test::LogRecorder log;
  ASSERT_TRUE(log.probe_and_clear()) << "log sink is not carrying records; the assertion below would be vacuous";
  const Value v =
      EvalSrc("=PIVOTBY({\"A\",\"x\";\"A\",\"y\";\"B\",\"x\"}, {\"X\";\"Y\";\"X\"}, {10;20;30}, SUM, 0, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_TRUE(log.empty()) << "unexpected diagnostic: " << log.joined();
}

TEST(PivotBy, ColTotalDepthTwoEmitsNoDiagnostic) {
  // Column subtotals are implemented, so the ±2 column request must not
  // announce a degraded grand-total-only layout.
  ::formulon::test::LogRecorder log;
  ASSERT_TRUE(log.probe_and_clear()) << "log sink is not carrying records; the assertion below would be vacuous";
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\",\"p\";\"X\",\"q\"}, {10;20}, SUM, 0, 1, , 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_TRUE(log.empty()) << "unexpected diagnostic: " << log.joined();
}

TEST(PivotBy, ColTotalDepthPositiveTwoMatchesMacExcelObservedMatrix) {
  // Mac Excel 16.111.3 ja-JP observation (V=1):
  //   [blank, X, X, Y, Y, 総計]
  //   [blank, M, blank, N, blank, blank]
  //   [合計, 4, 4, 2, 2, 6]
  //   [A, 4, 4, blank, blank, 4]
  //   [B, blank, blank, 2, 2, 2]
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\"}, {\"X\",\"M\";\"Y\",\"N\";\"X\",\"M\"},"
      "         {1;2;3}, SUM, 0, -1,, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_EQ(v.as_array_cols(), 6U);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 0)));
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 4).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 5).as_text()), "総計");
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "M");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 2)));
  EXPECT_EQ(std::string(Cell(v, 1, 3).as_text()), "N");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 4)));
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 4).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 5).as_number(), 6.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 4.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 3)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 4)));
  EXPECT_DOUBLE_EQ(Cell(v, 3, 5).as_number(), 4.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 1)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 2)));
  EXPECT_DOUBLE_EQ(Cell(v, 4, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 4).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 5).as_number(), 2.0);
}

TEST(PivotBy, ColTotalDepthNegativeTwoMatchesMacExcelObservedMatrix) {
  // The negative depth puts the hierarchy grand total first and each outer
  // subtotal immediately before its leaf blocks.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\"}, {\"X\",\"M\";\"Y\",\"N\";\"X\",\"M\"},"
      "         {1;2;3}, SUM, 0, -1,, -2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_EQ(v.as_array_cols(), 6U);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 0)));
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "総計");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 4).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 5).as_text()), "Y");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 1)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 2)));
  EXPECT_EQ(std::string(Cell(v, 1, 3).as_text()), "M");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 4)));
  EXPECT_EQ(std::string(Cell(v, 1, 5).as_text()), "N");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 6.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 4).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 5).as_number(), 2.0);
}

TEST(PivotBy, ColTotalDepthTwoTilesEveryValueColumn) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\"}, {\"X\",\"M\";\"Y\",\"N\";\"X\",\"M\"},"
      "         {1,10;2,20;3,30}, SUM, 0, 0,, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 4U);
  ASSERT_EQ(v.as_array_cols(), 11U);  // key + (leaf, subtotal) * 2 outers * V=2 + grand V=2
  // A: X leaf/subtotal = (4,40), Y is empty; B has only Y.
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 4).as_number(), 40.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 5)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 6)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 7)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 8)));
  EXPECT_DOUBLE_EQ(Cell(v, 3, 5).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 6).as_number(), 20.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 7).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 8).as_number(), 20.0);
}

TEST(PivotBy, RowAndColumnNestedSubtotalsIntersectAtBothOuterKeys) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\",\"x\";\"A\",\"y\";\"B\",\"x\"},"
      "         {\"X\",\"M\";\"Y\",\"N\";\"X\",\"M\"},"
      "         {10;20;30}, SUM, 0, 2,, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 8U);
  ASSERT_EQ(v.as_array_cols(), 7U);
  // A subtotal row: X leaf/subtotal=10, Y leaf/subtotal=20.
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "A");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 1)));
  EXPECT_DOUBLE_EQ(Cell(v, 4, 2).as_number(), 10.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 3).as_number(), 10.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 4).as_number(), 20.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 5).as_number(), 20.0);
  // B subtotal has no Y intersection, so both Y blocks remain blank.
  EXPECT_EQ(std::string(Cell(v, 6, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 6, 2).as_number(), 30.0);
  EXPECT_DOUBLE_EQ(Cell(v, 6, 3).as_number(), 30.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 6, 4)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 6, 5)));
}

TEST(PivotBy, ColSubtotalSortKeepsOuterGroupsContiguous) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"A\";\"A\"},"
      "         {\"Y\",\"N\";\"X\",\"M\";\"Y\",\"O\";\"X\",\"P\"},"
      "         {1;2;3;4}, SUM, 0, 0, , 2, 1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_cols(), 8U);
  // col_sort_order=1 orders leaves by totals Y(1), X(2), Y(3), X(4), but
  // the subtotal block plan re-bundles those leaves by outer key: Y,Y,Y,
  // then X,X,X.
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 4).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 5).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 6).as_text()), "X");
}

TEST(PivotBy, ColSubtotalFilterAndErrorRemainLocal) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"},"
      "         {\"X\",\"M\";\"Y\",\"N\";\"X\",\"M\";\"Y\",\"N\"},"
      "         {0;5;4;8}, LAMBDA(v, 1/SUM(v)), 0, 0,, 2,,"
      "         {TRUE;FALSE;TRUE;TRUE})");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 4U);
  ASSERT_EQ(v.as_array_cols(), 6U);
  // A/X is a local #DIV/0! in both its leaf and subtotal; A/Y was filtered
  // out and stays blank. B remains fully calculable.
  EXPECT_EQ(Cell(v, 2, 1).as_error(), ErrorCode::Div0);
  EXPECT_EQ(Cell(v, 2, 2).as_error(), ErrorCode::Div0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 3)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 4)));
  EXPECT_DOUBLE_EQ(Cell(v, 3, 1).as_number(), 0.25);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 0.25);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 0.125);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 0.125);
}

TEST(PivotBy, ColTotalDepthTwoWithOneColumnLevelKeepsOrdinaryLayout) {
  ::formulon::test::LogRecorder log;
  ASSERT_TRUE(log.probe_and_clear()) << "log sink is not carrying records; the assertion below would be vacuous";
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0,, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_cols(), 4U);  // key + two leaves + ordinary grand total
  // Row 0 is the col-axis label row; row 1 is the first body row (A).
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 2)));
  EXPECT_TRUE(log.empty()) << "unexpected diagnostic: " << log.joined();
}

TEST(PivotBy, DefaultDepthEmitsNoDiagnostic) {
  // The ordinary ±1 grand-total path must NOT emit the fallback warning.
  ::formulon::test::LogRecorder log;
  ASSERT_TRUE(log.probe_and_clear()) << "log sink is not carrying records; the assertion below would be vacuous";
  const Value v = EvalSrc("=PIVOTBY({\"A\",\"x\";\"A\",\"y\";\"B\",\"x\"}, {\"X\";\"Y\";\"X\"}, {10;20;30}, SUM, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_TRUE(log.empty()) << "unexpected diagnostic on the ±1 path: " << log.joined();
}

}  // namespace
}  // namespace eval
}  // namespace formulon
