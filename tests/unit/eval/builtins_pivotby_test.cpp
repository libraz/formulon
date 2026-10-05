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

TEST(PivotBy, BasicTwoByTwoLambdaAggregatorSum) {
  // Row groups: A, B. Col groups: X, Y.
  // Body: (A,X)=1, (A,Y)=2, (B,X)=3, (B,Y)=4.
  // With default field_headers=3, row_total_depth=1, col_total_depth=1:
  //   header row, 2 body rows, then bottom grand-total row.
  const Value v = EvalSrc(
      "=PIVOTBY({\"R\";\"A\";\"A\";\"B\";\"B\"}, {\"C\";\"X\";\"Y\";\"X\";\"Y\"},"
      "         {\"V\";1;2;3;4}, LAMBDA(v, SUM(v)))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Layout: header row + 2 body rows + bottom grand-total row = 4 rows.
  // Cols: row label + 2 col keys + grand-total = 4 cols.
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  // Header row: ["R", "X", "Y", "合計"].
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "R");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "合計");
  // A row: [A, 1, 2, 3].
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 3).as_number(), 3.0);
  // B row: [B, 3, 4, 7].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 7.0);
  // Bottom grand total row: ["合計", X-total=4, Y-total=6, total=10].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 3, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 6.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 10.0);
}

TEST(PivotBy, FieldHeadersZeroNoHeaders) {
  // No header row in the inputs and no row_fields header label emitted,
  // but the col-axis label row (which carries the pivot's column keys)
  // still renders with a blank corner cell -- Mac Excel always shows it,
  // independent of field_headers.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Row 0 is the col-axis label row; its corner cell is blank.
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  // First body row (row 1) carries the row key.
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
}

TEST(PivotBy, ZeroFieldHeadersAndZeroColTotalDepthStillEmitsColAxisRow) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0, , 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  // Row 0: blank corner, "X", "Y".
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "Y");
  // Row 1: "A", 1, blank (no B/X data point).
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_TRUE(Cell(v, 1, 2).is_blank());
  // Row 2: "B", blank (no A/Y data point), 2.
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  EXPECT_TRUE(Cell(v, 2, 1).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 2.0);
}

TEST(PivotBy, ZeroFieldHeadersAndZeroColTotalDepthLambdaAggregator) {
  // LAMBDA-form counterpart of the case above (oracle case
  // `pivotby_lambda_sum`); same shape as the bare-name case.
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, LAMBDA(v, SUM(v)), 0, 0, , 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_TRUE(Cell(v, 1, 2).is_blank());
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  EXPECT_TRUE(Cell(v, 2, 1).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 2.0);
}

TEST(PivotBy, FieldHeadersOneCopiesInputHeaders) {
  // Inputs have a header row; output also emits one. Top-left cell = the
  // row_fields header label ("Row").
  const Value v = EvalSrc(
      "=PIVOTBY({\"Row\";\"A\";\"A\";\"B\";\"B\"}, {\"Col\";\"X\";\"Y\";\"X\";\"Y\"},"
      "         {\"V\";1;2;3;4}, SUM, 1, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "Row");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "Y");
}

TEST(PivotBy, FieldHeadersTwoSynthesizesDefaults) {
  // Inputs have no header row; output emits a synthesised "Field 1".
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 2, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "Field 1");
}

TEST(PivotBy, FieldHeadersThreeBothInputsHaveAndOutputEmits) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"R\";\"A\";\"B\"}, {\"C\";\"X\";\"Y\"},"
      "         {\"V\";1;2}, SUM, 3, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "R");
}

TEST(PivotBy, FieldHeadersOutOfRangeYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, SUM, 5, 0,, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, RowTotalDepthZeroNoColumnTotals) {
  // No "column totals" row anywhere.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Col-axis row + 2 body rows, no totals row.
  EXPECT_EQ(v.as_array_rows(), 3U);
}

TEST(PivotBy, RowTotalDepthPositiveOneColumnTotalsAtBottom) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 1,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 4U);
  // Last row is the totals row: ["合計", X-total=4, Y-total=6].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 3, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 6.0);
}

TEST(PivotBy, RowTotalDepthNegativeOneColumnTotalsAtTop) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, -1,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 6.0);
}

TEST(PivotBy, RowTotalDepthOutOfRangeYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, SUM, 0, 99,, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, ColTotalDepthZeroNoRowTotals) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // No row-totals column.
  EXPECT_EQ(v.as_array_cols(), 3U);
}

TEST(PivotBy, ColTotalDepthPositiveOneRowTotalsOnRight) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0,, 1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_cols(), 4U);
  // Last column on each body row is the row-total: A->3, B->7 (row 0 is
  // the col-axis label row).
  EXPECT_DOUBLE_EQ(Cell(v, 1, 3).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 7.0);
}

TEST(PivotBy, ColTotalDepthNegativeOneRowTotalsOnLeft) {
  // col_total_depth=-1 => row totals appear immediately to the right of
  // the row label column.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0,, -1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_cols(), 4U);
  // Col layout: [row_label, row_total, X, Y]. Row 0 is the col-axis
  // label row; row 1 is the first body row (A).
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 3).as_number(), 2.0);
}

TEST(PivotBy, ColTotalDepthOutOfRangeYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, SUM, 0, 0,, 99)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, OmittedRowSortOrderPreservesFirstOccurrence) {
  // Row groups in input order: B, A, C. With the slot left out they stay
  // that way. Row 0 is the col-axis label row.
  const Value v = EvalSrc(
      "=PIVOTBY({\"B\";\"A\";\"C\";\"A\"}, {\"X\";\"X\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "C");
}

TEST(PivotBy, RowSortOrderPositiveAscendingByRowTotal) {
  // Row totals: A=2+4=6, B=1, C=3. Ascending row-total: B, C, A. Row 0 is
  // the col-axis label row.
  const Value v = EvalSrc(
      "=PIVOTBY({\"B\";\"A\";\"C\";\"A\"}, {\"X\";\"X\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, 1, 1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "A");
}

TEST(PivotBy, RowSortOrderNegativeDescendingByRowTotal) {
  // Row 0 is the col-axis label row.
  const Value v = EvalSrc(
      "=PIVOTBY({\"B\";\"A\";\"C\";\"A\"}, {\"X\";\"X\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, -1, 1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "B");
}

TEST(PivotBy, SuppliedRowSortOrderZeroYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, SuppliedColSortOrderZeroYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0, , 1, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, OmittedColSortOrderPreservesFirstOccurrence) {
  // Col groups in input order: Y, X, Z.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"A\";\"A\"}, {\"Y\";\"X\";\"Z\";\"X\"},"
      "         {1;2;3;4}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Col layout: row_label | Y | X | Z. Row 0 is the col-axis label row;
  // row 1 is the single "A" body row.
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);  // Y
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 6.0);  // X = 2+4
  EXPECT_DOUBLE_EQ(Cell(v, 1, 3).as_number(), 3.0);  // Z
}

TEST(PivotBy, ColSortOrderPositiveAscendingByColTotal) {
  // Col totals: Y=1, X=6, Z=3. Ascending col-total: Y(1), Z(3), X(6).
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"A\";\"A\"}, {\"Y\";\"X\";\"Z\";\"X\"},"
      "         {1;2;3;4}, SUM, 0, 1,, 0, 1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Layout: col-axis row (row 0, always emitted), the single "A" body
  // row (row 1), then the bottom totals row (row_total_depth=1, row 2).
  // col_total_depth=0 means no row-totals column. Cols sorted ascending:
  // row_label | Y | Z | X.
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);  // Y
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 3.0);  // Z
  EXPECT_DOUBLE_EQ(Cell(v, 1, 3).as_number(), 6.0);  // X
}

TEST(PivotBy, ColSortOrderNegativeDescendingByColTotal) {
  // Descending: X(6), Z(3), Y(1). Row 0 is the col-axis label row.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"A\";\"A\"}, {\"Y\";\"X\";\"Z\";\"X\"},"
      "         {1;2;3;4}, SUM, 0, 1,, 0, -1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 6.0);  // X
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 3.0);  // Z
  EXPECT_DOUBLE_EQ(Cell(v, 1, 3).as_number(), 1.0);  // Y
}

TEST(PivotBy, FilterArrayBasicIncludeExclude) {
  // Mask drops the second row (B, X, 3); only A rows and (B, Y, 4) survive.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\";\"B\"}, {\"X\";\"X\";\"Y\";\"Y\"},"
      "         {1;3;2;4}, SUM, 0, 0,, 0,,"
      "         {TRUE;FALSE;TRUE;TRUE})");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  // (A, X) = 1, (A, Y) = 2, (B, Y) = 4. (B, X) row was masked out -> Blank.
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 2.0);
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  // (B, X) has no surviving rows -> Blank cell.
  EXPECT_TRUE(Cell(v, 2, 1).is_blank()) << Cell(v, 2, 1).debug_to_string();
  EXPECT_FALSE(Cell(v, 2, 1).blank_projects_to_zero());
  EXPECT_FALSE(Cell(v, 2, 1).blank_counts_for_counta());
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
}

TEST(PivotBy, FilterArrayLengthMismatchYieldsValueError) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0,, 0,,"
      "         {TRUE;FALSE;TRUE})");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, FilterArrayAllExcludedYieldsValueError) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0,, 0,,"
      "         {FALSE;FALSE})");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, PerCellErrorIsolation) {
  // The aggregator divides 1 by the cell sum. (A, X) sums to 0 -> #DIV/0!
  // for that cell only; (A, Y) and (B, Y) sums are non-zero and produce
  // valid numbers.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\"}, {\"X\";\"Y\";\"Y\"},"
      "         {0;5;4}, LAMBDA(v, 1/SUM(v)), 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  // (A, X) errored.
  ASSERT_TRUE(Cell(v, 1, 1).is_error());
  EXPECT_EQ(Cell(v, 1, 1).as_error(), ErrorCode::Div0);
  // (A, Y) = 1/5 = 0.2.
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 0.2);
  // (B, X) is empty -> Blank, no aggregator call.
  EXPECT_TRUE(Cell(v, 2, 1).is_blank());
  // (B, Y) = 1/4 = 0.25.
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 0.25);
}

TEST(PivotBy, JapaneseFoldingAppliesToRowAndColKeys) {
  // Half-width "ｱ" (U+FF71) and full-width "ア" (U+30A2) fold to the same
  // key. The pivot collapses both rows into a single (row_group, col_group)
  // cell totalling 30. Same applies to the col-key axis, where two
  // half-width katakana col-keys also fold together.
  const Value v = EvalSrc(
      "=PIVOTBY({\"\xef\xbd\xb1\";\"\xe3\x82\xa2\"},"
      "         {\"\xef\xbd\xb6\";\"\xe3\x82\xab\"},"  // ｶ folds to カ
      "         {10;20}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 30.0);
}

TEST(PivotBy, RowSortOrderAbsTwoYieldsValueError) {
  // First-commit scope: only ±1 / 0 are supported on the sort orders.
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0, 2, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, ColSortOrderAbsTwoYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0,, 0, 2)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, TooFewArgsYieldsValueError) {
  // Fewer than 4 args is grammatically invalid (PIVOTBY requires the
  // first 4: row_fields, col_fields, values, function).
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1})");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
