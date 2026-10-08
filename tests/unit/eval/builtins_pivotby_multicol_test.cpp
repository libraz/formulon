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

TEST(PivotBy, MismatchedRowCountsRowVsValuesYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\";\"C\"}, {\"X\";\"Y\";\"Z\"}, {1;2}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, MismatchedRowCountsColVsValuesYieldsValueError) {
  // row_fields has 2 rows, values has 2 rows, col_fields has 3 -> #VALUE!.
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\";\"Z\"}, {1;2}, SUM, 0, 0,, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, MultiColumnRowFieldsBasic) {
  // K=2, L=1, V=1, fh=0. Formulon defaults for the rest:
  //   row_total_depth=1 (totals at BOTTOM), col_total_depth=1 (totals at RIGHT).
  // rows = L + 0 + nR + 1 = 1 + 0 + 2 + 1 = 4
  // cols = K + nC*V + V = 2 + 2 + 1 = 5
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\",\"P\";\"B\",\"Q\";\"A\",\"P\"}, {\"X\";\"Y\";\"X\"},"
      "         {1;2;3}, SUM, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 5U);
  // Row 0 (col-axis label row): [null, null, "X", "Y", "合計"].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 0)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 1)));
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 4).as_text()), "合計");
  // Row 1 ((A,P) group): X=1+3=4, Y=blank, total=4.
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "P");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 4.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 3)));
  EXPECT_DOUBLE_EQ(Cell(v, 1, 4).as_number(), 4.0);
  // Row 2 ((B,Q) group): X=blank, Y=2, total=2.
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "Q");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 2)));
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 4).as_number(), 2.0);
  // Row 3 (bottom grand-total row): ["合計", null, 4, 2, 6].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "合計");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 1)));
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 6.0);
}

TEST(PivotBy, MultiColumnColFieldsBasic) {
  // K=1, L=2, V=1, fh=0.
  // rows = 2 + 0 + 2 + 1 = 5, cols = 1 + 2 + 1 = 4.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\"}, {\"X\",\"M\";\"Y\",\"N\";\"X\",\"M\"},"
      "         {1;2;3}, SUM, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 5U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  // Row 0 (outermost col-axis level): [null, "X", "Y", "合計"].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 0)));
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "合計");
  // Row 1 (innermost col-axis level): [null, "M", "N", null].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 0)));
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "M");
  EXPECT_EQ(std::string(Cell(v, 1, 2).as_text()), "N");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 3)));
  // Row 2 (A): [A, 4, blank, 4].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 4.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 2)));
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 4.0);
  // Row 3 (B): [B, blank, 2, 2].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "B");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 1)));
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 2.0);
  // Row 4 (bottom grand total): ["合計", 4, 2, 6].
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 4, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 2).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 3).as_number(), 6.0);
}

TEST(PivotBy, MultiColumnValuesBasic) {
  // K=1, L=1, V=2, fh=0.
  // rows = 1 + 0 + 2 + 1 = 4, cols = 1 + 2*2 + 2 = 7.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\"}, {\"X\";\"Y\";\"X\"},"
      "         {1,10;2,20;3,30}, SUM, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 7U);
  // Row 0 (col-axis labels tiled V=2 per col group, "合計" tiled V=2):
  //   [null, "X", "X", "Y", "Y", "合計", "合計"].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 0)));
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 4).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 5).as_text()), "合計");
  EXPECT_EQ(std::string(Cell(v, 0, 6).as_text()), "合計");
  // Row 1 (A): [A, 4, 40, blank, blank, blank, blank].
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 40.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 3)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 4)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 5)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 6)));
  // Row 2 (B): [B, blank, blank, 2, 20, blank, blank].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 1)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 2)));
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 4).as_number(), 20.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 5)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 6)));
  // Row 3 (bottom grand totals).
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 3, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 20.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 5)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 6)));
}

TEST(PivotBy, MultiColumnAllAxes) {
  // K=2, L=2, V=2, fh=0.
  // rows = 2 + 0 + 2 + 1 = 5, cols = 2 + 2*2 + 2 = 8.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\",\"P\";\"B\",\"Q\";\"A\",\"P\"},"
      "         {\"X\",\"M\";\"Y\",\"N\";\"X\",\"M\"},"
      "         {1,10;2,20;3,30}, SUM, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 5U);
  EXPECT_EQ(v.as_array_cols(), 8U);
  // Row 0 (outer col-axis): [null, null, "X", "X", "Y", "Y",
  //                          "合計", "合計"].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 0)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 1)));
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 4).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 5).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 6).as_text()), "合計");
  EXPECT_EQ(std::string(Cell(v, 0, 7).as_text()), "合計");
  // Row 1 (inner col-axis): [null, null, "M", "M", "N", "N", null, null].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 0)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 1)));
  EXPECT_EQ(std::string(Cell(v, 1, 2).as_text()), "M");
  EXPECT_EQ(std::string(Cell(v, 1, 3).as_text()), "M");
  EXPECT_EQ(std::string(Cell(v, 1, 4).as_text()), "N");
  EXPECT_EQ(std::string(Cell(v, 1, 5).as_text()), "N");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 6)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 7)));
  // Row 2 ((A,P)): [A, P, 4, 40, null, null, null, null].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "P");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 40.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 4)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 5)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 6)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 7)));
  // Row 3 ((B,Q)): [B, Q, null, null, 2, 20, null, null].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 3, 1).as_text()), "Q");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 2)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 3)));
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 5).as_number(), 20.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 6)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 7)));
  // Row 4 (bottom grand total): ["合計", null, 4, 40, 2, 20, null, null].
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "合計");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 1)));
  EXPECT_DOUBLE_EQ(Cell(v, 4, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 3).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 4).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 5).as_number(), 20.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 6)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 7)));
}

TEST(PivotBy, MultiColumnRowsWithFieldHeadersThree) {
  // K=2, L=1, V=1, fh=3 (inputs have header, output emits header).
  // rows = 1 col-field header + 1 col-axis + 1 output header + 1 total + 2 body = 6.
  const Value v = EvalSrc(
      "=PIVOTBY({\"R1\",\"R2\";\"A\",\"P\";\"B\",\"Q\";\"A\",\"P\"},"
      "         {\"C\";\"X\";\"Y\";\"X\"},"
      "         {\"V\";1;2;3}, SUM, 3)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 6U);
  EXPECT_EQ(v.as_array_cols(), 5U);
  // Row 0 (col-field header): [null, null, "C", null, null].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 0)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 1)));
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "C");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 3)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 4)));
  // Row 1 (col-axis labels from col_fields data row 0 reps):
  //   [null, null, "X", "Y", "合計"].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 0)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 1)));
  EXPECT_EQ(std::string(Cell(v, 1, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 1, 3).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 1, 4).as_text()), "合計");
  // Row 2 (header row): ["R1", "R2", "V", "V", "V"].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "R1");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "R2");
  EXPECT_EQ(std::string(Cell(v, 2, 2).as_text()), "V");
  EXPECT_EQ(std::string(Cell(v, 2, 3).as_text()), "V");
  EXPECT_EQ(std::string(Cell(v, 2, 4).as_text()), "V");
  // Row 3 ((A,P)): [A, P, 4, blank, 4].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 3, 1).as_text()), "P");
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 4.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 3)));
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 4.0);
  // Row 4 ((B,Q)): [B, Q, blank, 2, 2].
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 4, 1).as_text()), "Q");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 2)));
  EXPECT_DOUBLE_EQ(Cell(v, 4, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 4).as_number(), 2.0);
  // Row 5 (bottom grand total).
  EXPECT_EQ(std::string(Cell(v, 5, 0).as_text()), "合計");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 5, 1)));
  EXPECT_DOUBLE_EQ(Cell(v, 5, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 5, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 5, 4).as_number(), 6.0);
}

TEST(PivotBy, MultiColumnValuesWithFieldHeadersThree) {
  // K=1, L=1, V=2, fh=3.
  // rows = 1 col-field header + 1 col-axis + 1 output header + 1 total + 2 body = 6.
  const Value v = EvalSrc(
      "=PIVOTBY({\"Region\";\"A\";\"B\";\"A\"}, {\"Cat\";\"X\";\"Y\";\"X\"},"
      "         {\"S\",\"C\";1,10;2,20;3,30}, SUM, 3)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 6U);
  EXPECT_EQ(v.as_array_cols(), 7U);
  // Row 0: [null, "Cat", null, null, null, null, null].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 0)));
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "Cat");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 2)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 3)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 4)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 5)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 6)));
  // Row 1: [null, "X", "X", "Y", "Y", "合計", "合計"].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 0)));
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 1, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 1, 3).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 1, 4).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 1, 5).as_text()), "合計");
  EXPECT_EQ(std::string(Cell(v, 1, 6).as_text()), "合計");
  // Row 2: ["Region", "S", "C", "S", "C", "S", "C"].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "Region");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "S");
  EXPECT_EQ(std::string(Cell(v, 2, 2).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 2, 3).as_text()), "S");
  EXPECT_EQ(std::string(Cell(v, 2, 4).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 2, 5).as_text()), "S");
  EXPECT_EQ(std::string(Cell(v, 2, 6).as_text()), "C");
  // Row 3 (A): [A, 4, 40, blank, blank, blank, blank].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 3, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 40.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 3)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 4)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 5)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 6)));
  // Row 4 (B): [B, blank, blank, 2, 20, blank, blank].
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "B");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 1)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 2)));
  EXPECT_DOUBLE_EQ(Cell(v, 4, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 4).as_number(), 20.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 5)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 6)));
  // Row 5 (bottom grand total).
  EXPECT_EQ(std::string(Cell(v, 5, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 5, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 5, 2).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 5, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 5, 4).as_number(), 20.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 5, 5)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 5, 6)));
}

TEST(PivotBy, MultiColumnValuesWithFieldHeadersTwoSynth) {
  // K=1, L=1, V=2, fh=2 (no input header, output emits synth labels).
  // rows = 1 + 1 + 1 + 2 + 1 = 6, cols = 7.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\"}, {\"X\";\"Y\";\"X\"},"
      "         {1,10;2,20;3,30}, SUM, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 6U);
  EXPECT_EQ(v.as_array_cols(), 7U);
  // Row 0: generated col-field label over the first body column.
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "列フィールド 1");
  // Row 1 (col-axis labels from col_fields data row 0 reps):
  //   [null, "X", "X", "Y", "Y", "合計", "合計"].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 0)));
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 1, 5).as_text()), "合計");
  // Row 2 (generated header row):
  //   ["行フィールド 1", "値 1", "値 2", "値 1", "値 2", "値 1", "値 2"].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "行フィールド 1");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "値 1");
  EXPECT_EQ(std::string(Cell(v, 2, 2).as_text()), "値 2");
  EXPECT_EQ(std::string(Cell(v, 2, 3).as_text()), "値 1");
  EXPECT_EQ(std::string(Cell(v, 2, 4).as_text()), "値 2");
  EXPECT_EQ(std::string(Cell(v, 2, 5).as_text()), "値 1");
  EXPECT_EQ(std::string(Cell(v, 2, 6).as_text()), "値 2");
  // Body sanity check (full coverage in MultiColumnValuesBasic).
  EXPECT_DOUBLE_EQ(Cell(v, 3, 1).as_number(), 4.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 3, 6)));
  EXPECT_DOUBLE_EQ(Cell(v, 5, 1).as_number(), 4.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 5, 6)));
}

TEST(PivotBy, MultiColumnRowsWithFilterArray) {
  // K=2, L=1, V=1, fh=0, with a filter that drops the (B,Q) data row.
  // After filter: only (A,P) row group, only X col group. nR=1, nC=1.
  // rows = 1 + 0 + 1 + 1 = 3, cols = 2 + 1 + 1 = 4.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\",\"P\";\"B\",\"Q\";\"A\",\"P\"}, {\"X\";\"Y\";\"X\"},"
      "         {1;2;3}, SUM, 0, -1,, 1,, {TRUE;FALSE;TRUE})");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  // Row 0: [null, null, "X", "合計"].
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 0)));
  EXPECT_TRUE(IsPlaceholder(Cell(v, 0, 1)));
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "合計");
  // Row 1 (top grand total): ["合計", null, 4, 4].
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "合計");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 1, 1)));
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 3).as_number(), 4.0);
  // Row 2 ((A,P)): [A, P, 4, 4].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "P");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 4.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
