//
// End-to-end tests for the lazy-dispatched VLOOKUP and HLOOKUP functions.
// Both share a `lookup_scan` helper in `tree_walker.cpp` that walks the
// first column (VLOOKUP) or first row (HLOOKUP) of the table_array and
// returns a cell offset by the caller-supplied column / row index.

#include <cstdint>
#include <string>
#include <string_view>

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
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;

void ExpectArrayShape(const Value& value, std::uint32_t rows, std::uint32_t cols) {
  ASSERT_TRUE(value.is_array()) << value.debug_to_string();
  EXPECT_EQ(value.as_array_rows(), rows);
  EXPECT_EQ(value.as_array_cols(), cols);
}

// ---------------------------------------------------------------------------
// VLOOKUP
// ---------------------------------------------------------------------------

TEST(BuiltinsVLookup, ExactNumericFirstColumn) {
  // First column is {10, 20, 30, 20}. Exact match (range_lookup=FALSE) on
  // 20 returns row 2's second column value.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("ten"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("twenty"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(2, 1, Value::text("thirty"));
  wb.sheet(0).set_cell_value(3, 0, Value::number(20.0));  // duplicate, ignored
  wb.sheet(0).set_cell_value(3, 1, Value::text("other twenty"));
  const Value v = EvalSourceIn("=VLOOKUP(20, A1:B4, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_text(), "twenty");
}

TEST(BuiltinsVLookup, ArrayLiteralTablePreservesTwoDimensions) {
  const Value v = EvalSource("=VLOOKUP(2,{1,\"a\";2,\"b\"},2,FALSE)");
  ASSERT_TRUE(v.is_text()) << v.debug_to_string();
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsVLookup, ArrayLookupValueExactPreservesShapeAndErrors) {
  const Value v = EvalSource("=VLOOKUP({20,99;#DIV/0!,10},{10,\"ten\";20,\"twenty\";30,\"thirty\"},2,FALSE)");
  ExpectArrayShape(v, 2U, 2U);
  const Value* cells = v.as_array_cells();
  ASSERT_TRUE(cells[0].is_text());
  EXPECT_EQ(cells[0].as_text(), "twenty");
  ASSERT_TRUE(cells[1].is_error());
  EXPECT_EQ(cells[1].as_error(), ErrorCode::NA);
  ASSERT_TRUE(cells[2].is_error());
  EXPECT_EQ(cells[2].as_error(), ErrorCode::Div0);
  ASSERT_TRUE(cells[3].is_text());
  EXPECT_EQ(cells[3].as_text(), "ten");
}

TEST(BuiltinsVLookup, ArrayLookupValueNestedSequenceMapsOnce) {
  const Value v = EvalSource("=VLOOKUP(SEQUENCE(3,1,20,-10),{10,\"ten\";20,\"twenty\";30,\"thirty\"},2,FALSE)");
  ExpectArrayShape(v, 3U, 1U);
  const Value* cells = v.as_array_cells();
  ASSERT_TRUE(cells[0].is_text());
  EXPECT_EQ(cells[0].as_text(), "twenty");
  ASSERT_TRUE(cells[1].is_text());
  EXPECT_EQ(cells[1].as_text(), "ten");
  ASSERT_TRUE(cells[2].is_error());
  EXPECT_EQ(cells[2].as_error(), ErrorCode::NA);
}

TEST(BuiltinsVLookup, ExactTextCaseInsensitive) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("apple"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("Banana"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("CHERRY"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=VLOOKUP(\"banana\", A1:B3, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsVLookup, ExactWildcardStar) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("Banana"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("Apple"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(200.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("Apricot"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(300.0));
  // "A*" matches "Apple" first.
  const Value v = EvalSourceIn("=VLOOKUP(\"A*\", A1:B3, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 200.0);
}

TEST(BuiltinsVLookup, ExactWildcardQuestion) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("cab"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("ab"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("abc"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(3.0));
  // "?b" matches exactly a 2-byte string whose 2nd char is 'b' -> "ab".
  const Value v = EvalSourceIn("=VLOOKUP(\"?b\", A1:B3, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsVLookup, ExactWildcardEscape) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("foo"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("*"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("bar"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(3.0));
  // "~*" is the literal asterisk.
  const Value v = EvalSourceIn("=VLOOKUP(\"~*\", A1:B3, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsVLookup, ExactNoMatchIsNa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  const Value v = EvalSourceIn("=VLOOKUP(99, A1:B2, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsVLookup, ApproximateMiddleHit) {
  // Ascending first column; target 7 -> last row with value <= 7 is row 2
  // (value 5), and we pull the 2nd column there.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("one"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("five"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 1, Value::text("ten"));
  const Value v = EvalSourceIn("=VLOOKUP(7, A1:B3, 2, TRUE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "five");
}

TEST(BuiltinsVLookup, ApproximateSmallerThanFirstIsNa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("five"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("ten"));
  const Value v = EvalSourceIn("=VLOOKUP(1, A1:B2, 2, TRUE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsVLookup, ApproximateBetweenCellsPicksSmaller) {
  // Target 7 falls between 5 and 10 -> picks 5 (last row with value <= 7).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(500.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(1000.0));
  const Value v = EvalSourceIn("=VLOOKUP(7, A1:B3, 2, TRUE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 500.0);
}

TEST(BuiltinsVLookup, ApproximateLargerThanAllPicksLast) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("one"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("five"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 1, Value::text("ten"));
  const Value v = EvalSourceIn("=VLOOKUP(999, A1:B3, 2, TRUE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ten");
}

TEST(BuiltinsVLookup, OmittedRangeLookupDefaultsToTrue) {
  // Omit the 4th arg: should behave identically to range_lookup=TRUE
  // (approximate match).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(500.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(1000.0));
  const Value v = EvalSourceIn("=VLOOKUP(7, A1:B3, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 500.0);
}

TEST(BuiltinsVLookup, ColIndexLessThanOneIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  const Value v = EvalSourceIn("=VLOOKUP(1, A1:B1, 0, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsVLookup, ColIndexBeyondColumnsIsRefError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  const Value v = EvalSourceIn("=VLOOKUP(1, A1:B1, 5, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsVLookup, LookupValueErrorPropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  const Value v = EvalSourceIn("=VLOOKUP(#DIV/0!, A1:B1, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsVLookup, SingleColumnRange) {
  // cols == 1, col_index_num = 1 -> returns the first-column cell itself.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  const Value v = EvalSourceIn("=VLOOKUP(20, A1:A3, 1, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

TEST(BuiltinsVLookup, CrossSheetTableArray) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(0, 0, Value::text("alpha"));
  wb.sheet(1).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 0, Value::text("beta"));
  wb.sheet(1).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(1).set_cell_value(2, 0, Value::text("gamma"));
  wb.sheet(1).set_cell_value(2, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=VLOOKUP(\"beta\", Data!A1:B3, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsVLookup, TwoDimensionalMixedColumns) {
  // Mixed text / numeric columns; return a text column using a numeric
  // first-column lookup.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("alpha"));
  wb.sheet(0).set_cell_value(0, 2, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("beta"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(200.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 1, Value::text("gamma"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(300.0));
  const Value v = EvalSourceIn("=VLOOKUP(2, A1:C3, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "beta");
}

TEST(BuiltinsVLookup, WrongArityIsValueError) {
  EXPECT_EQ(EvalSource("=VLOOKUP(1)").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=VLOOKUP(1, 2)").as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// HLOOKUP
// ---------------------------------------------------------------------------

TEST(BuiltinsHLookup, ExactNumericFirstRow) {
  // First row {10, 20, 30}; row_index_num=2 returns the cell directly
  // below the matched column.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(30.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("ten"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("twenty"));
  wb.sheet(0).set_cell_value(1, 2, Value::text("thirty"));
  const Value v = EvalSourceIn("=HLOOKUP(20, A1:C2, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "twenty");
}

TEST(BuiltinsHLookup, ArrayLookupValueExactPreservesShapeAndErrors) {
  const Value v = EvalSource("=HLOOKUP({20,99;#DIV/0!,10},{10,20,30;\"ten\",\"twenty\",\"thirty\"},2,FALSE)");
  ExpectArrayShape(v, 2U, 2U);
  const Value* cells = v.as_array_cells();
  ASSERT_TRUE(cells[0].is_text());
  EXPECT_EQ(cells[0].as_text(), "twenty");
  ASSERT_TRUE(cells[1].is_error());
  EXPECT_EQ(cells[1].as_error(), ErrorCode::NA);
  ASSERT_TRUE(cells[2].is_error());
  EXPECT_EQ(cells[2].as_error(), ErrorCode::Div0);
  ASSERT_TRUE(cells[3].is_text());
  EXPECT_EQ(cells[3].as_text(), "ten");
}

TEST(BuiltinsHLookup, ArrayLookupValueNestedSequenceMapsOnce) {
  const Value v = EvalSource("=HLOOKUP(SEQUENCE(1,3,25,-20),{10,20,30;\"low\",\"mid\",\"high\"},2,TRUE)");
  ExpectArrayShape(v, 1U, 3U);
  const Value* cells = v.as_array_cells();
  ASSERT_TRUE(cells[0].is_text());
  EXPECT_EQ(cells[0].as_text(), "mid");
  ASSERT_TRUE(cells[1].is_error());
  EXPECT_EQ(cells[1].as_error(), ErrorCode::NA);
  ASSERT_TRUE(cells[2].is_error());
  EXPECT_EQ(cells[2].as_error(), ErrorCode::NA);
}

TEST(BuiltinsLookup, ArrayLookupReferenceBlankPromotionForVAndHLookup) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 3, Value::number(1.0));  // D1
  wb.sheet(0).set_cell_value(1, 3, Value::number(2.0));  // D2
  wb.sheet(0).set_cell_value(2, 3, Value::number(3.0));  // D3
  wb.sheet(0).set_cell_value(0, 4, Value::number(2.0));  // E1
  // E2 intentionally remains blank.
  wb.sheet(0).set_cell_value(2, 4, Value::number(1.0));  // E3

  const Value vlookup_counta = EvalSourceIn("=COUNTA(VLOOKUP(SEQUENCE(3),D1:E3,2,FALSE))", wb, wb.sheet(0));
  ASSERT_TRUE(vlookup_counta.is_number());
  EXPECT_DOUBLE_EQ(vlookup_counta.as_number(), 3.0);
  const Value vlookup_count = EvalSourceIn("=COUNT(VLOOKUP(SEQUENCE(3),D1:E3,2,FALSE))", wb, wb.sheet(0));
  ASSERT_TRUE(vlookup_count.is_number());
  EXPECT_DOUBLE_EQ(vlookup_count.as_number(), 2.0);
  const Value vlookup_blank = EvalSourceIn("=ISBLANK(INDEX(VLOOKUP(SEQUENCE(3),D1:E3,2,FALSE),2,1))", wb, wb.sheet(0));
  ASSERT_TRUE(vlookup_blank.is_boolean());
  EXPECT_TRUE(vlookup_blank.as_boolean());
  const Value vlookup_scalar_counta = EvalSourceIn("=COUNTA(VLOOKUP(2,D1:E3,2,FALSE))", wb, wb.sheet(0));
  ASSERT_TRUE(vlookup_scalar_counta.is_number());
  EXPECT_DOUBLE_EQ(vlookup_scalar_counta.as_number(), 0.0);

  wb.sheet(0).set_cell_value(0, 6, Value::number(1.0));  // G1
  wb.sheet(0).set_cell_value(0, 7, Value::number(2.0));  // H1
  wb.sheet(0).set_cell_value(0, 8, Value::number(3.0));  // I1
  wb.sheet(0).set_cell_value(1, 6, Value::number(2.0));  // G2
  // H2 intentionally remains blank.
  wb.sheet(0).set_cell_value(1, 8, Value::number(1.0));  // I2

  const Value hlookup_counta = EvalSourceIn("=COUNTA(HLOOKUP(SEQUENCE(3),G1:I2,2,FALSE))", wb, wb.sheet(0));
  ASSERT_TRUE(hlookup_counta.is_number());
  EXPECT_DOUBLE_EQ(hlookup_counta.as_number(), 3.0);
  const Value hlookup_count = EvalSourceIn("=COUNT(HLOOKUP(SEQUENCE(3),G1:I2,2,FALSE))", wb, wb.sheet(0));
  ASSERT_TRUE(hlookup_count.is_number());
  EXPECT_DOUBLE_EQ(hlookup_count.as_number(), 2.0);
  const Value hlookup_blank = EvalSourceIn("=ISBLANK(INDEX(HLOOKUP(SEQUENCE(3),G1:I2,2,FALSE),2,1))", wb, wb.sheet(0));
  ASSERT_TRUE(hlookup_blank.is_boolean());
  EXPECT_TRUE(hlookup_blank.as_boolean());
  const Value hlookup_scalar_counta = EvalSourceIn("=COUNTA(HLOOKUP(H1,G1:I2,2,FALSE))", wb, wb.sheet(0));
  ASSERT_TRUE(hlookup_scalar_counta.is_number());
  EXPECT_DOUBLE_EQ(hlookup_scalar_counta.as_number(), 0.0);
}

TEST(BuiltinsHLookup, ExactWildcardStar) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("Banana"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("Apple"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("Apricot"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(200.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(300.0));
  const Value v = EvalSourceIn("=HLOOKUP(\"A*\", A1:C2, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 200.0);  // "Apple" is the first A* hit.
}

TEST(BuiltinsHLookup, ApproximateBetweenCellsPicksSmaller) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(500.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1000.0));
  const Value v = EvalSourceIn("=HLOOKUP(7, A1:C2, 2, TRUE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 500.0);
}

TEST(BuiltinsHLookup, ApproximateSmallerThanFirstIsNa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("five"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("ten"));
  const Value v = EvalSourceIn("=HLOOKUP(1, A1:B2, 2, TRUE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsHLookup, OmittedRangeLookupDefaultsToTrue) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(500.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1000.0));
  const Value v = EvalSourceIn("=HLOOKUP(7, A1:C2, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 500.0);
}

TEST(BuiltinsHLookup, RowIndexBeyondRowsIsRefError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  const Value v = EvalSourceIn("=HLOOKUP(1, A1:B2, 5, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsHLookup, LookupValueErrorPropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=HLOOKUP(#N/A, A1:A2, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsHLookup, CrossSheetTableArray) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(0, 0, Value::text("alpha"));
  wb.sheet(1).set_cell_value(0, 1, Value::text("beta"));
  wb.sheet(1).set_cell_value(0, 2, Value::text("gamma"));
  wb.sheet(1).set_cell_value(1, 0, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(1).set_cell_value(1, 2, Value::number(3.0));
  const Value v = EvalSourceIn("=HLOOKUP(\"beta\", Data!A1:C2, 2, FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsHLookup, WrongArityIsValueError) {
  EXPECT_EQ(EvalSource("=HLOOKUP(1)").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=HLOOKUP(1, 2, 3, TRUE, 5)").as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
