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

TEST(BuiltinsIndex, ColumnRangeRowSelects) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(40.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(50.0));
  const Value v = EvalSourceIn("=INDEX(A1:A5, 3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsIndex, RowRangeSoleIndexPicksColumn) {
  // 1-D row vector A1:E1 with two args (sole index = column).
  Workbook wb = Workbook::create();
  for (std::uint32_t c = 0; c < 5; ++c) {
    wb.sheet(0).set_cell_value(0, c, Value::number(10.0 * static_cast<double>(c + 1)));
  }
  const Value v = EvalSourceIn("=INDEX(A1:E1, 4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 40.0);
}

TEST(BuiltinsIndex, TwoDimensionalRowAndColumn) {
  Workbook wb = Workbook::create();
  // Populate A1:C3 with distinct values 1..9 row-major.
  for (std::uint32_t r = 0; r < 3; ++r) {
    for (std::uint32_t c = 0; c < 3; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 3 + c + 1)));
    }
  }
  const Value v = EvalSourceIn("=INDEX(A1:C3, 2, 3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);  // row 2, col 3 -> cell (1,2) -> 1*3+2+1 = 6
}

TEST(BuiltinsIndex, ArrayLiteralUsesSharedRangeMaterialization) {
  const Value v = EvalSource("=INDEX({1,2;3,4},2,1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsIndex, OutOfBoundsRowIsRefError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=INDEX(A1:A2, 5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndex, OutOfBoundsColumnIsRefError) {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 2; ++r) {
    for (std::uint32_t c = 0; c < 2; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(1.0));
    }
  }
  const Value v = EvalSourceIn("=INDEX(A1:B2, 1, 5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndex, NegativeIndexIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=INDEX(A1:A1, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIndex, ZeroRowOn2DRangeSpillsColumn) {
  // INDEX(A1:C3, 0, 2) spills the whole column 2 as a vertical array.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 3; ++r) {
    for (std::uint32_t c = 0; c < 3; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 3 + c + 1)));
    }
  }
  const Value v = EvalSourceIn("=INDEX(A1:C3, 0, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 2.0);  // (0,1)
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 5.0);  // (1,1)
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 8.0);  // (2,1)
}

TEST(BuiltinsIndex, ZeroColOn2DRangeSpillsRow) {
  // INDEX(A1:C3, 2, 0) spills the whole row 2 as a horizontal array. The
  // silent-divergence form INDEX(A1:C3, 2, 0) used to return only the
  // row's first cell.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 3; ++r) {
    for (std::uint32_t c = 0; c < 3; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 3 + c + 1)));
    }
  }
  const Value v = EvalSourceIn("=INDEX(A1:C3, 2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 4.0);  // (1,0)
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 5.0);  // (1,1)
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 6.0);  // (1,2)
}

TEST(BuiltinsIndex, TwoArg2DReferenceRowOnlyIsRefError) {
  // The reference form needs both indices on a 2-D area: Excel 365 gives
  // `INDEX(A1:B2,1)` #REF! (measured).
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 3; ++r) {
    for (std::uint32_t c = 0; c < 3; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 3 + c + 1)));
    }
  }
  const Value v = EvalSourceIn("=INDEX(A1:C3, 3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndex, ZeroBothOn2DRangeSpillsWholeArray) {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 2; ++r) {
    for (std::uint32_t c = 0; c < 2; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 2 + c + 1)));
    }
  }
  const Value v = EvalSourceIn("=INDEX(A1:B2, 0, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 4.0);
}

TEST(BuiltinsIndex, ZeroRowWithColumnOmittedSpillsAnArrayButNotAReference) {
  // Excel 365: `INDEX(Sheet2!A1:C2,0)` is #REF! like any row-only 2-D
  // reference, while the array form spills the whole array.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 2; ++r) {
    for (std::uint32_t c = 0; c < 2; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 2 + c + 1)));
    }
  }
  const Value two_arg = EvalSourceIn("=INDEX(A1:B2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(two_arg.is_error());
  EXPECT_EQ(two_arg.as_error(), ErrorCode::Ref);

  const Value array_literal = EvalSource("=INDEX({1,2;3,4},0)");
  ASSERT_TRUE(array_literal.is_array());
  EXPECT_EQ(array_literal.as_array_rows(), 2U);
  EXPECT_EQ(array_literal.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(array_literal.as_array_cells()[3].as_number(), 4.0);
}

TEST(BuiltinsIndex, ZeroColOutOfRangeColumnIsRefError) {
  // INDEX(A1:C3, 0, 5): whole column requested but col index out of range.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 3; ++r) {
    for (std::uint32_t c = 0; c < 3; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(1.0));
    }
  }
  const Value v = EvalSourceIn("=INDEX(A1:C3, 0, 5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndex, ColumnVectorZeroSpillsVector) {
  // 2-arg INDEX(A1:A3, 0) on a column vector spills the whole column.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  const Value v = EvalSourceIn("=INDEX(A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 30.0);
}

TEST(BuiltinsIndex, SingleCellRefAsArray) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("hello"));
  const Value hit = EvalSourceIn("=INDEX(A1, 1, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(hit.is_text());
  EXPECT_EQ(hit.as_text(), "hello");
  // Out of bounds on a single-cell Ref.
  const Value miss = EvalSourceIn("=INDEX(A1, 2, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(miss.is_error());
  EXPECT_EQ(miss.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndex, CrossSheetTwoDimensionalRange) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  for (std::uint32_t r = 0; r < 3; ++r) {
    for (std::uint32_t c = 0; c < 3; ++c) {
      wb.sheet(1).set_cell_value(r, c, Value::number(static_cast<double>((r + 1) * 10 + (c + 1))));
    }
  }
  const Value v = EvalSourceIn("=INDEX(Data!A1:C3, 3, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 32.0);
}

TEST(BuiltinsIndex, NonCoercibleIndexIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=INDEX(A1:A1, \"not a number\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIndex, WrongArityIsValueError) {
  EXPECT_EQ(EvalSource("=INDEX()").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=INDEX(1)").as_error(), ErrorCode::Value);
}

TEST(BuiltinsIndex, ArrayRowSelectorWithScalarColumn) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 4, Value::number(22.0));
  wb.sheet(0).set_cell_value(2, 3, Value::number(31.0));
  wb.sheet(0).set_cell_value(2, 4, Value::number(32.0));
  const Value v = EvalSourceIn("=INDEX(D1:E3,A1:A2,2)", wb, wb.sheet(0));
  ExpectArrayShape(v, 2U, 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 12.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 32.0);
}

TEST(BuiltinsIndex, ScalarRowWithArrayColumnSelector) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 4, Value::number(22.0));
  wb.sheet(0).set_cell_value(2, 3, Value::number(31.0));
  wb.sheet(0).set_cell_value(2, 4, Value::number(32.0));
  const Value v = EvalSourceIn("=INDEX(D1:E3,2,A1:B1)", wb, wb.sheet(0));
  ExpectArrayShape(v, 1U, 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 21.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 22.0);
}

TEST(BuiltinsIndex, ArrayRowAndColumnSelectorsBroadcastOuter) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 5, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 4, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 5, Value::number(22.0));
  wb.sheet(0).set_cell_value(2, 4, Value::number(31.0));
  wb.sheet(0).set_cell_value(2, 5, Value::number(32.0));
  const Value v = EvalSourceIn("=INDEX(E1:F3,A1:A3,B1:C1)", wb, wb.sheet(0));
  ExpectArrayShape(v, 3U, 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 12.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 21.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 22.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 31.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 32.0);
}

TEST(BuiltinsIndex, UnequalVerticalSelectorsFillNA) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 5, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 4, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 5, Value::number(22.0));
  wb.sheet(0).set_cell_value(2, 4, Value::number(31.0));
  wb.sheet(0).set_cell_value(2, 5, Value::number(32.0));
  const Value v = EvalSourceIn("=INDEX(E1:F3,A1:A3,B1:B2)", wb, wb.sheet(0));
  ExpectArrayShape(v, 3U, 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 22.0);
  ASSERT_TRUE(v.as_array_cells()[2].is_error());
  EXPECT_EQ(v.as_array_cells()[2].as_error(), ErrorCode::NA);
}

TEST(BuiltinsIndex, ArraySelectorBoundsAndErrorAreLaneLocal) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(9.0));
  wb.sheet(0).set_cell_formula(2, 0, "=NA()");
  wb.sheet(0).set_cell_value(0, 3, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 4, Value::number(22.0));
  wb.sheet(0).set_cell_value(2, 3, Value::number(31.0));
  wb.sheet(0).set_cell_value(2, 4, Value::number(32.0));
  const Value v = EvalSourceIn("=INDEX(D1:E3,A1:A3,1)", wb, wb.sheet(0));
  ExpectArrayShape(v, 3U, 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 11.0);
  ASSERT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(v.as_array_cells()[1].as_error(), ErrorCode::Ref);
  ASSERT_TRUE(v.as_array_cells()[2].is_error());
  EXPECT_EQ(v.as_array_cells()[2].as_error(), ErrorCode::NA);
}

TEST(BuiltinsIndex, ArrayRowZeroComposesColumnTile) {
  const Value v = EvalSource("=INDEX({11,12;21,22},{0;2},1)");
  ExpectArrayShape(v, 2U, 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 21.0);
}

TEST(BuiltinsIndex, ArrayColumnZeroComposesRowTile) {
  const Value v = EvalSource("=INDEX({11,12;21,22},1,{0,2})");
  ExpectArrayShape(v, 1U, 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 12.0);
}

TEST(BuiltinsIndex, ArraySelectorErrorPrecedesColumnErrorPerLane) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_formula(0, 0, "=NA()");
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 4, Value::number(22.0));
  const Value v = EvalSourceIn("=INDEX(D1:E2,A1:A2,#REF!)", wb, wb.sheet(0));
  ExpectArrayShape(v, 2U, 1U);
  ASSERT_TRUE(v.as_array_cells()[0].is_error());
  EXPECT_EQ(v.as_array_cells()[0].as_error(), ErrorCode::NA);
  ASSERT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(v.as_array_cells()[1].as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndex, ArraySelectorPreservesSourceErrorShape) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=INDEX(#REF!,A1:A2,1)", wb, wb.sheet(0));
  ExpectArrayShape(v, 2U, 1U);
  ASSERT_TRUE(v.as_array_cells()[0].is_error());
  EXPECT_EQ(v.as_array_cells()[0].as_error(), ErrorCode::Ref);
  ASSERT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(v.as_array_cells()[1].as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndex, ArraySelectorReferenceBlankCountsInCountA) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 3, Value::number(3.0));
  const Value v = EvalSourceIn("=COUNTA(INDEX(A1:A3,D1:D3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
