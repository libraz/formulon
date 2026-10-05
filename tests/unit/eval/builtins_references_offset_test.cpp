// Reference builtin tests grouped by reference constructor.

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>

#include "builtins_references_test_helpers.h"

namespace formulon {
namespace eval {
namespace {
using namespace references_test_helpers;

TEST(BuiltinsOffset, NegativeOffsetIntoGrid) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 2, Value::text("from"));  // C3
  const Value v = EvalSourceIn("=OFFSET(C3,-2,-2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsOffset, OffsetOutOfGridIsRef) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=OFFSET(A1,-1,0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsOffset, NegativeHeightEndsAtAnchor) {
  // Excel allows negative height: the rectangle extends upward and the
  // base cell is the bottom edge. `OFFSET(A1, 0, 0, -1, 1)` yields a
  // 1x1 rectangle at A1 (base is both top and bottom when |h|=1).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=OFFSET(A1,0,0,-1,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsOffset, NegativeHeightWalksUpTwoRows) {
  // `OFFSET(C3, 0, 0, -2, 1)` anchors at C3 and walks up one row, so the
  // rectangle spans C2:C3 and spills both cells.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 2, Value::number(42.0));  // C2
  wb.sheet(0).set_cell_value(2, 2, Value::number(99.0));  // C3
  const Value v = EvalSourceIn("=OFFSET(C3,0,0,-2,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 42.0);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[1].as_number(), 99.0);
}

TEST(BuiltinsOffset, NegativeHeightOffGridIsRef) {
  // Negative height that would walk off the top is still #REF!.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=OFFSET(A1,0,0,-2,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsOffset, ZeroHeightIsRef) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=OFFSET(A1,0,0,0,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsOffset, MultiCellSpillsRectangle) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  const Value v = EvalSourceIn("=OFFSET(A1,0,0,3,3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[4].as_number(), 4.0);
}

TEST(BuiltinsOffset, MultiCellSpillAdmitsRectanglesPastTheConjuredArrayCeiling) {
  // The spilled array is a copy of the rectangle the range reader already
  // admitted, so it is bounded like a read, not like an array a formula
  // conjures from its arguments. 2 x 550,000 sits past that narrower
  // ceiling (2^20 cells) and well inside the range-expansion bound.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(549999U, 1U, Value::number(2.0));
  const Value v = EvalSourceIn("=OFFSET(A1,0,0,550000,2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << "a rectangle the read admits must not fail to materialise";
  EXPECT_EQ(v.as_array_rows(), 550000U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1099999U].as_number(), 2.0);
}

TEST(BuiltinsOffset, MultiCellSpillMaterialisesBlankCellsAsZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=OFFSET(A1,2,2,2,-2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  for (std::uint32_t i = 0; i < 4U; ++i) {
    ASSERT_TRUE(v.as_array()->cells[i].is_number());
    EXPECT_DOUBLE_EQ(v.as_array()->cells[i].as_number(), 0.0);
  }
}

TEST(BuiltinsOffset, RangeConsumersRetainBlankProvenance) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));

  // Direct OFFSET projects the two interior source blanks to numeric zero,
  // but its range-aware consumers still receive genuine Blank cells.
  const Value count = EvalSourceIn("=COUNT(OFFSET(A1,0,0,2,2))", wb, wb.sheet(0));
  ASSERT_TRUE(count.is_number());
  EXPECT_DOUBLE_EQ(count.as_number(), 2.0);
  const Value counta = EvalSourceIn("=COUNTA(OFFSET(A1,0,0,2,2))", wb, wb.sheet(0));
  ASSERT_TRUE(counta.is_number());
  EXPECT_DOUBLE_EQ(counta.as_number(), 2.0);

  // An all-blank range has no logical values; it must not become FALSE just
  // because the direct spill surface renders its cells as zero.
  const Value all_blank_and = EvalSourceIn("=AND(OFFSET(D1,0,0,2,2))", wb, wb.sheet(0));
  ASSERT_TRUE(all_blank_and.is_error());
  EXPECT_EQ(all_blank_and.as_error(), ErrorCode::Value);
}

TEST(BuiltinsReferences, RawRangeIngressProjectsBlankAcrossReferenceForms) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(2.0));
  // A2 is intentionally untouched, and A3 closes the interior-blank range.
  sheet.set_cell_value(2, 0, Value::number(1.0));

  const std::string expressions[] = {
      "A1:A3",          "OFFSET(A1,0,0,3,1)",    "INDIRECT(\"A1:A3\")",
      "INDEX(A1:A3,0)", "CHOOSE(1,A1:A3,B1:B3)", "IF(TRUE,A1:A3,B1:B3)",
  };
  for (const std::string& expression : expressions) {
    const std::string direct_formula = "=" + expression;
    const Value direct = EvalSourceIn(direct_formula, wb, sheet);
    ASSERT_TRUE(direct.is_array()) << expression << ": " << direct.debug_to_string();
    ASSERT_EQ(direct.as_array_rows(), 3U) << expression;
    ASSERT_EQ(direct.as_array_cols(), 1U) << expression;
    ASSERT_TRUE(direct.as_array_cells()[0].is_number()) << expression;
    EXPECT_DOUBLE_EQ(direct.as_array_cells()[0].as_number(), 2.0) << expression;
    ASSERT_TRUE(direct.as_array_cells()[1].is_number()) << expression;
    EXPECT_DOUBLE_EQ(direct.as_array_cells()[1].as_number(), 0.0) << expression;
    ASSERT_TRUE(direct.as_array_cells()[2].is_number()) << expression;
    EXPECT_DOUBLE_EQ(direct.as_array_cells()[2].as_number(), 1.0) << expression;

    const Value count = EvalSourceIn("=COUNT(" + expression + ")", wb, sheet);
    ASSERT_TRUE(count.is_number()) << expression << ": " << count.debug_to_string();
    EXPECT_DOUBLE_EQ(count.as_number(), 2.0) << expression;
    const Value counta = EvalSourceIn("=COUNTA(" + expression + ")", wb, sheet);
    ASSERT_TRUE(counta.is_number()) << expression << ": " << counta.debug_to_string();
    EXPECT_DOUBLE_EQ(counta.as_number(), 2.0) << expression;
    const Value and_value = EvalSourceIn("=AND(" + expression + ")", wb, sheet);
    ASSERT_TRUE(and_value.is_boolean()) << expression << ": " << and_value.debug_to_string();
    EXPECT_TRUE(and_value.as_boolean()) << expression;
  }
}

TEST(BuiltinsOffset, BaseRangeSpillsRectangle) {
  // `OFFSET(A1:B2, 0, 0)` defaults height/width from the base -> a 2x2
  // rectangle starting at A1. Passing explicit `(1,1)` dimensions
  // confirms the scalar path.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(200.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(300.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(400.0));
  const Value v_multi = EvalSourceIn("=OFFSET(A1:B2,0,0)", wb, wb.sheet(0));
  ASSERT_TRUE(v_multi.is_array());
  ASSERT_EQ(v_multi.as_array_rows(), 2U);
  ASSERT_EQ(v_multi.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(v_multi.as_array()->cells[0].as_number(), 100.0);
  EXPECT_DOUBLE_EQ(v_multi.as_array()->cells[3].as_number(), 400.0);
  const Value v_scalar = EvalSourceIn("=OFFSET(A1:B2,0,0,1,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v_scalar.is_number());
  EXPECT_DOUBLE_EQ(v_scalar.as_number(), 100.0);
}

TEST(BuiltinsOffset, ArityBelowMinIsError) {
  EXPECT_TRUE(EvalSource("=OFFSET(A1)").is_error());
  EXPECT_TRUE(EvalSource("=OFFSET(A1,0)").is_error());
}

TEST(BuiltinsOffset, ArityAboveMaxIsError) {
  EXPECT_TRUE(EvalSource("=OFFSET(A1,0,0,1,1,1)").is_error());
}

TEST(BuiltinsOffset, NonRefBaseIsValueError) {
  // OFFSET(42, 0, 0) — base is a literal scalar, not a reference.
  const Value v = EvalSource("=OFFSET(42,0,0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsOffset, WholeColumnBaseResolvesInsteadOfValueError) {
  // OFFSET(A:A, 1, 0, 1, 1) anchors one row below A1: A2.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 0, Value::number(7.0));
  const Value v = EvalSourceIn("=OFFSET(A:A,1,0,1,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsOffset, WholeColumnBaseDefaultHeightIsFullAxis) {
  // Without an explicit height/width, OFFSET reuses the base's own shape
  // -- a whole column's declared shape spans the full row axis, so
  // ROWS(OFFSET(A:A,0,1)) (column-only shift, to B:B) reports the grid
  // height rather than #VALUE!. A *row* offset on a whole column is
  // rejected with #REF! regardless of this fix: the column already spans
  // every row, so shifting it down by even one pushes the bottom edge
  // past the grid, matching Excel.
  const Value v = EvalSource("=ROWS(OFFSET(A:A,0,1))");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), static_cast<double>(Sheet::kMaxRows));

  const Value shifted_down = EvalSource("=OFFSET(A:A,1,0)");
  ASSERT_TRUE(shifted_down.is_error());
  EXPECT_EQ(shifted_down.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsOffset, EmptyHeightOrWidthKeepsTheBaseSize) {
  // An empty height or width is the base's, not zero (Excel 365:
  // `@OFFSET(A1,0,0,,2)` is the 1x2 A1:B1, so #VALUE! off its row).
  const Value rows = EvalSource("=ROWS(OFFSET(A1:A3,0,0,,2))");
  ASSERT_TRUE(rows.is_number());
  EXPECT_DOUBLE_EQ(rows.as_number(), 3.0);
  const Value cols = EvalSource("=COLUMNS(OFFSET(A1:A3,0,0,2,))");
  ASSERT_TRUE(cols.is_number());
  EXPECT_DOUBLE_EQ(cols.as_number(), 1.0);
}

TEST(BuiltinsOffset, WholeRowBaseResolvesInsteadOfValueError) {
  // OFFSET(1:1, 0, 1, 1, 1) anchors one column right of row 1: B1.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 1, Value::number(9.0));
  const Value v = EvalSourceIn("=OFFSET(1:1,0,1,1,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 9.0);
}

TEST(BuiltinsOffset, MultiColumnWholeRangeBaseUnionsTheColumnSpan) {
  // OFFSET(A:C, 0, 0, 1, 1) keeps the full row axis and the A..C column
  // span, anchored at A1.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 2, Value::number(3.0));
  const Value v = EvalSourceIn("=OFFSET(A:C,0,2,1,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsOffset, LetBoundRangeBaseIsLookedThrough) {
  // LET(r, A1:A3, ROWS(OFFSET(r, 1, 0))) -- the LET binding's underlying
  // RangeOp shape (3 rows) must reach resolve_offset_base rather than
  // collapsing to its scalar spill anchor.
  const Value v = EvalSource("=LET(r,A1:A3,ROWS(OFFSET(r,1,0)))");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsOffset, LetBoundNameBaseWithExplicitHeightIsLookedThrough) {
  // The dynamic-named-range idiom (OFFSET(name,0,0,COUNTA(...))), scoped
  // to what `resolve_offset_base` actually looks through: a LET/LAMBDA
  // binding via `ctx.name_env()`, the same mechanism ROW / COLUMN / AREAS
  // / the conditional aggregators use. A workbook Name Manager entry is a
  // separate resolution path (`resolve_defined_name`) this finding's
  // invariant does not cover.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  const Value v = EvalSourceIn("=LET(data,A1:A2,SUM(OFFSET(data,0,0,2,1)))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsOffset, SumOfOffsetRectangle) {
  // Set up a 3x3 block with known sum.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 3; ++r) {
    for (std::uint32_t c = 0; c < 3; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 3 + c + 1)));
    }
  }
  // A1:C3 = {1,2,3; 4,5,6; 7,8,9} -> sum = 45.
  const Value v = EvalSourceIn("=SUM(OFFSET(A1,0,0,3,3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 45.0);
}

TEST(BuiltinsOffset, AverageOfOffsetRectangle) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(30.0));
  // AVERAGE(B1:C1) == 25, reached via OFFSET(A1, 0, 1, 1, 2).
  const Value v = EvalSourceIn("=AVERAGE(OFFSET(A1,0,1,1,2))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 25.0);
}

TEST(BuiltinsOffset, CountifOfOffsetRectangle) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(7.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=COUNTIF(OFFSET(A1,0,0,4,1),5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsOffset, OffsetAppliedAfterShift) {
  // The whole point of OFFSET + SUM: compute a moving sum. Shift the 3x1
  // window one column to the right (starts at B1).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(100.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(1000.0));
  const Value v = EvalSourceIn("=SUM(OFFSET(A1,0,1,1,3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1110.0);
}

TEST(BuiltinsOffset, SumOfOffsetOutOfGridPropagatesRef) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  // Trying to spill above row 1 — OFFSET produces #REF!, which SUM
  // propagates as its leftmost error.
  const Value v = EvalSourceIn("=SUM(OFFSET(A1,-1,0,2,1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsOffset, ScalarRowsCols) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(99.0));  // C2
  const Value v = EvalSourceIn("=OFFSET(A1,1,2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 99.0);
}

TEST(BuiltinsOffset, ZeroOffsetReturnsBase) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(4, 4, Value::number(42.0));  // E5
  const Value v = EvalSourceIn("=OFFSET(E5,0,0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
