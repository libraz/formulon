//
// End-to-end tests for the shape / geometry-inspection built-ins:
// ROWS, COLUMNS, ROW, COLUMN, and SUMPRODUCT.
//
// All five are routed through the lazy dispatch table
// (`eval_rows_lazy`, ..., `eval_sumproduct_lazy` in
// `eval/shape_ops_lazy.cpp`) because they must introspect each
// argument's AST shape. The tests here pin the three supported
// argument kinds — `Ref`, `RangeOp(Ref, Ref)`, and `ArrayLiteral` —
// plus the scalar / non-reference fallback and the error-propagation
// rules for each function.

#include <cstdint>
#include <string_view>

#include "cell.h"
#include "eval/cell_evaluator.h"
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

// ---------------------------------------------------------------------------
// Registry pins: lazy dispatch, never eager.
// ---------------------------------------------------------------------------

TEST(BuiltinsShapeRegistry, ShapeBuiltinsAreLazyOnly) {
  // None of the five impls live in the eager registry - they are routed
  // through `kLazyDispatch` in tree_walker.cpp. Pin that explicitly so a
  // future accidental `register_function(..., "ROWS", ...)` is caught.
  EXPECT_EQ(default_registry().lookup("ROWS"), nullptr);
  EXPECT_EQ(default_registry().lookup("COLUMNS"), nullptr);
  EXPECT_EQ(default_registry().lookup("ROW"), nullptr);
  EXPECT_EQ(default_registry().lookup("COLUMN"), nullptr);
  EXPECT_EQ(default_registry().lookup("SUMPRODUCT"), nullptr);
}

// ---------------------------------------------------------------------------
// ROWS / COLUMNS
// ---------------------------------------------------------------------------

TEST(BuiltinsRows, RangeRectangleReturnsRowCount) {
  Workbook wb = Workbook::create();
  // 3x2 rectangle across A1:B3.
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(6.0));
  const Value v = EvalSourceIn("=ROWS(A1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsColumns, RangeRectangleReturnsColumnCount) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  const Value v = EvalSourceIn("=COLUMNS(A1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsRows, SingleCellRefIsOne) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=ROWS(A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsColumns, SingleCellRefIsOne) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=COLUMNS(A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// Whole-axis references report the rectangle they declare, not the target
// sheet's populated extent. `expand_range` clamps the unbounded axis so
// enumerating the values stays affordable; that clamp is not a shape.
// ---------------------------------------------------------------------------

TEST(BuiltinsRows, WholeColumnIsGridHeightOnEmptySheet) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROWS(A:A)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), static_cast<double>(Sheet::kMaxRows));
}

TEST(BuiltinsRows, WholeColumnIsGridHeightOnPopulatedSheet) {
  // The same answer with content in the column: the declared rectangle does
  // not depend on what the sheet holds.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=ROWS(A:A)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), static_cast<double>(Sheet::kMaxRows));
}

TEST(BuiltinsColumns, WholeColumnSpanIsSpanWidth) {
  Workbook wb = Workbook::create();
  const Value empty = EvalSourceIn("=COLUMNS(A:C)", wb, wb.sheet(0));
  ASSERT_TRUE(empty.is_number());
  EXPECT_DOUBLE_EQ(empty.as_number(), 3.0);

  wb.sheet(0).set_cell_value(0, 1, Value::number(7.0));
  const Value populated = EvalSourceIn("=COLUMNS(A:C)", wb, wb.sheet(0));
  ASSERT_TRUE(populated.is_number());
  EXPECT_DOUBLE_EQ(populated.as_number(), 3.0);
}

TEST(BuiltinsRows, WholeColumnSpanIsGridHeight) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROWS(A:C)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), static_cast<double>(Sheet::kMaxRows));
}

TEST(BuiltinsColumns, WholeRowIsGridWidth) {
  Workbook wb = Workbook::create();
  const Value empty = EvalSourceIn("=COLUMNS(1:1)", wb, wb.sheet(0));
  ASSERT_TRUE(empty.is_number());
  EXPECT_DOUBLE_EQ(empty.as_number(), static_cast<double>(Sheet::kMaxCols));

  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value populated = EvalSourceIn("=COLUMNS(1:1)", wb, wb.sheet(0));
  ASSERT_TRUE(populated.is_number());
  EXPECT_DOUBLE_EQ(populated.as_number(), static_cast<double>(Sheet::kMaxCols));
}

TEST(BuiltinsRows, WholeRowSpanIsSpanHeight) {
  Workbook wb = Workbook::create();
  const Value empty = EvalSourceIn("=ROWS(1:3)", wb, wb.sheet(0));
  ASSERT_TRUE(empty.is_number());
  EXPECT_DOUBLE_EQ(empty.as_number(), 3.0);

  wb.sheet(0).set_cell_value(1, 0, Value::number(7.0));
  const Value populated = EvalSourceIn("=ROWS(1:3)", wb, wb.sheet(0));
  ASSERT_TRUE(populated.is_number());
  EXPECT_DOUBLE_EQ(populated.as_number(), 3.0);
}

TEST(BuiltinsRows, WholeAxisShapeIsAtLeastOne) {
  // Excel never answers 0 for a valid full-axis reference; the used-range
  // clamp used to make an empty sheet do exactly that.
  Workbook wb = Workbook::create();
  for (std::string_view src : {"=ROWS(A:A)", "=COLUMNS(A:A)", "=ROWS(1:1)", "=COLUMNS(1:1)", "=ROWS(A:C)",
                               "=COLUMNS(A:C)", "=ROWS(1:3)", "=COLUMNS(1:3)"}) {
    const Value v = EvalSourceIn(src, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << src;
    EXPECT_GE(v.as_number(), 1.0) << src;
  }
}

TEST(BuiltinsRows, UnknownSheetQualifierIsRef) {
  // Measuring a rectangle still requires the reference to resolve.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROWS(NoSuchSheet!A1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsRows, ArrayLiteralColumnShape) {
  // `{1;2;3}` is a 3x1 column literal (semicolons separate rows).
  const Value v = EvalSource("=ROWS({1;2;3})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsRows, ArrayLiteralRowIsOne) {
  // `{1,2,3}` is a 1x3 row literal (commas separate columns).
  const Value v = EvalSource("=ROWS({1,2,3})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsColumns, ArrayLiteralRowShape) {
  const Value v = EvalSource("=COLUMNS({1,2,3})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsRows, ScalarFallbackIsOne) {
  // Bare number expression: treated as 1x1 per Excel's scalar treatment.
  const Value v = EvalSource("=ROWS(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsRows, ErrorSubtreePropagates) {
  // Scalar argument that evaluates to an error must propagate that error
  // unchanged instead of silently becoming 1.
  const Value v = EvalSource("=ROWS(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsRows, ZeroArgIsArityViolation) {
  const Value v = EvalSource("=ROWS()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// ROW / COLUMN (1-arg form only; zero-arity would need current-cell ctx)
// ---------------------------------------------------------------------------

TEST(BuiltinsRow, SingleRefIsOneBasedRow) {
  // C5 -> 0-based (row=4, col=2) -> 1-based row=5.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROW(C5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsColumn, SingleRefIsOneBasedCol) {
  // C5 -> col=2 -> 1-based col=3.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(C5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsRow, RangeOpReturnsRowVector) {
  // B2:D4 -> {2;3;4}.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROW(B2:D4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[2].as_number(), 4.0);
}

TEST(BuiltinsColumn, RangeOpReturnsColumnVector) {
  // B2:D4 -> {2,3,4}.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(B2:D4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[2].as_number(), 4.0);
}

// ---------------------------------------------------------------------------
// ROW / COLUMN over a whole-axis reference project the declared rectangle,
// the same one ROWS / COLUMNS measure. A full-axis endpoint names a
// coordinate only on its bounded axis, so the projection cannot be read off
// the endpoint's raw row / col fields: `A:A` leaves `row` at its structure
// default, which is why the whole grid height used to collapse to 1.
// ---------------------------------------------------------------------------

TEST(BuiltinsRow, WholeColumnProjectsEveryGridRow) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROW(A:A)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_array_rows(), Sheet::kMaxRows);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[Sheet::kMaxRows - 1U].as_number(), static_cast<double>(Sheet::kMaxRows));
}

TEST(BuiltinsRow, WholeColumnSpanProjectsEveryGridRow) {
  // A multi-column band is still one column of row indices: only the
  // requested axis is projected.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROW(A:C)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_array_rows(), Sheet::kMaxRows);
  EXPECT_EQ(v.as_array_cols(), 1U);
}

TEST(BuiltinsColumn, WholeRowProjectsEveryGridColumn) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(1:1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), Sheet::kMaxCols);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[Sheet::kMaxCols - 1U].as_number(), static_cast<double>(Sheet::kMaxCols));
}

TEST(BuiltinsColumn, WholeRowSpanProjectsEveryGridColumn) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(1:3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), Sheet::kMaxCols);
}

TEST(BuiltinsRow, WholeAxisOffAxisProjectionStaysScalar) {
  // The axis a whole-axis reference bounds still projects to one index:
  // `ROW(1:1)` is row 1, `COLUMN(A:A)` is column 1. These are the two
  // spellings that were already right, and they must stay scalar rather
  // than becoming one-element arrays.
  Workbook wb = Workbook::create();
  const Value row = EvalSourceIn("=ROW(1:1)", wb, wb.sheet(0));
  ASSERT_TRUE(row.is_number()) << "kind=" << static_cast<int>(row.kind());
  EXPECT_DOUBLE_EQ(row.as_number(), 1.0);

  const Value column = EvalSourceIn("=COLUMN(A:A)", wb, wb.sheet(0));
  ASSERT_TRUE(column.is_number()) << "kind=" << static_cast<int>(column.kind());
  EXPECT_DOUBLE_EQ(column.as_number(), 1.0);
}

TEST(BuiltinsRow, WholeAxisSpanOffAxisProjectsItsSpan) {
  // `ROW(1:3)` and `COLUMN(A:C)` project the bounded axis, so they keep
  // reporting the span rather than the grid.
  Workbook wb = Workbook::create();
  const Value row = EvalSourceIn("=ROW(1:3)", wb, wb.sheet(0));
  ASSERT_TRUE(row.is_array()) << "kind=" << static_cast<int>(row.kind());
  EXPECT_EQ(row.as_array_rows(), 3U);
  EXPECT_EQ(row.as_array_cols(), 1U);

  const Value column = EvalSourceIn("=COLUMN(A:C)", wb, wb.sheet(0));
  ASSERT_TRUE(column.is_array()) << "kind=" << static_cast<int>(column.kind());
  EXPECT_EQ(column.as_array_rows(), 1U);
  EXPECT_EQ(column.as_array_cols(), 3U);
}

TEST(BuiltinsRow, WholeAxisProjectionAggregates) {
  // The aggregator seam expands the same rectangle: these are the closed
  // forms of 1..1048576 and 1..16384, so a projection that silently
  // clamped to the populated extent could not produce them.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));

  const Value rows = EvalSourceIn("=SUM(ROW(A:A))", wb, wb.sheet(0));
  ASSERT_TRUE(rows.is_number()) << "kind=" << static_cast<int>(rows.kind());
  const double n_rows = static_cast<double>(Sheet::kMaxRows);
  EXPECT_DOUBLE_EQ(rows.as_number(), n_rows * (n_rows + 1.0) / 2.0);

  const Value cols = EvalSourceIn("=SUM(COLUMN(1:1))", wb, wb.sheet(0));
  ASSERT_TRUE(cols.is_number()) << "kind=" << static_cast<int>(cols.kind());
  const double n_cols = static_cast<double>(Sheet::kMaxCols);
  EXPECT_DOUBLE_EQ(cols.as_number(), n_cols * (n_cols + 1.0) / 2.0);
}

TEST(BuiltinsRow, WholeColumnSpillsFromTheTopRowOnly) {
  // A grid-height projection fits the sheet only when it is anchored at
  // row 1; anywhere else it runs off the bottom and Excel reports #SPILL!.
  // The committed region is a descriptor, not a million materialised
  // cells, so the sheet-side cost does not follow the projection's size.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Cell anchored;
  anchored.formula_text = "=ROW(A:A)";
  Arena arena;
  const Value at_top = evaluate_cell_for_recalc(wb, sheet, anchored, 0U, 0U, default_registry(), arena);
  ASSERT_TRUE(at_top.is_number()) << "kind=" << static_cast<int>(at_top.kind());
  EXPECT_DOUBLE_EQ(at_top.as_number(), 1.0);
  ASSERT_NE(sheet.spill_region_at_anchor(0U, 0U), nullptr);
  EXPECT_EQ(sheet.cell_at(Sheet::kMaxRows - 1U, 0U), nullptr) << "the region is a descriptor, not stored cells";

  const Value mid = EvalSourceIn("=A500000", wb, sheet);
  ASSERT_TRUE(mid.is_number()) << "kind=" << static_cast<int>(mid.kind());
  EXPECT_DOUBLE_EQ(mid.as_number(), 500000.0) << "read back out of the committed region";

  Workbook lower = Workbook::create();
  Sheet& lower_sheet = lower.sheet(0);
  Cell below;
  below.formula_text = "=ROW(A:A)";
  Arena lower_arena;
  const Value off_top = evaluate_cell_for_recalc(lower, lower_sheet, below, 5U, 3U, default_registry(), lower_arena);
  ASSERT_TRUE(off_top.is_error()) << "kind=" << static_cast<int>(off_top.kind());
  EXPECT_EQ(off_top.as_error(), ErrorCode::Spill);
}

TEST(BuiltinsRow, ArrayLiteralIsValue) {
  // Array literal is not a reference; Excel's `=ROW({1,2,3})` is #VALUE!.
  const Value v = EvalSource("=ROW({1,2,3})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsRow, ScalarLiteralIsValue) {
  // A bare number is not a reference.
  const Value v = EvalSource("=ROW(42)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsRow, ErrorSubtreePropagates) {
  const Value v = EvalSource("=ROW(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsRow, ZeroArgWithoutAnchorIsValue) {
  // Without a formula-cell anchor the 0-arg form has no meaningful answer,
  // so surface #VALUE! rather than inventing a row number.
  const Value v = EvalSource("=ROW()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsRow, ZeroArgReturnsAnchorRowOneBased) {
  // The oracle harness anchors each formula at its own cell so ROW() / COLUMN()
  // can report the containing cell's coordinates. Mirror that here by
  // constructing an EvalContext with the anchor explicitly.
  Workbook wb = Workbook::create();
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=ROW()", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EvalState state;
  const EvalContext ctx = EvalContext(wb, wb.sheet(0), state).with_formula_cell(4U, 2U);  // C5
  const Value v = evaluate(*root, eval_arena, default_registry(), ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsColumn, ZeroArgReturnsAnchorColumnOneBased) {
  Workbook wb = Workbook::create();
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=COLUMN()", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EvalState state;
  const EvalContext ctx = EvalContext(wb, wb.sheet(0), state).with_formula_cell(4U, 2U);  // C5
  const Value v = evaluate(*root, eval_arena, default_registry(), ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// SUMPRODUCT
// ---------------------------------------------------------------------------

TEST(BuiltinsSumproduct, TwoParallelRanges) {
  // [1,2,3] * [4,5,6] => 1*4 + 2*5 + 3*6 = 32.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(6.0));
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A3,B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 32.0);
}

TEST(BuiltinsSumproduct, SingleRangeIsPlainSum) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsSumproduct, ShapeMismatchIsValue) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(5.0));
  // A1:A2 is 2x1; B1:B3 is 3x1 -> mismatched rectangles.
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A2,B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSumproduct, TextCellTreatedAsZero) {
  // [1, "x", 3] * [10, 20, 30] = 1*10 + 0*20 + 3*30 = 100. SUMPRODUCT
  // specifically does NOT coerce non-numeric cells - they contribute 0.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A3,B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0);
}

TEST(BuiltinsSumproduct, ErrorInRangePropagates) {
  // An error cell in the first (leftmost) range wins.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_formula(1, 0, "=1/0");
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A3,B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsSumproduct, ArrayLiteralColumns) {
  // {1;2;3} * {4;5;6} = 32, identical to the range variant.
  const Value v = EvalSource("=SUMPRODUCT({1;2;3},{4;5;6})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 32.0);
}

TEST(BuiltinsSumproduct, ThreeArrays) {
  // {1;2} * {3;4} * {5;6} = 1*3*5 + 2*4*6 = 15 + 48 = 63.
  const Value v = EvalSource("=SUMPRODUCT({1;2},{3;4},{5;6})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 63.0);
}

TEST(BuiltinsSumproduct, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=SUMPRODUCT()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSumproduct, ScalarArgsMultiply) {
  // All 1x1 arguments: 2 * 3 * 4 = 24.
  const Value v = EvalSource("=SUMPRODUCT(2,3,4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 24.0);
}

// ---------------------------------------------------------------------------
// Chained `:` — the operator composes, so the shape is the bounding box.
// ---------------------------------------------------------------------------
//
// Excel accepts a `:` chain and stores it verbatim, and the rectangle it
// names is the bounding box of every endpoint: `=COLUMNS(A:C:E:G)` is 7,
// measured on Excel 365 (Mac, ja-JP, 16.112.1). The two-endpoint shape is
// not a special case of the grammar, so a derivation that stops there
// answers `#VALUE!` for a spelling Excel evaluates.

TEST(BuiltinsShapeChainedRange, WholeColumnChainSpansEveryEndpoint) {
  Workbook wb = Workbook::create();
  const Value cols = EvalSourceIn("=COLUMNS(A:C:E:G)", wb, wb.sheet(0));
  ASSERT_TRUE(cols.is_number()) << "expected a count, got an error";
  EXPECT_DOUBLE_EQ(cols.as_number(), 7.0);

  // Out-of-order endpoints bound the same rectangle; the chain is not
  // required to run left to right.
  const Value reversed = EvalSourceIn("=COLUMNS(G:E:C:A)", wb, wb.sheet(0));
  ASSERT_TRUE(reversed.is_number());
  EXPECT_DOUBLE_EQ(reversed.as_number(), 7.0);

  const Value longer = EvalSourceIn("=COLUMNS(A:C:E:G:I:K)", wb, wb.sheet(0));
  ASSERT_TRUE(longer.is_number());
  EXPECT_DOUBLE_EQ(longer.as_number(), 11.0);
}

TEST(BuiltinsShapeChainedRange, RepeatedEndpointsDoNotChangeTheRectangle) {
  // `H:H:E:E` is the four-endpoint spelling of `H:E`, and the six-endpoint
  // `h:h:h:h:e:e` names the same four columns. Each was answering
  // `#VALUE!` past the depth the two-endpoint derivation reached, so the
  // answer used to depend on how many times a column was written.
  Workbook wb = Workbook::create();
  for (const char* src : {"=COLUMNS(H:E)", "=COLUMNS(H:H:E:E)", "=COLUMNS(h:h:h:h:e:e)"}) {
    const Value cols = EvalSourceIn(src, wb, wb.sheet(0));
    ASSERT_TRUE(cols.is_number()) << src;
    EXPECT_DOUBLE_EQ(cols.as_number(), 4.0) << src;
  }
}

TEST(BuiltinsShapeChainedRange, WholeRowChainSpansEveryEndpoint) {
  Workbook wb = Workbook::create();
  const Value rows = EvalSourceIn("=ROWS(1:3:5:7)", wb, wb.sheet(0));
  ASSERT_TRUE(rows.is_number());
  EXPECT_DOUBLE_EQ(rows.as_number(), 7.0);
  // The unbounded axis stays whole-grid, exactly as for the pair form.
  const Value cols = EvalSourceIn("=COLUMNS(1:3:5:7)", wb, wb.sheet(0));
  ASSERT_TRUE(cols.is_number());
  EXPECT_DOUBLE_EQ(cols.as_number(), static_cast<double>(Sheet::kMaxCols));
}

TEST(BuiltinsShapeChainedRange, BoundedChainTakesEachAxisFromADifferentEndpoint) {
  // Neither corner of `A1:C3:E5`'s bounding box is one of the endpoints on
  // both axes at once, which is why the chain cannot be reduced by picking
  // one endpoint per side.
  Workbook wb = Workbook::create();
  const Value cols = EvalSourceIn("=COLUMNS(A1:C3:E5)", wb, wb.sheet(0));
  ASSERT_TRUE(cols.is_number());
  EXPECT_DOUBLE_EQ(cols.as_number(), 5.0);
  const Value rows = EvalSourceIn("=ROWS(A1:C3:E5)", wb, wb.sheet(0));
  ASSERT_TRUE(rows.is_number());
  EXPECT_DOUBLE_EQ(rows.as_number(), 5.0);

  const Value out_of_order = EvalSourceIn("=ROWS(C1:A3:B2)", wb, wb.sheet(0));
  ASSERT_TRUE(out_of_order.is_number());
  EXPECT_DOUBLE_EQ(out_of_order.as_number(), 3.0);
}

TEST(BuiltinsShapeChainedRange, MixedAxisChainBoundsTheWholeGrid) {
  // The endpoints do not have to bound the same axis. Composing a
  // whole-column span with a whole-row span leaves neither axis bounded,
  // so the rectangle is the entire grid -- not an error. Measured on
  // Excel 365 (Mac, ja-JP, 16.112.1); `shape.yaml` carries the goldens.
  Workbook wb = Workbook::create();
  const Value cols = EvalSourceIn("=COLUMNS(A:C:1:3)", wb, wb.sheet(0));
  ASSERT_TRUE(cols.is_number()) << "a mixed-axis chain is not an error in Excel";
  EXPECT_DOUBLE_EQ(cols.as_number(), static_cast<double>(Sheet::kMaxCols));
  const Value rows = EvalSourceIn("=ROWS(A:C:1:3)", wb, wb.sheet(0));
  ASSERT_TRUE(rows.is_number());
  EXPECT_DOUBLE_EQ(rows.as_number(), static_cast<double>(Sheet::kMaxRows));
}

TEST(BuiltinsShapeChainedRange, WholeAxisChainedWithACellExtendsOnlyTheBoundedAxis) {
  // `A:C:E1` keeps every row and widens the columns to A..E; the bounded
  // endpoint cannot narrow an axis the span already covers.
  Workbook wb = Workbook::create();
  const Value col_span_cols = EvalSourceIn("=COLUMNS(A:C:E1)", wb, wb.sheet(0));
  ASSERT_TRUE(col_span_cols.is_number());
  EXPECT_DOUBLE_EQ(col_span_cols.as_number(), 5.0);
  const Value col_span_rows = EvalSourceIn("=ROWS(A:C:E1)", wb, wb.sheet(0));
  ASSERT_TRUE(col_span_rows.is_number());
  EXPECT_DOUBLE_EQ(col_span_rows.as_number(), static_cast<double>(Sheet::kMaxRows));

  const Value row_span_rows = EvalSourceIn("=ROWS(1:3:E5)", wb, wb.sheet(0));
  ASSERT_TRUE(row_span_rows.is_number());
  EXPECT_DOUBLE_EQ(row_span_rows.as_number(), 5.0);
  const Value row_span_cols = EvalSourceIn("=COLUMNS(1:3:E5)", wb, wb.sheet(0));
  ASSERT_TRUE(row_span_cols.is_number());
  EXPECT_DOUBLE_EQ(row_span_cols.as_number(), static_cast<double>(Sheet::kMaxCols));
}

TEST(BuiltinsShapeChainedRange, ChainedRangeAggregatesOverTheBoundingBox) {
  // The shape functions are not a separate path: a plain aggregate reads
  // the same rectangle, so the chain has to reach range expansion too.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1.0));  // A1
  sheet.set_cell_value(0, 2, Value::number(2.0));  // C1
  sheet.set_cell_value(0, 4, Value::number(4.0));  // E1
  sheet.set_cell_value(0, 6, Value::number(8.0));  // G1
  const Value sum = EvalSourceIn("=SUM(A:C:E:G)", wb, sheet);
  ASSERT_TRUE(sum.is_number()) << "chained range did not reach range expansion";
  EXPECT_DOUBLE_EQ(sum.as_number(), 15.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
