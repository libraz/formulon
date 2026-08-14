//
// End-to-end tests for the reference-manipulation builtins `ADDRESS`,
// `INDIRECT`, and `OFFSET`. ADDRESS is eager; INDIRECT and OFFSET are
// lazy (see `eval/reference_lazy.*`). OFFSET's multi-cell range path is
// integrated with the range-aware aggregators through
// `resolve_range_arg`, so tests covering `=SUM(OFFSET(...))` and
// `=AVERAGE(OFFSET(...))` live here alongside the scalar OFFSET cases.

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
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it through the default function registry with
// no bound workbook.
Value EvalSource(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena);
}

// Parses `src` and evaluates it against a bound workbook + current sheet.
Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  const EvalContext ctx = test::workbook_context(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// ADDRESS
// ---------------------------------------------------------------------------

TEST(BuiltinsAddress, DefaultAbsoluteA1) {
  const Value v = EvalSource("=ADDRESS(1,1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "$A$1");
}

TEST(BuiltinsAddress, AbsoluteRowRelativeColumn) {
  // abs_num = 2 -> A$1 (row absolute, column relative)
  const Value v = EvalSource("=ADDRESS(1,1,2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "A$1");
}

TEST(BuiltinsAddress, RelativeRowAbsoluteColumn) {
  // abs_num = 3 -> $A1
  const Value v = EvalSource("=ADDRESS(1,1,3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "$A1");
}

TEST(BuiltinsAddress, BothRelative) {
  // abs_num = 4 -> A1
  const Value v = EvalSource("=ADDRESS(1,1,4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "A1");
}

TEST(BuiltinsAddress, InvalidAbsNumIsValueError) {
  EXPECT_TRUE(EvalSource("=ADDRESS(1,1,0)").is_error());
  EXPECT_TRUE(EvalSource("=ADDRESS(1,1,5)").is_error());
  EXPECT_TRUE(EvalSource("=ADDRESS(1,1,-1)").is_error());
}

TEST(BuiltinsAddress, ColumnLettersBoundary) {
  // 26 -> Z
  const Value v26 = EvalSource("=ADDRESS(1,26,4)");
  ASSERT_TRUE(v26.is_text());
  EXPECT_EQ(std::string(v26.as_text()), "Z1");
  // 27 -> AA
  const Value v27 = EvalSource("=ADDRESS(1,27,4)");
  ASSERT_TRUE(v27.is_text());
  EXPECT_EQ(std::string(v27.as_text()), "AA1");
  // 702 -> ZZ
  const Value v702 = EvalSource("=ADDRESS(1,702,4)");
  ASSERT_TRUE(v702.is_text());
  EXPECT_EQ(std::string(v702.as_text()), "ZZ1");
  // 703 -> AAA
  const Value v703 = EvalSource("=ADDRESS(1,703,4)");
  ASSERT_TRUE(v703.is_text());
  EXPECT_EQ(std::string(v703.as_text()), "AAA1");
  // 16384 -> XFD (Excel's last column)
  const Value vMax = EvalSource("=ADDRESS(1,16384,4)");
  ASSERT_TRUE(vMax.is_text());
  EXPECT_EQ(std::string(vMax.as_text()), "XFD1");
}

TEST(BuiltinsAddress, RowAndColumnOutOfRange) {
  EXPECT_TRUE(EvalSource("=ADDRESS(0,1)").is_error());
  EXPECT_TRUE(EvalSource("=ADDRESS(1048577,1)").is_error());
  EXPECT_TRUE(EvalSource("=ADDRESS(1,0)").is_error());
  EXPECT_TRUE(EvalSource("=ADDRESS(1,16385)").is_error());
}

TEST(BuiltinsAddress, R1C1AbsoluteStyle) {
  // a1 = FALSE, abs_num = 1 -> R1C1
  const Value v = EvalSource("=ADDRESS(1,1,1,FALSE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "R1C1");
}

TEST(BuiltinsAddress, R1C1MixedStyles) {
  // abs_num = 2 -> R1C[1]
  const Value v2 = EvalSource("=ADDRESS(1,1,2,FALSE)");
  ASSERT_TRUE(v2.is_text());
  EXPECT_EQ(std::string(v2.as_text()), "R1C[1]");
  // abs_num = 3 -> R[1]C1
  const Value v3 = EvalSource("=ADDRESS(1,1,3,FALSE)");
  ASSERT_TRUE(v3.is_text());
  EXPECT_EQ(std::string(v3.as_text()), "R[1]C1");
  // abs_num = 4 -> R[1]C[1]
  const Value v4 = EvalSource("=ADDRESS(1,1,4,FALSE)");
  ASSERT_TRUE(v4.is_text());
  EXPECT_EQ(std::string(v4.as_text()), "R[1]C[1]");
}

TEST(BuiltinsAddress, SheetPrefixBare) {
  const Value v = EvalSource("=ADDRESS(1,1,1,TRUE,\"Sheet1\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "Sheet1!$A$1");
}

TEST(BuiltinsAddress, SheetPrefixWithSpaceQuoted) {
  const Value v = EvalSource("=ADDRESS(1,1,1,TRUE,\"Sheet 1\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "'Sheet 1'!$A$1");
}

TEST(BuiltinsAddress, SheetPrefixWithApostropheEscaped) {
  const Value v = EvalSource("=ADDRESS(1,1,4,TRUE,\"O'Brien\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "'O''Brien'!A1");
}

TEST(BuiltinsAddress, SheetPrefixLeadingDigitQuoted) {
  const Value v = EvalSource("=ADDRESS(1,1,4,TRUE,\"2020\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "'2020'!A1");
}

TEST(BuiltinsAddress, SheetPrefixEmptyStillEmitsSeparator) {
  // Excel 365 emits the trailing `!` whenever the sheet_text arg is
  // supplied, even if the name is an empty string. Matches oracle.
  const Value v = EvalSource("=ADDRESS(1,1,4,TRUE,\"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "!A1");
}

TEST(BuiltinsAddress, LargeRowAndColumn) {
  const Value v = EvalSource("=ADDRESS(100,50,4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "AX100");
}

TEST(BuiltinsAddress, AbsoluteWithSheet) {
  const Value v = EvalSource("=ADDRESS(10,5,1,TRUE,\"Data\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "Data!$E$10");
}

TEST(BuiltinsAddress, R1C1WithSheet) {
  const Value v = EvalSource("=ADDRESS(10,5,1,FALSE,\"Data\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "Data!R10C5");
}

TEST(BuiltinsAddress, ArityBelowMinIsError) {
  EXPECT_TRUE(EvalSource("=ADDRESS(1)").is_error());
}

TEST(BuiltinsAddress, ErrorArgumentPropagates) {
  EXPECT_TRUE(EvalSource("=ADDRESS(1/0,1)").is_error());
}

TEST(BuiltinsAddress, NumericCoercionOnRowCol) {
  // Fractional rows truncate toward zero (matches std::trunc).
  const Value v = EvalSource("=ADDRESS(3.7, 2.9)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "$B$3");
}

// ---------------------------------------------------------------------------
// INDIRECT
// ---------------------------------------------------------------------------

TEST(BuiltinsIndirect, SingleCellLocal) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=INDIRECT(\"A1\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

TEST(BuiltinsIndirect, SingleCellAbsolute) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(2, 1, Value::number(7.5));  // B3 = 7.5
  const Value v = EvalSourceIn("=INDIRECT(\"$B$3\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.5);
}

TEST(BuiltinsIndirect, QualifiedSheetBare) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(4, 2, Value::text("hit"));  // Data!C5
  const Value v = EvalSourceIn("=INDIRECT(\"Data!C5\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "hit");
}

TEST(BuiltinsIndirect, QualifiedSheetQuotedWithSpace) {
  Workbook wb = Workbook::create();
  wb.add_sheet("My Data");
  wb.sheet(1).set_cell_value(0, 0, Value::number(99.0));
  const Value v = EvalSourceIn("=INDIRECT(\"'My Data'!A1\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 99.0);
}

TEST(BuiltinsIndirect, EmptyTextIsRefError) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=INDIRECT(\"\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndirect, GarbageTextIsRefError) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=INDIRECT(\"not-a-ref\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndirect, UnknownSheetIsRefError) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=INDIRECT(\"NoSuch!A1\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndirect, RangeTextSpillsArray) {
  // Multi-cell range text resolves to the full rectangle as a Value::Array
  // so the result spills (and is consumable by range-aware callers).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=INDIRECT(\"A1:A2\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
}

TEST(BuiltinsIndirect, RangeTextAdmitsRectanglesPastTheConjuredArrayCeiling) {
  // Same rule as the spilled OFFSET rectangle: the array is a copy of an
  // already-admitted read, so the narrower ceiling for argument-synthesised
  // arrays (2^20 cells) must not apply to it.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(549999U, 1U, Value::number(2.0));
  const Value v = EvalSourceIn("=INDIRECT(\"A1:B550000\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << "a rectangle the read admits must not fail to materialise";
  EXPECT_EQ(v.as_array_rows(), 550000U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1099999U].as_number(), 2.0);
}

TEST(BuiltinsIndirect, RangeTextInSumAggregates) {
  // SUM(INDIRECT("A1:A3")) aggregates the multi-cell range.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  const Value v = EvalSourceIn("=SUM(INDIRECT(\"A1:A3\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 60.0);
}

TEST(BuiltinsIndirect, MultiColTableInVlookup) {
  // VLOOKUP with an INDIRECT multi-column table argument resolves the
  // table rectangle and returns the column-2 cell of the matched row.
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(1).set_cell_value(0, 1, Value::text("one"));
  wb.sheet(1).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(1).set_cell_value(1, 1, Value::text("two"));
  wb.sheet(1).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(1).set_cell_value(2, 1, Value::text("three"));
  const Value v = EvalSourceIn("=VLOOKUP(2, INDIRECT(\"Data!A1:B3\"), 2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "two");
}

TEST(BuiltinsIndirect, RangeTextCollapsingToSingleCellIsScalar) {
  // Range syntax whose rectangle collapses to one cell (A1:A1) returns
  // the scalar cell, not a 1x1 array.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=INDIRECT(\"A1:A1\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

TEST(BuiltinsIndirect, R1C1StyleNotSupportedMvp) {
  // R1C1 style is deferred and surfaces as #REF! per the MVP spec.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=INDIRECT(\"R1C1\",FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndirect, BuildFromConcat) {
  // INDIRECT("A" & "1") evaluates the concat first, then resolves.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(111.0));
  const Value v = EvalSourceIn("=INDIRECT(\"A\" & \"1\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 111.0);
}

TEST(BuiltinsIndirect, ArityMustBeOneOrTwo) {
  const Value v0 = EvalSource("=INDIRECT()");
  ASSERT_TRUE(v0.is_error());
  const Value v3 = EvalSource("=INDIRECT(\"A1\",TRUE,3)");
  ASSERT_TRUE(v3.is_error());
}

TEST(BuiltinsIndirect, ErrorTextPropagates) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=INDIRECT(1/0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// Full-column / full-row INDIRECT shapes expand against the sheet's used
// range: aggregating one via SUM sees the populated cells, and an empty
// column collapses to a blank scalar (no cells to spill). The ROW / COLUMN
// family below exercises the rectangle-aware metadata path.
TEST(BuiltinsIndirect, FullColumnTextExpandsInSum) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 3, Value::number(4.0));  // D1
  wb.sheet(0).set_cell_value(1, 3, Value::number(5.0));  // D2
  const Value v = EvalSourceIn("=SUM(INDIRECT(\"D:D\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 9.0);
}

TEST(BuiltinsIndirect, FullColumnTextEmptyResolvesToZero) {
  // An empty column has no cells to spill; the expansion collapses to a
  // blank scalar, which the top-level resolves to 0 exactly like a bare
  // reference to an empty cell (`=Z1`).
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=INDIRECT(\"D:D\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsIndirect, FullRowTextExpandsInSum) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(4, 0, Value::number(1.0));  // A5
  wb.sheet(0).set_cell_value(4, 2, Value::number(2.0));  // C5
  const Value v = EvalSourceIn("=SUM(INDIRECT(\"5:5\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsIndirectFullColRow, RowOfFullColumn) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROW(INDIRECT(\"D:D\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsIndirectFullColRow, ColumnOfFullColumn) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(INDIRECT(\"D:D\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsIndirectFullColRow, RowOfFullRow) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROW(INDIRECT(\"5:5\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsIndirectFullColRow, ColumnOfFullRow) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(INDIRECT(\"5:5\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsIndirectFullColRow, AbsoluteMixedFullColumn) {
  // `$FF:FG` -> leftmost column is FF = 162.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(INDIRECT(\"$FF:FG\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 162.0);
}

TEST(BuiltinsIndirectFullColRow, ReversedFullColumnNormalisedByMin) {
  // `C:A` should report column 1 (A) as leftmost.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(INDIRECT(\"C:A\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsIndirectFullColRow, AbsoluteFullRow) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROW(INDIRECT(\"$12:$23\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 12.0);
}

TEST(BuiltinsIndirectFullColRow, LastColumnXfd) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(INDIRECT(\"XFD:XFD\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 16384.0);
}

TEST(BuiltinsIndirectFullColRow, LowercaseColumn) {
  // `s:s` -> column 19 (S).
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(INDIRECT(\"s:s\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 19.0);
}

TEST(BuiltinsIndirectFullColRow, LeadingZeroRow) {
  // `05:05` -> row 5 (leading zeros tolerated in the row part).
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROW(INDIRECT(\"05:05\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

// ---------------------------------------------------------------------------
// OFFSET (scalar form)
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// OFFSET inside lazy aggregators (range-producing path)
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// `:` operator with OFFSET / INDIRECT call endpoints.
// Mac Excel 365 (ja-JP) accepts these and unions the resolved rectangles;
// see tests/oracle/cases/range_endpoint_calls.yaml for the goldens.
// ---------------------------------------------------------------------------

TEST(RangeEndpoint, RefToOffsetSingleCol) {
  // SUM(A1:OFFSET(A1,4,0)) — lhs literal Ref, rhs OFFSET single-cell.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  const Value v = EvalSourceIn("=SUM(A1:OFFSET(A1,4,0))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(RangeEndpoint, OffsetToOffsetSingleCol) {
  // SUM(OFFSET(A1,0,0):OFFSET(A1,4,0)) — both endpoints are OFFSET calls.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  const Value v = EvalSourceIn("=SUM(OFFSET(A1,0,0):OFFSET(A1,4,0))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(RangeEndpoint, IndirectToIndirect) {
  // SUM(INDIRECT("B2"):INDIRECT("B6")) — both endpoints INDIRECT.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(3, 1, Value::number(6.0));
  wb.sheet(0).set_cell_value(4, 1, Value::number(8.0));
  wb.sheet(0).set_cell_value(5, 1, Value::number(10.0));
  const Value v = EvalSourceIn("=SUM(INDIRECT(\"B2\"):INDIRECT(\"B6\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(RangeEndpoint, OffsetAboveRefNormalises) {
  // SUM(OFFSET(A5,-4,0):A5) — lhs OFFSET resolves to A1, rhs A5;
  // endpoint ordering is normalised so the union is A1:A5.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  const Value v = EvalSourceIn("=SUM(OFFSET(A5,-4,0):A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(RangeEndpoint, RefToMultiCellOffset) {
  // SUM(A1:OFFSET(A1,2,1,2,1)) — rhs OFFSET resolves to a 2x1
  // rectangle B3:B4. The rectangle's bottom-right corner participates
  // in the union so the unioned range is A1:B4 =
  //   1 + 2 + 3 + 4 + 10 + 20 + 30 + 40 = 110.
  // Verified against Mac Excel 365 in
  // tests/oracle/golden/range_endpoint_calls.golden.json.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
    wb.sheet(0).set_cell_value(r, 1, Value::number(static_cast<double>((r + 1) * 10)));
  }
  const Value v = EvalSourceIn("=SUM(A1:OFFSET(A1,2,1,2,1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 110.0);
}

TEST(RangeEndpoint, RefToOffsetTwoCol) {
  // SUM(A1:OFFSET(A1,4,1)) — rhs OFFSET steps to column B; spans A1:B5.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
    wb.sheet(0).set_cell_value(r, 1, Value::number(static_cast<double>((r + 1) * 10)));
  }
  const Value v = EvalSourceIn("=SUM(A1:OFFSET(A1,4,1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 165.0);
}

TEST(RangeEndpoint, RefToIndirect) {
  // SUM(A1:INDIRECT("A3")) — lhs Ref, rhs INDIRECT.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(200.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(300.0));
  const Value v = EvalSourceIn("=SUM(A1:INDIRECT(\"A3\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 600.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
