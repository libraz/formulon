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

Value EvalAnchoredAt(std::string_view src, const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser parser_instance(src, parse_arena);
  parser::AstNode* root = parser_instance.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  const EvalContext ctx = test::workbook_context(wb, wb.sheet(0), state).with_formula_cell(row, col);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

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

TEST(BuiltinsIndirect, SingleCellNumericTextIsSkippedByAggregators) {
  // A single-cell INDIRECT is a reference, so SUM / PRODUCT skip its text
  // exactly as they do for SUM(A1); the text is not coerced to 7.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("7"));
  for (const char* src : {"=SUM(INDIRECT(\"A1\"))", "=PRODUCT(INDIRECT(\"A1\"))", "=SUM(A1)"}) {
    const Value v = EvalSourceIn(src, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << src << ": " << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), 0.0) << src;
  }
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

TEST(BuiltinsIndirect, R1C1AbsoluteResolvesTheCell) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=INDIRECT(\"R1C1\",FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsIndirect, R1C1RelativeResolvesAgainstTheFormulaCell) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(2, 2, Value::number(7.0));
  const Value v = EvalAnchoredAt("=INDIRECT(\"R[1]C[1]\",FALSE)", wb, 1, 1);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsIndirect, R1C1BareAxesAreTheFormulaCellItself) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 1, Value::number(5.0));
  const Value v = EvalAnchoredAt("=SUM(INDIRECT(\"RC\",FALSE))", wb, 1, 1);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsIndirect, R1C1RelativeWithoutAFormulaCellIsRef) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=INDIRECT(\"RC\",FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndirect, R1C1AbsoluteWithoutAFormulaCellStillResolves) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 2, Value::number(42.0));
  const Value v = EvalSourceIn("=INDIRECT(\"R2C3\",FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

TEST(BuiltinsIndirect, R1C1RangeSumsTheRectangle) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  const Value v = EvalSourceIn("=SUM(INDIRECT(\"R1C1:R2C2\",FALSE))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsIndirect, R1C1SingleAxisIsAWholeRow) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 5, Value::number(20.0));
  wb.sheet(0).set_cell_value(0, 0, Value::number(99.0));
  const Value v = EvalSourceIn("=SUM(INDIRECT(\"R2\",FALSE))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsIndirect, A1TextUnderR1C1FlagIsRef) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=INDIRECT(\"A1\",FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndirect, R1C1OffsetOutsideTheGridIsRef) {
  Workbook wb = Workbook::create();
  const Value v = EvalAnchoredAt("=INDIRECT(\"R[-1]C\",FALSE)", wb, 0, 0);
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndirect, R1C1MismatchedRangeEndpointsAreRef) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=INDIRECT(\"R2:R3C4\",FALSE)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
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

}  // namespace
}  // namespace eval
}  // namespace formulon
