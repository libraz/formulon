//
// End-to-end tests for SUBTOTAL — the multi-mode aggregator that dispatches
// on a numeric function code (1..11 / 101..111). Tests pin:
//   * Each of the 11 modes returns the right scalar.
//   * The 100+ "ignore-hidden" codes drop range cells sitting on a hidden
//     row, and agree with 1..11 when the range holds no hidden row.
//   * Code 3 (COUNTA) sees Bool / Text / Error cells inside a range and
//     counts them, while every other code skips non-numeric range cells.
//   * Out-of-range or non-finite codes surface `#VALUE!`.
//   * Truncating dispatch: 9.7 -> SUM, 109.4 -> SUM (matches Excel).

#include <cmath>
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

// Seeds A1:A5 with 10, 20, 30, 40, 50 and returns a workbook bound to that
// sheet so the parameterized tests below can assert against a known column.
Workbook MakeNumericRange() {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(40.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(50.0));
  return wb;
}

// Marks `row` (0-based) hidden on `sheet`, the same shape the OOXML reader
// and `fm_sheet_set_row_hidden` produce.
void HideRow(Sheet& sheet, std::uint32_t row) {
  RowLayout layout;
  layout.row = row;
  layout.hidden = true;
  sheet.mutable_layout().row_overrides.push_back(layout);
}

// ---------------------------------------------------------------------------
// Registry
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotalRegistry, Registered) {
  const FunctionDef* def = default_registry().lookup("SUBTOTAL");
  ASSERT_NE(def, nullptr);
  EXPECT_TRUE(def->accepts_ranges);
  EXPECT_FALSE(def->propagate_errors);  // COUNTA must see error cells.
  EXPECT_FALSE(def->range_filter_numeric_only);
  EXPECT_EQ(def->min_arity, 2u);
}

// ---------------------------------------------------------------------------
// Mode dispatch
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, Code1IsAverage) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(1,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsSubtotal, Code2IsCount) {
  Workbook wb = MakeNumericRange();
  // Add a non-numeric cell to A6 to confirm COUNT skips it.
  wb.sheet(0).set_cell_value(5, 0, Value::text("text"));
  const Value v = EvalSourceIn("=SUBTOTAL(2,A1:A6)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsSubtotal, Code3IsCountA) {
  Workbook wb = MakeNumericRange();
  // Mix of types: 5 numbers + 1 text + 1 bool + 1 blank should yield 7.
  wb.sheet(0).set_cell_value(5, 0, Value::text("text"));
  wb.sheet(0).set_cell_value(6, 0, Value::boolean(true));
  // Row 8 (index 7) intentionally left blank; COUNTA must skip it.
  const Value v = EvalSourceIn("=SUBTOTAL(3,A1:A8)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsSubtotal, Code4IsMax) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(4,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
}

TEST(BuiltinsSubtotal, Code5IsMin) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(5,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsSubtotal, Code6IsProduct) {
  // Use a smaller range to keep the product representable as an exact int.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(4.0));
  const Value v = EvalSourceIn("=SUBTOTAL(6,A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 24.0);
}

TEST(BuiltinsSubtotal, Code9IsSum) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(9,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
}

// VAR / VARP / STDEV / STDEVP — small symmetric range so the expected values
// are easy to verify by hand. Sample variance of {1, 2, 3, 4, 5} = 2.5;
// population variance = 2.0.
TEST(BuiltinsSubtotal, Code10IsVarSample) {
  Workbook wb = Workbook::create();
  for (uint32_t i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>(i + 1)));
  }
  const Value v = EvalSourceIn("=SUBTOTAL(10,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.5);
}

TEST(BuiltinsSubtotal, Code11IsVarPopulation) {
  Workbook wb = Workbook::create();
  for (uint32_t i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>(i + 1)));
  }
  const Value v = EvalSourceIn("=SUBTOTAL(11,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsSubtotal, Code7IsStdevSample) {
  Workbook wb = Workbook::create();
  for (uint32_t i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>(i + 1)));
  }
  const Value v = EvalSourceIn("=SUBTOTAL(7,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), std::sqrt(2.5));
}

TEST(BuiltinsSubtotal, Code8IsStdevPopulation) {
  Workbook wb = Workbook::create();
  for (uint32_t i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>(i + 1)));
  }
  const Value v = EvalSourceIn("=SUBTOTAL(8,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), std::sqrt(2.0));
}

// ---------------------------------------------------------------------------
// 100+ ignore-hidden codes agree with 1..11 when nothing is hidden.
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, Code109EqualsCode9) {
  Workbook wb = MakeNumericRange();
  const Value v9 = EvalSourceIn("=SUBTOTAL(9,A1:A5)", wb, wb.sheet(0));
  const Value v109 = EvalSourceIn("=SUBTOTAL(109,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v9.is_number());
  ASSERT_TRUE(v109.is_number());
  EXPECT_DOUBLE_EQ(v9.as_number(), v109.as_number());
}

TEST(BuiltinsSubtotal, Code103EqualsCode3) {
  Workbook wb = MakeNumericRange();
  wb.sheet(0).set_cell_value(5, 0, Value::text("foo"));
  const Value v3 = EvalSourceIn("=SUBTOTAL(3,A1:A6)", wb, wb.sheet(0));
  const Value v103 = EvalSourceIn("=SUBTOTAL(103,A1:A6)", wb, wb.sheet(0));
  ASSERT_TRUE(v3.is_number());
  ASSERT_TRUE(v103.is_number());
  EXPECT_DOUBLE_EQ(v3.as_number(), v103.as_number());
}

// ---------------------------------------------------------------------------
// 100+ codes exclude cells sitting on a hidden row; 1..11 include them.
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, Code109ExcludesHiddenRows) {
  Workbook wb = MakeNumericRange();
  HideRow(wb.sheet(0), 1);  // A2 = 20
  HideRow(wb.sheet(0), 3);  // A4 = 40
  const Value v9 = EvalSourceIn("=SUBTOTAL(9,A1:A5)", wb, wb.sheet(0));
  const Value v109 = EvalSourceIn("=SUBTOTAL(109,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v9.is_number());
  ASSERT_TRUE(v109.is_number());
  EXPECT_DOUBLE_EQ(v9.as_number(), 150.0);
  EXPECT_DOUBLE_EQ(v109.as_number(), 90.0);
  EXPECT_NE(v9.as_number(), v109.as_number());
}

TEST(BuiltinsSubtotal, Code101AveragesVisibleRowsOnly) {
  Workbook wb = MakeNumericRange();
  HideRow(wb.sheet(0), 0);  // A1 = 10
  const Value v = EvalSourceIn("=SUBTOTAL(101,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 35.0);  // (20+30+40+50)/4
}

TEST(BuiltinsSubtotal, Code104MaxSkipsHiddenExtreme) {
  Workbook wb = MakeNumericRange();
  HideRow(wb.sheet(0), 4);  // A5 = 50, the maximum
  const Value v = EvalSourceIn("=SUBTOTAL(104,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 40.0);
}

TEST(BuiltinsSubtotal, Code103CountAExcludesHiddenRows) {
  Workbook wb = MakeNumericRange();
  wb.sheet(0).set_cell_value(5, 0, Value::text("foo"));
  HideRow(wb.sheet(0), 5);
  const Value v3 = EvalSourceIn("=SUBTOTAL(3,A1:A6)", wb, wb.sheet(0));
  const Value v103 = EvalSourceIn("=SUBTOTAL(103,A1:A6)", wb, wb.sheet(0));
  ASSERT_TRUE(v3.is_number());
  ASSERT_TRUE(v103.is_number());
  EXPECT_DOUBLE_EQ(v3.as_number(), 6.0);
  EXPECT_DOUBLE_EQ(v103.as_number(), 5.0);
}

TEST(BuiltinsSubtotal, HiddenErrorCellNoLongerPropagates) {
  // The hidden-row filter runs before the mode dispatch, so an error on a
  // hidden row cannot short-circuit the numeric branch.
  Workbook wb = MakeNumericRange();
  wb.sheet(0).set_cell_value(2, 0, Value::error(ErrorCode::NA));
  HideRow(wb.sheet(0), 2);
  const Value v9 = EvalSourceIn("=SUBTOTAL(9,A1:A5)", wb, wb.sheet(0));
  const Value v109 = EvalSourceIn("=SUBTOTAL(109,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v9.is_error());
  EXPECT_EQ(v9.as_error(), ErrorCode::NA);
  ASSERT_TRUE(v109.is_number());
  EXPECT_DOUBLE_EQ(v109.as_number(), 120.0);
}

TEST(BuiltinsSubtotal, HiddenRowOutsideRangeDoesNotShiftFilter) {
  // The flag is matched against the range's own top row, so a hidden row
  // below the rectangle must not drop the cell at the same offset.
  Workbook wb = MakeNumericRange();
  HideRow(wb.sheet(0), 8);
  const Value v = EvalSourceIn("=SUBTOTAL(109,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
}

TEST(BuiltinsSubtotal, HiddenRowHonoredForRangeNotStartingAtRowOne) {
  Workbook wb = MakeNumericRange();
  HideRow(wb.sheet(0), 3);  // A4 = 40
  const Value v = EvalSourceIn("=SUBTOTAL(109,A3:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 80.0);  // 30 + 50
}

TEST(BuiltinsSubtotal, HiddenRowHonoredForWholeColumnReference) {
  // A whole-column reference keeps its origin at row 0, so the offsets line
  // up with the sheet's own row numbering.
  Workbook wb = MakeNumericRange();
  HideRow(wb.sheet(0), 1);  // A2 = 20
  const Value v = EvalSourceIn("=SUBTOTAL(109,A:A)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 130.0);
}

TEST(BuiltinsSubtotal, HiddenRowHonoredForTwoDimensionalRange) {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 3; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
    wb.sheet(0).set_cell_value(r, 1, Value::number(static_cast<double>(10 * (r + 1))));
  }
  HideRow(wb.sheet(0), 1);  // drops both cells of the middle row
  const Value v = EvalSourceIn("=SUBTOTAL(109,A1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 44.0);  // 1 + 10 + 3 + 30
}

TEST(BuiltinsSubtotal, HiddenRowHonoredAcrossSheetQualifier) {
  Workbook wb = Workbook::create();
  Sheet& data = wb.sheet(wb.add_sheet("Data"));
  for (std::uint32_t r = 0; r < 3; ++r) {
    data.set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  HideRow(data, 0);
  const Value v = EvalSourceIn("=SUBTOTAL(109,Data!A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);  // 2 + 3
}

TEST(BuiltinsSubtotal, HiddenRowHonoredWhenOnlyRightEndpointIsQualified) {
  // `A1:Data!A3` is the mirrored qualifier shape `expand_range` also
  // accepts; the visibility lookup must read the same sheet it does.
  Workbook wb = Workbook::create();
  Sheet& data = wb.sheet(wb.add_sheet("Data"));
  for (std::uint32_t r = 0; r < 3; ++r) {
    data.set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
    wb.sheet(0).set_cell_value(r, 0, Value::number(100.0));
  }
  HideRow(data, 2);
  const Value v = EvalSourceIn("=SUBTOTAL(109,A1:Data!A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);  // 1 + 2, the Data sheet's rows
}

TEST(BuiltinsSubtotal, ArrayLiteralIgnoresRowVisibility) {
  // A literal array has no sheet row behind it, so every element counts
  // regardless of what the current sheet hides.
  Workbook wb = MakeNumericRange();
  HideRow(wb.sheet(0), 0);
  HideRow(wb.sheet(0), 1);
  const Value v = EvalSourceIn("=SUBTOTAL(109,{1;2;3})", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsSubtotal, DirectScalarArgsIgnoreRowVisibility) {
  Workbook wb = MakeNumericRange();
  HideRow(wb.sheet(0), 0);
  const Value v = EvalSourceIn("=SUBTOTAL(109,1,2,3,4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsSubtotal, HiddenRowIgnoredByFractionalHundredCode) {
  // 109.4 truncates to the hidden-skipping SUM, same as 109.
  Workbook wb = MakeNumericRange();
  HideRow(wb.sheet(0), 1);
  const Value v = EvalSourceIn("=SUBTOTAL(109.4,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 130.0);
}

TEST(BuiltinsSubtotal, MultipleRangeArgsFilterIndependently) {
  Workbook wb = MakeNumericRange();
  HideRow(wb.sheet(0), 0);  // A1 = 10
  HideRow(wb.sheet(0), 4);  // A5 = 50
  const Value v = EvalSourceIn("=SUBTOTAL(109,A1:A2,A4:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 60.0);  // 20 + 40
}

// ---------------------------------------------------------------------------
// Truncating code (matches Excel's TRUNC-before-dispatch rule).
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, FractionalCodeTruncatesToSum) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(9.7,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
}

// ---------------------------------------------------------------------------
// Bad codes
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, ZeroCodeRejected) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSubtotal, OutOfRangeCodeRejected) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(12,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSubtotal, BetweenBlockCodeRejected) {
  // Codes 12..100 fall in the gap between the two valid blocks.
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(50,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSubtotal, AboveHiddenBlockCodeRejected) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(112,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Provenance: range cells of non-numeric kind are skipped (numeric modes)
// or counted (COUNTA mode). Errors propagate in numeric modes; COUNTA
// counts them.
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, SumSkipsNonNumericInRange) {
  Workbook wb = MakeNumericRange();
  wb.sheet(0).set_cell_value(5, 0, Value::text("nope"));
  wb.sheet(0).set_cell_value(6, 0, Value::boolean(true));
  // Blank at row 8 (index 7) silently skipped.
  const Value v = EvalSourceIn("=SUBTOTAL(9,A1:A8)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
}

TEST(BuiltinsSubtotal, CountASkipsBlankButCountsErrors) {
  Workbook wb = MakeNumericRange();
  wb.sheet(0).set_cell_value(5, 0, Value::error(ErrorCode::NA));
  // Row 7 blank.
  const Value v = EvalSourceIn("=SUBTOTAL(3,A1:A7)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  // 5 numbers + 1 error = 6; the blank at A7 is the only skipped cell.
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsSubtotal, SumPropagatesErrorInRange) {
  Workbook wb = MakeNumericRange();
  wb.sheet(0).set_cell_value(2, 0, Value::error(ErrorCode::NA));
  const Value v = EvalSourceIn("=SUBTOTAL(9,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Empty / degenerate ranges
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, AverageOverEmptyRangeIsDiv0) {
  Workbook wb = Workbook::create();
  // A1:A3 entirely blank -> SUBTOTAL(1, A1:A3) sees no numerics.
  const Value v = EvalSourceIn("=SUBTOTAL(1,A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsSubtotal, MinOverEmptyRangeIsZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=SUBTOTAL(5,A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsSubtotal, VarSingleSampleIsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=SUBTOTAL(10,A1:A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Direct (literal) arguments
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, DirectArgsSum) {
  // Mac Excel 365 accepts direct numeric literals as inputs to SUBTOTAL.
  // Our impl drops non-Number direct values silently (matching the range
  // rule); for purely numeric direct args the behaviour matches Mac.
  const Value v = EvalSource("=SUBTOTAL(9,1,2,3,4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsSubtotal, MissingRangeRejected) {
  // Single-arg form is rejected by min_arity = 2 (no implicit empty data).
  const Value v = EvalSource("=SUBTOTAL(9)");
  ASSERT_TRUE(v.is_error());
}

}  // namespace
}  // namespace eval
}  // namespace formulon
