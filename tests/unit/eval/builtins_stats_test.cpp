//
// End-to-end tests for the statistical aggregators: MEDIAN, MODE /
// MODE.SNGL, LARGE / SMALL, PERCENTILE[.INC], QUARTILE[.INC], STDEV[.S],
// STDEV.P, VAR[.S], VAR.P.
//
// These share the `accepts_ranges = true` path with the counting family
// but keep `propagate_errors = true`; text / bool / blank arguments are
// silently skipped inside the impl, errors short-circuit in the
// dispatcher. See the block comment above `Median` in
// src/eval/builtins.cpp for the divergence from SUM / AVERAGE.

#include <cmath>
#include <cstdint>
#include <limits>
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

// Invokes a registered function impl directly; used for Blank-bearing
// cases that cannot be expressed as a formula without a workbook.
Value CallDirect(std::string_view name, const Value* args, std::uint32_t arity) {
  static thread_local Arena arena;
  arena.reset();
  const FunctionDef* def = default_registry().lookup(name);
  EXPECT_NE(def, nullptr) << "function not registered: " << name;
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return def->impl(args, arity, arena);
}

// ---------------------------------------------------------------------------
// Registry pins
// ---------------------------------------------------------------------------

TEST(BuiltinsStatsRegistry, AllNamesRegistered) {
  for (const char* name : {"MEDIAN", "MODE", "MODE.SNGL", "LARGE", "SMALL", "PERCENTILE.INC", "PERCENTILE",
                           "QUARTILE.INC", "QUARTILE", "STDEV.S", "STDEV", "STDEV.P", "VAR.S", "VAR", "VAR.P"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

TEST(BuiltinsStatsRegistry, AllRangeAwareAndPropagating) {
  for (const char* name : {"MEDIAN", "MODE", "MODE.SNGL", "LARGE", "SMALL", "PERCENTILE.INC", "PERCENTILE",
                           "QUARTILE.INC", "QUARTILE", "STDEV.S", "STDEV", "STDEV.P", "VAR.S", "VAR", "VAR.P"}) {
    const FunctionDef* def = default_registry().lookup(name);
    ASSERT_NE(def, nullptr) << name;
    EXPECT_TRUE(def->accepts_ranges) << name;
    EXPECT_TRUE(def->propagate_errors) << name;
  }
}

// ---------------------------------------------------------------------------
// MEDIAN
// ---------------------------------------------------------------------------

TEST(BuiltinsMedian, SingleValue) {
  const Value v = EvalSource("=MEDIAN(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

TEST(BuiltinsMedian, OddCountMiddle) {
  // Sorted: 1, 3, 5, 7, 9 -> middle element is 5.
  const Value v = EvalSource("=MEDIAN(5, 1, 9, 3, 7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsMedian, EvenCountAverageOfTwoMiddle) {
  // Sorted: 1, 2, 3, 4 -> (2 + 3) / 2 = 2.5.
  const Value v = EvalSource("=MEDIAN(1, 2, 3, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.5);
}

TEST(BuiltinsMedian, EvenCountExtremeFiniteValuesStayFinite) {
  const Value v = EvalSource("=MEDIAN(1E308,1E308)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1E308);
}

TEST(BuiltinsMedian, NegativeValues) {
  const Value v = EvalSource("=MEDIAN(-5, -3, -1, 1, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -1.0);
}

TEST(BuiltinsMedian, DirectBoolAndNumericTextCount) {
  // Direct arguments coerce like VAR/STDEV: TRUE is 1, "5" is 5.
  const Value v = EvalSource("=MEDIAN(1, TRUE, 3, \"5\", 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsMedian, DirectNonNumericTextIsValue) {
  const Value v = EvalSource("=MEDIAN(1, \"ignored\", 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMedian, ArrayLiteralNonNumericsAreSkipped) {
  const Value v = EvalSource("=MEDIAN({TRUE, FALSE})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMedian, DirectErrorPropagates) {
  const Value v = EvalSource("=MEDIAN(1, #DIV/0!, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMedian, OmittedSlotCountsAsZero) {
  const Value args[] = {Value::number(1.0), Value::blank(), Value::number(3.0), Value::number(5.0)};
  const Value v = CallDirect("MEDIAN", args, 4u);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMedian, BlankCellReferenceIsSkipped) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 0, Value::number(4.0));
  const Value v = EvalSourceIn("=MEDIAN(A1, A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsMedian, AllBlankCellsIsNum) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=MEDIAN(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMedian, Range) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(20.0));
  const Value v = EvalSourceIn("=MEDIAN(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

TEST(BuiltinsMedian, CrossSheetRangeSkipsNonNumeric) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.sheet(1).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 0, Value::text("skip"));
  wb.sheet(1).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(1).set_cell_value(3, 0, Value::boolean(true));
  wb.sheet(1).set_cell_value(4, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=MEDIAN(Sheet2!A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// MODE / MODE.SNGL
// ---------------------------------------------------------------------------

TEST(BuiltinsMode, UniqueMode) {
  const Value v = EvalSource("=MODE(1, 2, 3, 2, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMode, TripleBeatsDouble) {
  const Value v = EvalSource("=MODE(1, 2, 2, 3, 3, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsMode, TieBreakByFirstOccurrence) {
  // Both 2 and 7 appear twice; 7 appears first in source order, so Excel
  // returns 7.
  const Value v = EvalSource("=MODE(7, 2, 7, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsMode, AllUniqueIsNa) {
  const Value v = EvalSource("=MODE(1, 2, 3, 4, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMode, SingleValueIsNa) {
  // One value cannot repeat, so Excel returns #N/A.
  const Value v = EvalSource("=MODE(42)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMode, NonNumericsAreSkipped) {
  // The non-numerics are ignored; only the two 4s repeat.
  const Value v = EvalSource("=MODE(1, \"4\", 4, TRUE, 4, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsMode, SnglAliasSameAsMode) {
  const Value a = EvalSource("=MODE(3, 1, 3, 2)");
  const Value b = EvalSource("=MODE.SNGL(3, 1, 3, 2)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsMode, Range) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=MODE(A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// MODE.MULT
// ---------------------------------------------------------------------------

TEST(BuiltinsModeMult, SingleModeReturnsOneRowArray) {
  // Only 2 repeats; result is a 1x1 array containing 2.
  const Value v = EvalSource("=MODE.MULT(1, 2, 3, 2, 4)");
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 1u);
  EXPECT_EQ(arr->cols, 1u);
  ASSERT_TRUE(arr->cells[0].is_number());
  EXPECT_DOUBLE_EQ(arr->cells[0].as_number(), 2.0);
}

TEST(BuiltinsModeMult, TiedModesReturnedInFirstOccurrenceOrder) {
  // 1 appears 2x, 2 appears 3x, 3 appears 3x. Ties between 2 and 3 at
  // frequency 3; first occurrence order is 2 (pos 1) then 3 (pos 2).
  const Value v = EvalSource("=MODE.MULT(1, 2, 3, 1, 2, 3, 2, 3)");
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 2u);
  EXPECT_EQ(arr->cols, 1u);
  ASSERT_TRUE(arr->cells[0].is_number());
  ASSERT_TRUE(arr->cells[1].is_number());
  EXPECT_DOUBLE_EQ(arr->cells[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(arr->cells[1].as_number(), 3.0);
}

TEST(BuiltinsModeMult, TieBreakDoesNotReorderByMagnitude) {
  // 7 and 2 each appear twice; 7 appears first in source so it leads the
  // output even though it is the larger value.
  const Value v = EvalSource("=MODE.MULT(7, 2, 7, 2)");
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 2u);
  EXPECT_EQ(arr->cols, 1u);
  EXPECT_DOUBLE_EQ(arr->cells[0].as_number(), 7.0);
  EXPECT_DOUBLE_EQ(arr->cells[1].as_number(), 2.0);
}

TEST(BuiltinsModeMult, AllUniqueIsNa) {
  const Value v = EvalSource("=MODE.MULT(1, 2, 3, 4, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsModeMult, EmptyNumericSliceIsNa) {
  // Only non-numeric values: text and bool are silently skipped (same rule
  // as MODE / MODE.SNGL), leaving an empty numeric slice.
  const Value v = EvalSource("=MODE.MULT(\"a\", \"b\", TRUE, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsModeMult, NonNumericsAreSkippedSameAsMode) {
  // Only the two 4s repeat (text "4" and TRUE are skipped, not coerced).
  // Output is a 1x1 array containing 4.
  const Value v = EvalSource("=MODE.MULT(1, \"4\", 4, TRUE, 4, 2)");
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 1u);
  EXPECT_EQ(arr->cols, 1u);
  EXPECT_DOUBLE_EQ(arr->cells[0].as_number(), 4.0);
}

TEST(BuiltinsModeMult, ErrorPropagates) {
  const Value v = EvalSource("=MODE.MULT(1, 2, 1/0, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsModeMult, RangeInput) {
  // 3 appears twice; result is 1x1 array containing 3.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=MODE.MULT(A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 1u);
  EXPECT_EQ(arr->cols, 1u);
  EXPECT_DOUBLE_EQ(arr->cells[0].as_number(), 3.0);
}

TEST(BuiltinsModeMult, RangeWithMultipleModes) {
  // 2 and 4 both appear twice in the range; output preserves first-
  // occurrence order (2 in row 1, 4 in row 2).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(4.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(4.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(7.0));
  const Value v = EvalSourceIn("=MODE.MULT(A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  EXPECT_EQ(arr->rows, 2u);
  EXPECT_EQ(arr->cols, 1u);
  EXPECT_DOUBLE_EQ(arr->cells[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(arr->cells[1].as_number(), 4.0);
}

TEST(BuiltinsModeMult, ZeroArgRejected) {
  const Value v = EvalSource("=MODE.MULT()");
  ASSERT_TRUE(v.is_error());
}

// ---------------------------------------------------------------------------
// LARGE / SMALL
// ---------------------------------------------------------------------------

TEST(BuiltinsLarge, K1IsMax) {
  const Value v = EvalSource("=LARGE(3, 1, 4, 1, 5, 9, 2, 6, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 9.0);
}

TEST(BuiltinsLarge, K2SecondLargest) {
  // Sorted desc: 9, 6, 5, 4, 3, 2, 1, 1, 1 -> 2nd is 6.
  const Value v = EvalSource("=LARGE(3, 1, 4, 1, 5, 9, 2, 6, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 9.0);
  const Value v2 = EvalSource("=LARGE(3, 1, 4, 1, 5, 9, 2, 6, 2)");
  ASSERT_TRUE(v2.is_number());
  EXPECT_DOUBLE_EQ(v2.as_number(), 6.0);
}

TEST(BuiltinsLarge, KEqualsNIsMin) {
  // 5 numerics: 1, 2, 3, 4, 5. k=5 -> smallest.
  const Value v = EvalSource("=LARGE(1, 2, 3, 4, 5, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsLarge, KTooLargeIsNum) {
  const Value v = EvalSource("=LARGE(1, 2, 3, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsLarge, KZeroIsNum) {
  const Value v = EvalSource("=LARGE(1, 2, 3, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsLarge, KNegativeIsNum) {
  const Value v = EvalSource("=LARGE(1, 2, 3, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsLarge, KFractionalTruncates) {
  // Mac Excel 365 LARGE indexes via TRUNC(N + 1 - k), i.e. CEIL-on-k:
  // for N=5 and k=2.7 the index is TRUNC(5 + 1 - 2.7) = TRUNC(3.3) = 3,
  // so the result is the 3rd element of the ascending sort -> 3.
  // (Previously this test asserted the floor-on-k value 4; that
  // matched the old impl, not Excel. See oracle probes
  // `large_k_fractional_truncates` and `large_k_fractional_at_half` in
  // `tests/oracle/cases/comparison_rank_probes.yaml` for the Mac
  // ground truth.)
  const Value v = EvalSource("=LARGE(1, 2, 3, 4, 5, 2.7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsLarge, DirectScalarTextUnparseableIsValue) {
  // Direct scalar Text arguments go through strict numeric coercion, so an
  // unparseable `"a"` surfaces #VALUE! rather than being silently skipped.
  // Matches Mac Excel 365 and IronCalc's SMALL_LARGE fixture (see J4
  // `=SMALL("Hello", 1)` -> #VALUE!). This replaces the previous lenient
  // behaviour where non-numeric scalars were dropped.
  const Value v = EvalSource("=LARGE(10, \"a\", 20, TRUE, 30, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsLarge, DirectScalarBoolCoerces) {
  // Direct scalar Bool arguments coerce to 1 / 0 (TRUE -> 1, FALSE -> 0)
  // for SMALL / LARGE, matching the Mac Excel 365 oracle.
  //   data = [10, 1, 20, 30] after coercion; k=1 -> max = 30.
  const Value v = EvalSource("=LARGE(10, TRUE, 20, 30, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsLarge, DirectScalarNumericTextCoerces) {
  // Direct scalar numeric-looking Text coerces to its number; `"3.4"` -> 3.4.
  const Value v = EvalSource("=LARGE(\"3.4\", 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.4);
}

TEST(BuiltinsSmall, DirectScalarNumericTextCoerces) {
  // Mirror of LARGE; IronCalc fixture J3 `=SMALL("3.4", 1)` -> 3.4.
  const Value v = EvalSource("=SMALL(\"3.4\", 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.4);
}

TEST(BuiltinsSmall, DirectScalarBoolCoerces) {
  // IronCalc fixture H4 `=SMALL(TRUE, 1)` -> 1.0.
  const Value v = EvalSource("=SMALL(TRUE, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsSmall, DirectScalarTextUnparseableIsValue) {
  // IronCalc fixture J4 `=SMALL("Hello", 1)` -> #VALUE!.
  const Value v = EvalSource("=SMALL(\"Hello\", 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSmall, ArrayLiteralNonNumericSkipped) {
  // Array-literal elements are treated like range cells: non-Number
  // elements are dropped silently by `range_filter_numeric_only` inside
  // the ArrayLiteral dispatch branch. Here `{1, FALSE, TRUE}` reduces to
  // xs=[1.0], so SMALL(..., 1) -> 1.0. IronCalc's F10 fixture uses a Text
  // element (`{1, "-1"}`) which the parser does not yet accept inside a
  // brace literal (array literal grammar does not yet allow scalar
  // expressions). When parser support for Text/scalar expressions in
  // array literals lands, an equivalent `{1,"-1"}` case can be added.
  const Value v = EvalSource("=SMALL({1, FALSE, TRUE}, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsLarge, RangeAndScalarK) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(40.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(30.0));
  const Value v = EvalSourceIn("=LARGE(A1:A4, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsLarge, CrossSheetRange) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.sheet(1).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(1).set_cell_value(1, 0, Value::number(15.0));
  wb.sheet(1).set_cell_value(2, 0, Value::number(25.0));
  const Value v = EvalSourceIn("=LARGE(Sheet2!A1:A3, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 25.0);
}

TEST(BuiltinsSmall, K1IsMin) {
  const Value v = EvalSource("=SMALL(3, 1, 4, 1, 5, 9, 2, 6, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsSmall, K2SecondSmallest) {
  // Sorted asc: 1, 2, 3, 4, 5 -> 2nd is 2.
  const Value v = EvalSource("=SMALL(5, 4, 3, 2, 1, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsSmall, KEqualsNIsMax) {
  const Value v = EvalSource("=SMALL(1, 2, 3, 4, 5, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsSmall, KOutOfRangeIsNum) {
  const Value v = EvalSource("=SMALL(1, 2, 3, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsSmall, Range) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(20.0));
  const Value v = EvalSourceIn("=SMALL(A1:A3, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

// ---------------------------------------------------------------------------
// PERCENTILE.INC / PERCENTILE
// ---------------------------------------------------------------------------

TEST(BuiltinsPercentileInc, K0IsMin) {
  const Value v = EvalSource("=PERCENTILE.INC(10, 20, 30, 40, 50, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsPercentileInc, K1IsMax) {
  const Value v = EvalSource("=PERCENTILE.INC(10, 20, 30, 40, 50, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
}

TEST(BuiltinsPercentileInc, KHalfIsMedian) {
  // 5 values, k=0.5 -> pos = 0.5 * 4 = 2 -> exact index -> 30.
  const Value v = EvalSource("=PERCENTILE.INC(10, 20, 30, 40, 50, 0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsPercentileInc, KQuarterInterpolates) {
  // 5 values, k=0.25 -> pos = 0.25 * 4 = 1 -> exact -> 20.
  const Value v = EvalSource("=PERCENTILE.INC(10, 20, 30, 40, 50, 0.25)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

TEST(BuiltinsPercentileInc, FractionalInterpolation) {
  // 4 values, k=0.5 -> pos = 0.5 * 3 = 1.5 -> blend of idx 1 (20) and 2 (30).
  // r = 20 + 0.5 * (30 - 20) = 25.
  const Value v = EvalSource("=PERCENTILE.INC(10, 20, 30, 40, 0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 25.0);
}

TEST(BuiltinsPercentileInc, OppositeExtremeFiniteEndpointsInterpolate) {
  const Value v = EvalSource("=PERCENTILE.INC(-1E308,1E308,0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsPercentileInc, KAbove1IsNum) {
  const Value v = EvalSource("=PERCENTILE.INC(10, 20, 30, 1.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsPercentileInc, KBelow0IsNum) {
  const Value v = EvalSource("=PERCENTILE.INC(10, 20, 30, -0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsPercentileInc, EmptyNumericSliceIsNum) {
  const Value v = EvalSource("=PERCENTILE.INC(\"a\", \"b\", 0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsPercentileInc, AliasMatchesCanonical) {
  const Value a = EvalSource("=PERCENTILE.INC(1, 2, 3, 4, 0.3)");
  const Value b = EvalSource("=PERCENTILE(1, 2, 3, 4, 0.3)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsPercentileInc, RangeAndScalarK) {
  Workbook wb = Workbook::create();
  for (std::uint32_t i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>((i + 1) * 10)));
  }
  const Value v = EvalSourceIn("=PERCENTILE.INC(A1:A5, 0.5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

// ---------------------------------------------------------------------------
// QUARTILE.INC / QUARTILE
// ---------------------------------------------------------------------------

TEST(BuiltinsQuartileInc, KnownDataset) {
  // 7 values sorted: 1, 2, 4, 7, 8, 9, 10.
  //   q=0 -> min=1
  //   q=4 -> max=10
  //   q=2 -> median=7
  //   q=1 -> pos = 0.25 * 6 = 1.5 -> blend(2, 4) = 3.0
  //   q=3 -> pos = 0.75 * 6 = 4.5 -> blend(8, 9) = 8.5
  const Value q0 = EvalSource("=QUARTILE.INC(1, 2, 4, 7, 8, 9, 10, 0)");
  const Value q1 = EvalSource("=QUARTILE.INC(1, 2, 4, 7, 8, 9, 10, 1)");
  const Value q2 = EvalSource("=QUARTILE.INC(1, 2, 4, 7, 8, 9, 10, 2)");
  const Value q3 = EvalSource("=QUARTILE.INC(1, 2, 4, 7, 8, 9, 10, 3)");
  const Value q4 = EvalSource("=QUARTILE.INC(1, 2, 4, 7, 8, 9, 10, 4)");
  ASSERT_TRUE(q0.is_number());
  ASSERT_TRUE(q1.is_number());
  ASSERT_TRUE(q2.is_number());
  ASSERT_TRUE(q3.is_number());
  ASSERT_TRUE(q4.is_number());
  EXPECT_DOUBLE_EQ(q0.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(q1.as_number(), 3.0);
  EXPECT_DOUBLE_EQ(q2.as_number(), 7.0);
  EXPECT_DOUBLE_EQ(q3.as_number(), 8.5);
  EXPECT_DOUBLE_EQ(q4.as_number(), 10.0);
}

TEST(BuiltinsQuartileInc, QAbove4IsNum) {
  const Value v = EvalSource("=QUARTILE.INC(1, 2, 3, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsQuartileInc, QNegativeIsNum) {
  const Value v = EvalSource("=QUARTILE.INC(1, 2, 3, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsQuartileInc, QFractionalTruncatesTowardZero) {
  // Excel truncates a fractional quart toward zero: 1.5 -> Q1 -> P25 of {1,2,3,4}.
  const Value q15 = EvalSource("=QUARTILE.INC(1, 2, 3, 4, 1.5)");
  const Value q1 = EvalSource("=QUARTILE.INC(1, 2, 3, 4, 1)");
  ASSERT_TRUE(q15.is_number());
  ASSERT_TRUE(q1.is_number());
  EXPECT_DOUBLE_EQ(q15.as_number(), q1.as_number());
}

TEST(BuiltinsQuartileInc, AliasMatchesCanonical) {
  const Value a = EvalSource("=QUARTILE.INC(1, 2, 4, 7, 8, 9, 10, 1)");
  const Value b = EvalSource("=QUARTILE(1, 2, 4, 7, 8, 9, 10, 1)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

// ---------------------------------------------------------------------------
// STDEV.S / STDEV / STDEV.P
// ---------------------------------------------------------------------------

TEST(BuiltinsStdevS, KnownDataset) {
  // {2, 4, 4, 4, 5, 5, 7, 9}: mean=5, SS=32, sample stdev=sqrt(32/7)~2.1380899.
  // This is the same dataset as the Wikipedia standard deviation article,
  // where sample stdev = sqrt(32/7) and population stdev = sqrt(4) = 2.
  const Value v = EvalSource("=STDEV.S(2, 4, 4, 4, 5, 5, 7, 9)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::sqrt(32.0 / 7.0), 1e-12);
}

TEST(BuiltinsStdevS, AliasMatchesCanonical) {
  const Value a = EvalSource("=STDEV.S(1, 2, 3, 4, 5)");
  const Value b = EvalSource("=STDEV(1, 2, 3, 4, 5)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsStdevS, SingleValueIsDiv0) {
  const Value v = EvalSource("=STDEV.S(5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsStdevS, DirectNonNumericTextIsValueError) {
  // Microsoft's documented rule: "text that cannot be translated into
  // numbers cause errors" for DIRECT arguments. Unlike MEDIAN / MODE,
  // STDEV.S does not silently skip an unparseable direct text argument.
  const Value v = EvalSource("=STDEV.S(1, \"a\", 2, TRUE, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsStdevS, DirectBooleanAndNumericTextAreCounted) {
  // Microsoft's documented rule: "Logical values and text representations
  // of numbers that you type directly into the list of arguments are
  // counted". {2, TRUE, "3"} -> {2, 1, 3}: mean=2, SS=2, sample stdev=1.
  const Value v = EvalSource("=STDEV.S(2, TRUE, \"3\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsStdevS, RangeSourcedNonNumericStillSkipped) {
  // The direct-argument coercion rule does not extend to range-sourced
  // cells: A2 (TRUE) and A3 ("x") are dropped, leaving only {1, 3}.
  // sample stdev of {1, 3} = sqrt(2/1) = sqrt(2).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 0, Value::number(3.0));
  const Value v = EvalSourceIn("=STDEV.S(A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::sqrt(2.0), 1e-12);
}

TEST(BuiltinsStdevS, Range) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalSourceIn("=STDEV.S(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsStdevP, KnownDataset) {
  // Same {2,4,4,4,5,5,7,9} -> population stdev = sqrt(32/8) = 2.
  const Value v = EvalSource("=STDEV.P(2, 4, 4, 4, 5, 5, 7, 9)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.0, 1e-12);
}

TEST(BuiltinsStdevP, SingleValueIsZero) {
  const Value v = EvalSource("=STDEV.P(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsStdevP, DirectNonNumericTextIsValueError) {
  // "a" is a direct argument that does not parse as a number.
  const Value v = EvalSource("=STDEV.P(\"a\", \"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// VAR.S / VAR / VAR.P
// ---------------------------------------------------------------------------

TEST(BuiltinsVarS, KnownDataset) {
  // {2,4,4,4,5,5,7,9}: sample variance = 32/7.
  const Value v = EvalSource("=VAR.S(2, 4, 4, 4, 5, 5, 7, 9)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 32.0 / 7.0, 1e-12);
}

TEST(BuiltinsVarS, AliasMatchesCanonical) {
  const Value a = EvalSource("=VAR.S(1, 2, 3, 4, 5)");
  const Value b = EvalSource("=VAR(1, 2, 3, 4, 5)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsVarS, SingleValueIsDiv0) {
  const Value v = EvalSource("=VAR.S(5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsVarP, KnownDataset) {
  // {2,4,4,4,5,5,7,9}: population variance = 32/8 = 4.
  const Value v = EvalSource("=VAR.P(2, 4, 4, 4, 5, 5, 7, 9)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 4.0, 1e-12);
}

TEST(BuiltinsVarP, SingleValueIsZero) {
  const Value v = EvalSource("=VAR.P(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsVarP, DirectNonNumericTextIsValueError) {
  // "a" is a direct argument that does not parse as a number; TRUE would
  // count (direct Bool -> 1) but "a" errors the whole call first.
  const Value v = EvalSource("=VAR.P(\"a\", TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsVarP, DirectBooleanIsCounted) {
  // {2, TRUE} -> {2, 1}: mean=1.5, SS=0.5, population variance=0.25.
  const Value v = EvalSource("=VAR.P(2, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.25, 1e-12);
}

TEST(BuiltinsVarP, Range) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(4.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(4.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(4.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(5, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(6, 0, Value::number(7.0));
  wb.sheet(0).set_cell_value(7, 0, Value::number(9.0));
  const Value v = EvalSourceIn("=VAR.P(A1:A8)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 4.0, 1e-12);
}

TEST(BuiltinsStatsExtreme, EqualLargeRangeOverflowIsNumError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0e308));
  wb.sheet(0).set_cell_value(1, 0, Value::number(1.0e308));

  for (const char* formula : {"=VAR.S(A1:A2)", "=VAR.P(A1:A2)", "=STDEV.S(A1:A2)", "=STDEV.P(A1:A2)"}) {
    const Value result = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(result.is_error()) << formula << ": " << result.debug_to_string();
    EXPECT_EQ(result.as_error(), ErrorCode::Num) << formula;
  }
}

TEST(BuiltinsStatsExtreme, OpposingMaximumPopulationStdDevIsNumError) {
  Workbook wb = Workbook::create();
  const double max = std::numeric_limits<double>::max();
  wb.sheet(0).set_cell_value(0, 0, Value::number(max));
  wb.sheet(0).set_cell_value(1, 0, Value::number(-max));

  const Value stdev = EvalSourceIn("=STDEV.P(A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(stdev.is_error()) << stdev.debug_to_string();
  EXPECT_EQ(stdev.as_error(), ErrorCode::Num);
}

TEST(BuiltinsStatsExtreme, SquaredDeviationOverflowIsNumError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0e155));
  for (std::uint32_t row = 1U; row < 1000U; ++row) {
    wb.sheet(0).set_cell_value(row, 0, Value::number(0.0));
  }

  const Value variance = EvalSourceIn("=VAR.P(A1:A1000)", wb, wb.sheet(0));
  ASSERT_TRUE(variance.is_error()) << variance.debug_to_string();
  EXPECT_EQ(variance.as_error(), ErrorCode::Num);
}

TEST(BuiltinsStatsExtreme, AdjacentLargeValuesMatchExcelPopulationVariance) {
  constexpr double first = 1.0e150;
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(first));
  wb.sheet(0).set_cell_value(1, 0, Value::number(std::nextafter(first, 0.0)));

  const Value variance = EvalSourceIn("=VAR.P(A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(variance.is_number()) << variance.debug_to_string();
  EXPECT_DOUBLE_EQ(variance.as_number(), 1.650920409798954e+268);
}

TEST(BuiltinsStatsEdges, TrimMeanBlankPercentCellTrimsNothing) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  const double data[] = {12.4, 5.0, 7.3, -4.0, 0.0, 2.3, 4.3, 7.0, -2.0};
  for (std::uint32_t i = 0; i < 9U; ++i) {
    s.set_cell_value(i, 0, Value::number(data[i]));
  }
  const Value v = EvalSourceIn("=TRIMMEAN(A1:A9,B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 3.5888888888888886);
}

TEST(BuiltinsStatsEdges, FisherInvSaturatesToOne) {
  for (const char* formula : {"=FISHERINV(1000)", "=FISHERINV(10000000000)"}) {
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_DOUBLE_EQ(v.as_number(), 1.0) << formula;
  }
}

TEST(BuiltinsStatsEdges, TInvRoundTripsAnExtremeLeftTail) {
  const Value v = EvalSource("=T.INV(T.DIST(-1000000,10,TRUE),10)");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_NEAR(v.as_number(), -1000000.0000000001, 1e-6);
}

// The k of LARGE / SMALL / PERCENTILE / QUARTILE / TRIMMEAN is a single value:
// a cell reference there coerces as a direct argument instead of being
// filtered like the data range (Mac Excel 365).
TEST(BuiltinsStatsScalarArg, CellKCoercesAsADirectArgument) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  for (std::uint32_t i = 0; i < 7U; ++i) {
    s.set_cell_value(i, 0, Value::number(10.0 * (i + 1)));
  }
  s.set_cell_value(0, 1, Value::boolean(true));
  s.set_cell_value(1, 1, Value::text("0.4"));
  s.set_cell_value(2, 1, Value::text("abc"));
  struct Case {
    const char* source;
    double expected;
  };
  for (const Case& c : {Case{"=LARGE(A1:A7,B1)", 70.0}, Case{"=SMALL(A1:A7,B1)", 10.0},
                        Case{"=QUARTILE.EXC(A1:A7,B1)", 20.0}, Case{"=PERCENTILE.EXC(A1:A7,B2)", 32.0},
                        Case{"=PERCENTILE.INC(A1:A7,B2)", 34.0}, Case{"=TRIMMEAN(A1:A7,B4)", 40.0}}) {
    const Value v = EvalSourceIn(c.source, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << c.source << ": " << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), c.expected) << c.source;
  }
  for (const char* source : {"=LARGE(A1:A7,B3)", "=SMALL(A1:A7,B3)", "=PERCENTILE.EXC(A1:A7,B3)"}) {
    const Value v = EvalSourceIn(source, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_error()) << source;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << source;
  }
}

TEST(BuiltinsStatsScalarArg, ArrayKEvaluatesPerElement) {
  const Value large = EvalSource("=LARGE({10;20;30;40},{1,2})");
  ASSERT_TRUE(large.is_array()) << large.debug_to_string();
  ASSERT_EQ(large.as_array()->rows, 1U);
  ASSERT_EQ(large.as_array()->cols, 2U);
  EXPECT_DOUBLE_EQ(large.as_array()->cells[0].as_number(), 40.0);
  EXPECT_DOUBLE_EQ(large.as_array()->cells[1].as_number(), 30.0);
  const Value pct = EvalSource("=PERCENTILE.INC({10;20;30},{0;2})");
  ASSERT_TRUE(pct.is_array()) << pct.debug_to_string();
  EXPECT_DOUBLE_EQ(pct.as_array()->cells[0].as_number(), 10.0);
  EXPECT_EQ(pct.as_array()->cells[1].as_error(), ErrorCode::Num);
}

TEST(BuiltinsStatsScalarArg, PercentileExcAcceptsBothEndsOfThePosition) {
  struct Case {
    const char* source;
    double expected;
  };
  for (const Case& c : {Case{"=PERCENTILE.EXC({1,2,3},0.75)", 3.0}, Case{"=PERCENTILE.EXC({1,2,3},0.25)", 1.0},
                        Case{"=QUARTILE.EXC({10},2)", 10.0}, Case{"=PERCENTILE.EXC({10,20},2/3)", 20.0}}) {
    const Value v = EvalSource(c.source);
    ASSERT_TRUE(v.is_number()) << c.source << ": " << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), c.expected) << c.source;
  }
  const Value past = EvalSource("=PERCENTILE.EXC({10},0.501)");
  ASSERT_TRUE(past.is_error());
  EXPECT_EQ(past.as_error(), ErrorCode::Num);
}

TEST(BuiltinsStatsScalarArg, AverageASnapsCancellation) {
  const Value v = EvalSource("=AVERAGEA(0.5,-0.4,-0.1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
  EXPECT_FALSE(std::signbit(v.as_number()));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
