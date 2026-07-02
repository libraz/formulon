// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Regression tests for a batch of eval-layer fixes:
//   1. General Excel array broadcasting in `broadcast_binop` (outer product,
//      RxC / Rx1 / 1xC combinations, #N/A padding on mismatched non-1 axes).
//   2. MATCH approximate mode skipping blank cells (not treating them as 0).
//   3. XLOOKUP explicit-empty `if_not_found` behaving like an omitted arg.
//   4. UnionOp `(A1:A2,B1:B2)` flowing through range-aware aggregators.
//   5. Half-width voicing composition applied to classic lookups (VLOOKUP /
//      MATCH / HLOOKUP), matching XLOOKUP.

#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/tree_walker/broadcast.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "util/test_eval_helpers.h"
#include "util/test_value_macros.h"
#include "util/test_workbook_builder.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using test::EvalSource;
using test::EvalSourceIn;
using test::WorkbookBuilder;

// Builds a row-major Array `Value` from `cells` (rows*cols entries) in `arena`.
Value MakeArray(Arena& arena, std::uint32_t rows, std::uint32_t cols, const std::vector<Value>& cells) {
  const std::size_t n = static_cast<std::size_t>(rows) * cols;
  Value* buf = arena.create_array<Value>(n);
  for (std::size_t i = 0; i < n && i < cells.size(); ++i) {
    buf[i] = cells[i];
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  arr->rows = rows;
  arr->cols = cols;
  arr->cells = buf;
  return Value::array(arr);
}

// ---------------------------------------------------------------------------
// 1. broadcast_binop general shape rules
// ---------------------------------------------------------------------------

TEST(BroadcastBinop, ColumnTimesRowIsOuterProduct) {
  // {1;2;3} (3x1) * {10,20} (1x2) -> 3x2 outer product.
  Arena a;
  const Value lhs = MakeArray(a, 3U, 1U, {Value::number(1), Value::number(2), Value::number(3)});
  const Value rhs = MakeArray(a, 1U, 2U, {Value::number(10), Value::number(20)});
  const Value out = broadcast_binop(parser::BinOp::Mul, lhs, rhs, a);
  ASSERT_TRUE(out.is_array());
  const ArrayValue* r = out.as_array();
  ASSERT_EQ(r->rows, 3U);
  ASSERT_EQ(r->cols, 2U);
  const double expected[6] = {10, 20, 20, 40, 30, 60};
  for (std::size_t i = 0; i < 6; ++i) {
    ASSERT_TRUE(r->cells[i].is_number()) << "cell " << i;
    EXPECT_DOUBLE_EQ(r->cells[i].as_number(), expected[i]) << "cell " << i;
  }
}

TEST(BroadcastBinop, RowVectorPlusColumnVectorBroadcasts) {
  // {1,2,3} (1x3) + {10;20} (2x1) -> 2x3, every cell defined (clean broadcast).
  Arena a;
  const Value lhs = MakeArray(a, 1U, 3U, {Value::number(1), Value::number(2), Value::number(3)});
  const Value rhs = MakeArray(a, 2U, 1U, {Value::number(10), Value::number(20)});
  const Value out = broadcast_binop(parser::BinOp::Add, lhs, rhs, a);
  ASSERT_TRUE(out.is_array());
  const ArrayValue* r = out.as_array();
  ASSERT_EQ(r->rows, 2U);
  ASSERT_EQ(r->cols, 3U);
  const double expected[6] = {11, 12, 13, 21, 22, 23};
  for (std::size_t i = 0; i < 6; ++i) {
    ASSERT_TRUE(r->cells[i].is_number()) << "cell " << i;
    EXPECT_DOUBLE_EQ(r->cells[i].as_number(), expected[i]) << "cell " << i;
  }
}

TEST(BroadcastBinop, MismatchedNonUnitAxisPadsWithNA) {
  // {1,2,3} (1x3) + {1,2} (1x2): cols 3 vs 2, both > 1 -> extend to 3 with the
  // shortfall filled with #N/A -> {2, 4, #N/A}.
  Arena a;
  const Value lhs = MakeArray(a, 1U, 3U, {Value::number(1), Value::number(2), Value::number(3)});
  const Value rhs = MakeArray(a, 1U, 2U, {Value::number(1), Value::number(2)});
  const Value out = broadcast_binop(parser::BinOp::Add, lhs, rhs, a);
  ASSERT_TRUE(out.is_array());
  const ArrayValue* r = out.as_array();
  ASSERT_EQ(r->rows, 1U);
  ASSERT_EQ(r->cols, 3U);
  EXPECT_DOUBLE_EQ(r->cells[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(r->cells[1].as_number(), 4.0);
  ASSERT_TRUE(r->cells[2].is_error());
  EXPECT_EQ(r->cells[2].as_error(), ErrorCode::NA);
}

TEST(BroadcastBinop, EqualShapesAreElementwise) {
  // Two 3x1 columns multiply element-wise.
  Arena a;
  const Value lhs = MakeArray(a, 3U, 1U, {Value::number(2), Value::number(3), Value::number(4)});
  const Value rhs = MakeArray(a, 3U, 1U, {Value::number(10), Value::number(100), Value::number(1000)});
  const Value out = broadcast_binop(parser::BinOp::Mul, lhs, rhs, a);
  ASSERT_TRUE(out.is_array());
  const ArrayValue* r = out.as_array();
  ASSERT_EQ(r->rows, 3U);
  ASSERT_EQ(r->cols, 1U);
  EXPECT_DOUBLE_EQ(r->cells[0].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(r->cells[1].as_number(), 300.0);
  EXPECT_DOUBLE_EQ(r->cells[2].as_number(), 4000.0);
}

TEST(BroadcastBinop, ScalarBroadcastsAcrossArray) {
  Arena a;
  const Value lhs = MakeArray(a, 1U, 1U, {Value::number(5)});
  const Value rhs = MakeArray(a, 2U, 2U, {Value::number(1), Value::number(2), Value::number(3), Value::number(4)});
  const Value out = broadcast_binop(parser::BinOp::Add, lhs, rhs, a);
  ASSERT_TRUE(out.is_array());
  const ArrayValue* r = out.as_array();
  ASSERT_EQ(r->rows, 2U);
  ASSERT_EQ(r->cols, 2U);
  const double expected[4] = {6, 7, 8, 9};
  for (std::size_t i = 0; i < 4; ++i) {
    EXPECT_DOUBLE_EQ(r->cells[i].as_number(), expected[i]) << "cell " << i;
  }
}

TEST(BroadcastBinop, OuterProductEndToEndViaSequence) {
  // SEQUENCE produces genuine Value::Arrays, so `=SEQUENCE(3,1)*SEQUENCE(1,2)`
  // reaches broadcast_binop through the top-level BinaryOp path. Summed via
  // SUMPRODUCT: outer product {1,2;2,4;3,6} totals 18.
  const Value v = EvalSource("=SUMPRODUCT(SEQUENCE(3,1)*SEQUENCE(1,2))");
  EXPECT_VALUE_NUMBER(v, 18.0);
}

// ---------------------------------------------------------------------------
// 2. MATCH approximate skips blank cells
// ---------------------------------------------------------------------------

TEST(MatchApproxBlank, BlankCellIsSkippedNotTreatedAsZero) {
  // Ascending column with a blank at A3: [1, 3, <blank>, 5]. MATCH(4, .., 1)
  // must skip the blank (not rank it as 0) and return the position of 3 (=2).
  Workbook wb = WorkbookBuilder().sheet("Sheet1").cell("A1", 1.0).cell("A2", 3.0).cell("A4", 5.0).build();
  const Value v = EvalSourceIn("=MATCH(4, A1:A4, 1)", wb, wb.sheet(0));
  EXPECT_VALUE_NUMBER(v, 2.0);
}

// ---------------------------------------------------------------------------
// 3. XLOOKUP if_not_found: omitted vs explicit-empty vs ""
// ---------------------------------------------------------------------------

TEST(XlookupIfNotFound, OmittedFourthArgIsNA) {
  Workbook wb =
      WorkbookBuilder().sheet("Sheet1").cell("A1", 1.0).cell("A2", 2.0).cell("B1", 10.0).cell("B2", 20.0).build();
  const Value v = EvalSourceIn("=XLOOKUP(\"zz\", A1:A2, B1:B2)", wb, wb.sheet(0));
  EXPECT_VALUE_ERROR(v, ErrorCode::NA);
}

TEST(XlookupIfNotFound, ExplicitEmptySlotIsNA) {
  // Trailing comma = empty arg slot; the parser injects a Blank literal and
  // XLOOKUP treats it as "no if_not_found supplied" -> #N/A (Mac Excel 365).
  Workbook wb =
      WorkbookBuilder().sheet("Sheet1").cell("A1", 1.0).cell("A2", 2.0).cell("B1", 10.0).cell("B2", 20.0).build();
  const Value v = EvalSourceIn("=XLOOKUP(\"zz\", A1:A2, B1:B2,)", wb, wb.sheet(0));
  EXPECT_VALUE_ERROR(v, ErrorCode::NA);
}

TEST(XlookupIfNotFound, ExplicitEmptyStringReturnsEmptyString) {
  Workbook wb =
      WorkbookBuilder().sheet("Sheet1").cell("A1", 1.0).cell("A2", 2.0).cell("B1", 10.0).cell("B2", 20.0).build();
  const Value v = EvalSourceIn("=XLOOKUP(\"zz\", A1:A2, B1:B2, \"\")", wb, wb.sheet(0));
  EXPECT_VALUE_TEXT(v, "");
}

// ---------------------------------------------------------------------------
// 4. UnionOp as a range-aware aggregator argument
// ---------------------------------------------------------------------------

TEST(UnionArg, SumOverTwoAreas) {
  Workbook wb =
      WorkbookBuilder().sheet("Sheet1").cell("A1", 1.0).cell("A2", 2.0).cell("B1", 10.0).cell("B2", 20.0).build();
  const Value v = EvalSourceIn("=SUM((A1:A2,B1:B2))", wb, wb.sheet(0));
  EXPECT_VALUE_NUMBER(v, 33.0);
}

TEST(UnionArg, ProductOverTwoAreas) {
  // PRODUCT is another eager range-aware aggregator: 1*2*10*20 = 400.
  // (COUNT's union support lives in the lazy `eval_count_lazy` path owned by a
  // separate lane; see the batch report.)
  Workbook wb =
      WorkbookBuilder().sheet("Sheet1").cell("A1", 1.0).cell("A2", 2.0).cell("B1", 10.0).cell("B2", 20.0).build();
  const Value v = EvalSourceIn("=PRODUCT((A1:A2,B1:B2))", wb, wb.sheet(0));
  EXPECT_VALUE_NUMBER(v, 400.0);
}

TEST(UnionArg, AverageOverTwoAreas) {
  Workbook wb =
      WorkbookBuilder().sheet("Sheet1").cell("A1", 1.0).cell("A2", 2.0).cell("B1", 10.0).cell("B2", 20.0).build();
  const Value v = EvalSourceIn("=AVERAGE((A1:A2,B1:B2))", wb, wb.sheet(0));
  EXPECT_VALUE_NUMBER(v, 8.25);
}

TEST(UnionArg, OverlappingAreasAreDoubleCounted) {
  // Excel's union does NOT de-duplicate overlapping cells: (A1:A2,A1:A2)
  // counts A1:A2 twice, so SUM = 2*(1+2) = 6.
  Workbook wb = WorkbookBuilder().sheet("Sheet1").cell("A1", 1.0).cell("A2", 2.0).build();
  const Value v = EvalSourceIn("=SUM((A1:A2,A1:A2))", wb, wb.sheet(0));
  EXPECT_VALUE_NUMBER(v, 6.0);
}

// ---------------------------------------------------------------------------
// 5. Half-width voicing composition in classic lookups (VLOOKUP / MATCH)
// ---------------------------------------------------------------------------

// Half-width katakana KA (U+FF76) + half-width voiced mark (U+FF9E): a base +
// standalone voicing mark that Excel composes to full-width GA (U+30AC) before
// text comparison. The lookup key uses the decomposed half-width form; the
// stored cell uses the composed full-width form.
constexpr const char* kHalfWidthKaVoiced = "\xEF\xBD\xB6\xEF\xBE\x9E";  // ｶﾞ
constexpr const char* kFullWidthGa = "\xE3\x82\xAC";                    // ガ

TEST(HalfWidthVoicing, VlookupComposesLikeXlookup) {
  Workbook wb = WorkbookBuilder().sheet("Sheet1").text_cell("A1", kFullWidthGa).cell("B1", 100.0).build();
  const std::string vlookup = std::string("=VLOOKUP(\"") + kHalfWidthKaVoiced + "\", A1:B1, 2, FALSE)";
  const Value v = EvalSourceIn(vlookup, wb, wb.sheet(0));
  EXPECT_VALUE_NUMBER(v, 100.0);
  // XLOOKUP already composes; confirm classic now agrees.
  const std::string xlookup = std::string("=XLOOKUP(\"") + kHalfWidthKaVoiced + "\", A1:A1, B1:B1)";
  const Value vx = EvalSourceIn(xlookup, wb, wb.sheet(0));
  EXPECT_VALUE_NUMBER(vx, 100.0);
}

TEST(HalfWidthVoicing, MatchComposesLikeXmatch) {
  Workbook wb = WorkbookBuilder().sheet("Sheet1").text_cell("A1", kFullWidthGa).build();
  const std::string match = std::string("=MATCH(\"") + kHalfWidthKaVoiced + "\", A1:A1, 0)";
  const Value v = EvalSourceIn(match, wb, wb.sheet(0));
  EXPECT_VALUE_NUMBER(v, 1.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
