// Copyright 2026 libraz. Licensed under the MIT License.
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
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
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
