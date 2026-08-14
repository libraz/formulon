//
// End-to-end tests for the lazy-dispatched XLOOKUP and XMATCH functions.
// Both share the `xlookup_scan` helper in `tree_walker.cpp` that linearly
// (FirstToLast / LastToFirst) or binary-searches (BinaryAsc / BinaryDesc) a
// 1-D lookup_array using one of four match modes (Exact / Smaller / Larger
// / Wildcard).

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
  return evaluate(*root, eval_arena, default_registry(), test::mac_context());
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
  const EvalContext ctx = test::mac_context(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// Populates column A with `keys` and column B with `values` on sheet 0.
void SeedColumnPair(Workbook& wb, std::initializer_list<Value> keys, std::initializer_list<Value> values) {
  std::uint32_t row = 0;
  auto kit = keys.begin();
  auto vit = values.begin();
  for (; kit != keys.end() && vit != values.end(); ++kit, ++vit, ++row) {
    wb.sheet(0).set_cell_value(row, 0, *kit);
    wb.sheet(0).set_cell_value(row, 1, *vit);
  }
}

TEST(BuiltinsXLookup, ArrayLiteralLookupAndReturnArrays) {
  const Value v = EvalSource("=XLOOKUP(2,{1,2,3},{10,20,30})");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

TEST(BuiltinsXLookup, ArrayLookupValuePreservesVerticalAndHorizontalRangeShape) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0)},
                 {Value::number(100.0), Value::number(200.0), Value::number(300.0)});
  wb.sheet(0).set_cell_value(0, 3, Value::number(20.0));  // D1
  wb.sheet(0).set_cell_value(1, 3, Value::number(10.0));  // D2
  wb.sheet(0).set_cell_value(2, 3, Value::number(99.0));  // D3
  const Value vertical = EvalSourceIn("=XLOOKUP(D1:D3,A1:A3,B1:B3,\"missing\")", wb, wb.sheet(0));
  ASSERT_TRUE(vertical.is_array()) << vertical.debug_to_string();
  EXPECT_EQ(vertical.as_array_rows(), 3U);
  EXPECT_EQ(vertical.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(vertical.as_array_cells()[0].as_number(), 200.0);
  EXPECT_DOUBLE_EQ(vertical.as_array_cells()[1].as_number(), 100.0);
  ASSERT_TRUE(vertical.as_array_cells()[2].is_text());
  EXPECT_EQ(vertical.as_array_cells()[2].as_text(), "missing");

  wb.sheet(0).set_cell_value(3, 3, Value::number(30.0));  // D4
  wb.sheet(0).set_cell_value(3, 4, Value::number(10.0));  // E4
  wb.sheet(0).set_cell_value(3, 5, Value::number(20.0));  // F4
  const Value horizontal = EvalSourceIn("=XLOOKUP(D4:F4,A1:A3,B1:B3,\"missing\")", wb, wb.sheet(0));
  ASSERT_TRUE(horizontal.is_array()) << horizontal.debug_to_string();
  EXPECT_EQ(horizontal.as_array_rows(), 1U);
  EXPECT_EQ(horizontal.as_array_cols(), 3U);
  EXPECT_DOUBLE_EQ(horizontal.as_array_cells()[0].as_number(), 300.0);
  EXPECT_DOUBLE_EQ(horizontal.as_array_cells()[1].as_number(), 100.0);
  EXPECT_DOUBLE_EQ(horizontal.as_array_cells()[2].as_number(), 200.0);
}

TEST(BuiltinsXLookup, TwoDimensionalArrayLookupValuePreservesRowMajorLanes) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0)},
                 {Value::number(100.0), Value::number(200.0), Value::number(300.0)});

  const Value xlookup = EvalSourceIn("=XLOOKUP({30,10;99,20},A1:A3,B1:B3,\"missing\")", wb, wb.sheet(0));
  ASSERT_TRUE(xlookup.is_array()) << xlookup.debug_to_string();
  ASSERT_EQ(xlookup.as_array_rows(), 2U);
  ASSERT_EQ(xlookup.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(xlookup.as_array_cells()[0].as_number(), 300.0);
  EXPECT_DOUBLE_EQ(xlookup.as_array_cells()[1].as_number(), 100.0);
  ASSERT_TRUE(xlookup.as_array_cells()[2].is_text());
  EXPECT_EQ(xlookup.as_array_cells()[2].as_text(), "missing");
  EXPECT_DOUBLE_EQ(xlookup.as_array_cells()[3].as_number(), 200.0);

  const Value xmatch = EvalSourceIn("=XMATCH({30,10;99,20},A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(xmatch.is_array()) << xmatch.debug_to_string();
  ASSERT_EQ(xmatch.as_array_rows(), 2U);
  ASSERT_EQ(xmatch.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(xmatch.as_array_cells()[0].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(xmatch.as_array_cells()[1].as_number(), 1.0);
  ASSERT_TRUE(xmatch.as_array_cells()[2].is_error());
  EXPECT_EQ(xmatch.as_array_cells()[2].as_error(), ErrorCode::NA);
  EXPECT_DOUBLE_EQ(xmatch.as_array_cells()[3].as_number(), 2.0);
}

TEST(BuiltinsXLookup, ArrayLookupValueUsesFirstCellOfMultiCellReturnSlice) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(101.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(1001.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(202.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(2002.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(303.0));
  wb.sheet(0).set_cell_value(2, 2, Value::number(3003.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(30.0));  // D1
  wb.sheet(0).set_cell_value(1, 3, Value::number(10.0));  // D2
  const Value vertical = EvalSourceIn("=XLOOKUP(D1:D2,A1:A3,B1:C3,\"missing\")", wb, wb.sheet(0));
  ASSERT_TRUE(vertical.is_array()) << vertical.debug_to_string();
  EXPECT_EQ(vertical.as_array_rows(), 2U);
  EXPECT_EQ(vertical.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(vertical.as_array_cells()[0].as_number(), 303.0);
  EXPECT_DOUBLE_EQ(vertical.as_array_cells()[1].as_number(), 101.0);

  wb.sheet(0).set_cell_value(4, 0, Value::number(10.0));  // A5
  wb.sheet(0).set_cell_value(4, 1, Value::number(20.0));  // B5
  wb.sheet(0).set_cell_value(4, 2, Value::number(30.0));  // C5
  wb.sheet(0).set_cell_value(5, 0, Value::number(101.0));
  wb.sheet(0).set_cell_value(5, 1, Value::number(202.0));
  wb.sheet(0).set_cell_value(5, 2, Value::number(303.0));
  wb.sheet(0).set_cell_value(6, 0, Value::number(1001.0));
  wb.sheet(0).set_cell_value(6, 1, Value::number(2002.0));
  wb.sheet(0).set_cell_value(6, 2, Value::number(3003.0));
  wb.sheet(0).set_cell_value(3, 3, Value::number(20.0));  // D4
  wb.sheet(0).set_cell_value(3, 4, Value::number(10.0));  // E4
  const Value horizontal = EvalSourceIn("=XLOOKUP(D4:E4,A5:C5,A6:C7,\"missing\")", wb, wb.sheet(0));
  ASSERT_TRUE(horizontal.is_array()) << horizontal.debug_to_string();
  EXPECT_EQ(horizontal.as_array_rows(), 1U);
  EXPECT_EQ(horizontal.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(horizontal.as_array_cells()[0].as_number(), 202.0);
  EXPECT_DOUBLE_EQ(horizontal.as_array_cells()[1].as_number(), 101.0);
}

TEST(BuiltinsXLookup, ArrayLookupReferenceBlankPromotionAndScalarControls) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));  // A1
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));  // A2
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));  // A3
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));  // B1
  // B2 intentionally remains blank.
  wb.sheet(0).set_cell_value(2, 1, Value::number(1.0));  // B3
  const Value counta = EvalSourceIn("=COUNTA(XLOOKUP(SEQUENCE(3),A1:A3,B1:B3))", wb, wb.sheet(0));
  ASSERT_TRUE(counta.is_number());
  EXPECT_DOUBLE_EQ(counta.as_number(), 3.0);
  const Value count = EvalSourceIn("=COUNT(XLOOKUP(SEQUENCE(3),A1:A3,B1:B3))", wb, wb.sheet(0));
  ASSERT_TRUE(count.is_number());
  EXPECT_DOUBLE_EQ(count.as_number(), 2.0);
  const Value nested_blank = EvalSourceIn("=ISBLANK(INDEX(XLOOKUP(SEQUENCE(3),A1:A3,B1:B3),2,1))", wb, wb.sheet(0));
  ASSERT_TRUE(nested_blank.is_boolean());
  EXPECT_TRUE(nested_blank.as_boolean());

  const Value scalar_counta = EvalSourceIn("=COUNTA(XLOOKUP(2,A1:A3,B1:B3))", wb, wb.sheet(0));
  ASSERT_TRUE(scalar_counta.is_number());
  EXPECT_DOUBLE_EQ(scalar_counta.as_number(), 0.0);

  wb.sheet(0).set_cell_value(0, 3, Value::number(1.0));   // D1
  wb.sheet(0).set_cell_value(0, 4, Value::number(2.0));   // E1
  wb.sheet(0).set_cell_value(0, 5, Value::number(3.0));   // F1
  wb.sheet(0).set_cell_value(1, 3, Value::number(10.0));  // D2
  // E2 intentionally remains blank.
  wb.sheet(0).set_cell_value(1, 5, Value::number(30.0));   // F2
  wb.sheet(0).set_cell_value(2, 3, Value::number(100.0));  // D3
  // E3 intentionally remains blank.
  wb.sheet(0).set_cell_value(2, 5, Value::number(300.0));  // F3

  const Value multi = EvalSourceIn("=XLOOKUP(D1:F1,D1:F1,D2:F3)", wb, wb.sheet(0));
  ASSERT_TRUE(multi.is_array()) << multi.debug_to_string();
  EXPECT_EQ(multi.as_array_rows(), 1U);
  EXPECT_EQ(multi.as_array_cols(), 3U);
  EXPECT_DOUBLE_EQ(multi.as_array_cells()[0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(multi.as_array_cells()[2].as_number(), 30.0);
  const Value multi_nested_blank = EvalSourceIn("=ISBLANK(INDEX(XLOOKUP(D1:F1,D1:F1,D2:F3),1,2))", wb, wb.sheet(0));
  ASSERT_TRUE(multi_nested_blank.is_boolean());
  EXPECT_TRUE(multi_nested_blank.as_boolean());
  const Value scalar_slice_counta = EvalSourceIn("=COUNTA(XLOOKUP(2,D1:F1,D2:F3))", wb, wb.sheet(0));
  ASSERT_TRUE(scalar_slice_counta.is_number());
  EXPECT_DOUBLE_EQ(scalar_slice_counta.as_number(), 0.0);

  // INDEX over G1:G1 supplies a reference-derived blank as if_not_found.
  // Array lanes promote it at the output boundary, while scalar XLOOKUP
  // preserves the raw blank.
  const Value array_fallback_counta =
      EvalSourceIn("=COUNTA(XLOOKUP({1;2;99},A1:A3,B1:B3,INDEX(G1:G1,1)))", wb, wb.sheet(0));
  ASSERT_TRUE(array_fallback_counta.is_number());
  EXPECT_DOUBLE_EQ(array_fallback_counta.as_number(), 3.0);
  const Value scalar_fallback_counta = EvalSourceIn("=COUNTA(XLOOKUP(99,A1:A3,B1:B3,INDEX(G1:G1,1)))", wb, wb.sheet(0));
  ASSERT_TRUE(scalar_fallback_counta.is_number());
  EXPECT_DOUBLE_EQ(scalar_fallback_counta.as_number(), 0.0);
}

TEST(BuiltinsXLookup, ArrayLookupValueErrorsAndTrueMissesAreIndependent) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0)}, {Value::number(100.0), Value::number(200.0)});
  wb.sheet(0).set_cell_value(0, 3, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 3, Value::error(ErrorCode::Div0));
  wb.sheet(0).set_cell_value(2, 3, Value::number(99.0));
  const Value v = EvalSourceIn("=XLOOKUP(D1:D3,A1:A2,B1:B2,\"missing\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 100.0);
  ASSERT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(v.as_array_cells()[1].as_error(), ErrorCode::Div0);
  ASSERT_TRUE(v.as_array_cells()[2].is_text());
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "missing");
}

TEST(BuiltinsXLookup, ArrayLookupValueAllHitsDoNotEvaluateFallback) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0)}, {Value::number(100.0), Value::number(200.0)});
  wb.sheet(0).set_cell_value(0, 3, Value::number(20.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(10.0));
  const Value v = EvalSourceIn("=XLOOKUP(D1:D2,A1:A2,B1:B2,1/0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 200.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 100.0);
}

TEST(BuiltinsXLookup, ArrayFallbackIsValueErrorOnlyInMissLanes) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0)}, {Value::number(100.0), Value::number(200.0)});
  wb.sheet(0).set_cell_value(0, 2, Value::number(900.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(901.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(99.0));
  const Value v = EvalSourceIn("=XLOOKUP(D1:D2,A1:A2,B1:B2,C1:C2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 100.0);
  ASSERT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(v.as_array_cells()[1].as_error(), ErrorCode::Value);
}

TEST(BuiltinsXLookup, ScalarLookupArrayFallbackPreservesArrayResult) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0)}, {Value::number(100.0), Value::number(200.0)});
  wb.sheet(0).set_cell_value(0, 2, Value::number(900.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(901.0));
  const Value v = EvalSourceIn("=XLOOKUP(99,A1:A2,B1:B2,C1:C2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 900.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 901.0);
}

// ---------------------------------------------------------------------------
// XLOOKUP
// ---------------------------------------------------------------------------

TEST(BuiltinsXLookup, ExactNumericHit) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0)},
                 {Value::text("ten"), Value::text("twenty"), Value::text("thirty")});
  const Value v = EvalSourceIn("=XLOOKUP(20, A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "twenty");
}

TEST(BuiltinsXLookup, ExactTextCaseInsensitive) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::text("Apple"), Value::text("Banana"), Value::text("CHERRY")},
                 {Value::number(1.0), Value::number(2.0), Value::number(3.0)});
  const Value v = EvalSourceIn("=XLOOKUP(\"banana\", A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsXLookup, MatchModeSmallerExactHit) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0)},
                 {Value::text("a"), Value::text("b"), Value::text("c")});
  const Value v = EvalSourceIn("=XLOOKUP(20, A1:A3, B1:B3, \"nf\", -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsXLookup, MatchModeSmallerBetweenPicksBelow) {
  // Lookup 25 against {10, 20, 30} with match_mode=-1 returns the row for 20.
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0)},
                 {Value::text("a"), Value::text("b"), Value::text("c")});
  const Value v = EvalSourceIn("=XLOOKUP(25, A1:A3, B1:B3, \"nf\", -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsXLookup, MatchModeSmallerBelowFirstReturnsIfNotFound) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0)},
                 {Value::text("a"), Value::text("b"), Value::text("c")});
  const Value v = EvalSourceIn("=XLOOKUP(5, A1:A3, B1:B3, \"nf\", -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "nf");
}

TEST(BuiltinsXLookup, MatchModeLargerExactHit) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0)},
                 {Value::text("a"), Value::text("b"), Value::text("c")});
  const Value v = EvalSourceIn("=XLOOKUP(20, A1:A3, B1:B3, \"nf\", 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsXLookup, MatchModeLargerBetweenPicksAbove) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0)},
                 {Value::text("a"), Value::text("b"), Value::text("c")});
  const Value v = EvalSourceIn("=XLOOKUP(25, A1:A3, B1:B3, \"nf\", 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "c");
}

TEST(BuiltinsXLookup, MatchModeLargerAboveLastReturnsIfNotFound) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0)},
                 {Value::text("a"), Value::text("b"), Value::text("c")});
  const Value v = EvalSourceIn("=XLOOKUP(99, A1:A3, B1:B3, \"nf\", 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "nf");
}

TEST(BuiltinsXLookup, WildcardStar) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::text("Banana"), Value::text("Apple"), Value::text("Apricot")},
                 {Value::number(100.0), Value::number(200.0), Value::number(300.0)});
  // "A*" matches "Apple" first.
  const Value v = EvalSourceIn("=XLOOKUP(\"A*\", A1:A3, B1:B3, \"nf\", 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 200.0);
}

TEST(BuiltinsXLookup, WildcardQuestion) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::text("cab"), Value::text("ab"), Value::text("abc")},
                 {Value::number(1.0), Value::number(2.0), Value::number(3.0)});
  // "?b" matches the first 2-char string whose 2nd char is 'b' -> "ab".
  const Value v = EvalSourceIn("=XLOOKUP(\"?b\", A1:A3, B1:B3, \"nf\", 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsXLookup, WildcardEscapeLiteralStar) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::text("x"), Value::text("*"), Value::text("y")},
                 {Value::number(1.0), Value::number(2.0), Value::number(3.0)});
  // "~*" should match the literal "*" only.
  const Value v = EvalSourceIn("=XLOOKUP(\"~*\", A1:A3, B1:B3, \"nf\", 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsXLookup, SearchModeReverseDuplicates) {
  // {10, 20, 20, 30} — forward scan hits row 2, reverse scan hits row 3.
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(20.0), Value::number(30.0)},
                 {Value::text("a"), Value::text("b1"), Value::text("b2"), Value::text("c")});
  const Value fwd = EvalSourceIn("=XLOOKUP(20, A1:A4, B1:B4, \"nf\", 0, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(fwd.is_text());
  EXPECT_EQ(fwd.as_text(), "b1");
  const Value rev = EvalSourceIn("=XLOOKUP(20, A1:A4, B1:B4, \"nf\", 0, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(rev.is_text());
  EXPECT_EQ(rev.as_text(), "b2");
}

TEST(BuiltinsXLookup, OmittedModesUseDefaultAndExplicitReverseFindsLastDuplicate) {
  // With all optional slots omitted, the default forward search finds the
  // first duplicate. Supplying only search_mode=-1 finds the last duplicate.
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(20.0), Value::number(20.0), Value::number(30.0)},
                 {Value::text("first"), Value::text("last"), Value::text("other")});
  const Value default_search = EvalSourceIn("=XLOOKUP(20, A1:A3, B1:B3,,,)", wb, wb.sheet(0));
  ASSERT_TRUE(default_search.is_text()) << default_search.debug_to_string();
  EXPECT_EQ(default_search.as_text(), "first");
  const Value reverse_search = EvalSourceIn("=XLOOKUP(20, A1:A3, B1:B3,,,-1)", wb, wb.sheet(0));
  ASSERT_TRUE(reverse_search.is_text()) << reverse_search.debug_to_string();
  EXPECT_EQ(reverse_search.as_text(), "last");
}

TEST(BuiltinsXLookup, SearchModeBinaryAscExactHit) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0), Value::number(40.0)},
                 {Value::text("a"), Value::text("b"), Value::text("c"), Value::text("d")});
  const Value v = EvalSourceIn("=XLOOKUP(30, A1:A4, B1:B4, \"nf\", 0, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "c");
}

TEST(BuiltinsXLookup, SearchModeBinaryAscSmaller) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0), Value::number(40.0)},
                 {Value::text("a"), Value::text("b"), Value::text("c"), Value::text("d")});
  // Lookup 25 with match_mode=-1, binary ascending -> "b" (for key 20).
  const Value v = EvalSourceIn("=XLOOKUP(25, A1:A4, B1:B4, \"nf\", -1, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsXLookup, SearchModeBinaryAscLarger) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0), Value::number(30.0), Value::number(40.0)},
                 {Value::text("a"), Value::text("b"), Value::text("c"), Value::text("d")});
  // Lookup 25 with match_mode=1, binary ascending -> "c" (for key 30).
  const Value v = EvalSourceIn("=XLOOKUP(25, A1:A4, B1:B4, \"nf\", 1, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "c");
}

TEST(BuiltinsXLookup, SearchModeBinaryDescExactHit) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(40.0), Value::number(30.0), Value::number(20.0), Value::number(10.0)},
                 {Value::text("a"), Value::text("b"), Value::text("c"), Value::text("d")});
  const Value v = EvalSourceIn("=XLOOKUP(30, A1:A4, B1:B4, \"nf\", 0, -2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsXLookup, IfNotFoundDefaultIsNa) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0)}, {Value::text("a"), Value::text("b")});
  const Value v = EvalSourceIn("=XLOOKUP(99, A1:A2, B1:B2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsXLookup, IfNotFoundCustomScalar) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0)}, {Value::text("a"), Value::text("b")});
  const Value v = EvalSourceIn("=XLOOKUP(99, A1:A2, B1:B2, \"missing\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "missing");
}

TEST(BuiltinsXLookup, IfNotFoundNotEvaluatedOnHit) {
  // 1/0 lives in the if_not_found slot; on a hit it must NOT be evaluated,
  // so no #DIV/0! should leak through.
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0), Value::number(20.0)}, {Value::text("a"), Value::text("b")});
  const Value v = EvalSourceIn("=XLOOKUP(10, A1:A2, B1:B2, 1/0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_text(), "a");
}

TEST(BuiltinsXLookup, InvalidMatchModeIsValueError) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0)}, {Value::text("a")});
  EXPECT_EQ(EvalSourceIn("=XLOOKUP(10, A1:A1, B1:B1, \"nf\", 5)", wb, wb.sheet(0)).as_error(), ErrorCode::Value);
}

TEST(BuiltinsXLookup, InvalidSearchModeIsValueError) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0)}, {Value::text("a")});
  EXPECT_EQ(EvalSourceIn("=XLOOKUP(10, A1:A1, B1:B1, \"nf\", 0, 0)", wb, wb.sheet(0)).as_error(), ErrorCode::Value);
}

TEST(BuiltinsXLookup, TwoDimensionalLookupArrayIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(0, 2, Value::text("r1"));
  wb.sheet(0).set_cell_value(1, 2, Value::text("r2"));
  EXPECT_EQ(EvalSourceIn("=XLOOKUP(1, A1:B2, C1:C2)", wb, wb.sheet(0)).as_error(), ErrorCode::Value);
}

TEST(BuiltinsXLookup, MultiColumnReturnSpillsMatchedRow) {
  // Vertical lookup on A1:A3, multi-column return B1:D3 -> matched row
  // spills across all three return columns.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  for (std::uint32_t r = 0; r < 3; ++r) {
    wb.sheet(0).set_cell_value(r, 1, Value::number(static_cast<double>(r * 10 + 1)));  // B
    wb.sheet(0).set_cell_value(r, 2, Value::number(static_cast<double>(r * 10 + 2)));  // C
    wb.sheet(0).set_cell_value(r, 3, Value::number(static_cast<double>(r * 10 + 3)));  // D
  }
  const Value v = EvalSourceIn("=XLOOKUP(20, A1:A3, B1:D3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 11.0);  // B2
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 12.0);  // C2
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 13.0);  // D2
}

TEST(BuiltinsXLookup, MultiRowReturnSpillsMatchedColumn) {
  // Horizontal lookup on A1:C1, multi-row return A2:C4 -> matched column
  // spills across all three return rows.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(30.0));
  for (std::uint32_t r = 1; r < 4; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r * 100 + 1)));
    wb.sheet(0).set_cell_value(r, 1, Value::number(static_cast<double>(r * 100 + 2)));
    wb.sheet(0).set_cell_value(r, 2, Value::number(static_cast<double>(r * 100 + 3)));
  }
  const Value v = EvalSourceIn("=XLOOKUP(20, A1:C1, A2:C4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 102.0);  // row2 col B (match col=1)
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 202.0);  // row3 col B
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 302.0);  // row4 col B
}

TEST(BuiltinsXLookup, MultiColumnReturnAxisMismatchIsValueError) {
  // Vertical lookup length 3 but return_array has 2 rows -> #VALUE!.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(4.0));
  EXPECT_EQ(EvalSourceIn("=XLOOKUP(20, A1:A3, B1:C2)", wb, wb.sheet(0)).as_error(), ErrorCode::Value);
}

TEST(BuiltinsXLookup, RowLookupWithRowReturn) {
  // Both arrays are rows of length 3; lookup column orientation is flipped
  // but flat-index translation still picks the right cell.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(30.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 2, Value::text("c"));
  const Value v = EvalSourceIn("=XLOOKUP(20, A1:C1, A2:C2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsXLookup, ReturnArraySizeMismatchIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("b"));
  EXPECT_EQ(EvalSourceIn("=XLOOKUP(10, A1:A3, B1:B2)", wb, wb.sheet(0)).as_error(), ErrorCode::Value);
}

TEST(BuiltinsXLookup, CrossSheetRanges) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(0, 0, Value::text("alpha"));
  wb.sheet(1).set_cell_value(1, 0, Value::text("beta"));
  wb.sheet(1).set_cell_value(2, 0, Value::text("gamma"));
  wb.sheet(1).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(1).set_cell_value(2, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=XLOOKUP(\"beta\", Data!A1:A3, Data!B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsXLookup, LookupValueErrorPropagates) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0)}, {Value::text("a")});
  EXPECT_EQ(EvalSourceIn("=XLOOKUP(#DIV/0!, A1:A1, B1:B1)", wb, wb.sheet(0)).as_error(), ErrorCode::Div0);
}

TEST(BuiltinsXLookup, ReturnCellErrorIsForwarded) {
  // The return cell contains #DIV/0!; XLOOKUP must propagate it unchanged.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 1, Value::error(ErrorCode::Div0));
  const Value v = EvalSourceIn("=XLOOKUP(10, A1:A1, B1:B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsXLookup, ArityTooFewIsValueError) {
  EXPECT_EQ(EvalSource("=XLOOKUP(1)").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=XLOOKUP(1, 2)").as_error(), ErrorCode::Value);
}

TEST(BuiltinsXLookup, ArityTooManyIsValueError) {
  Workbook wb = Workbook::create();
  SeedColumnPair(wb, {Value::number(10.0)}, {Value::text("a")});
  EXPECT_EQ(EvalSourceIn("=XLOOKUP(10, A1:A1, B1:B1, \"nf\", 0, 1, 99)", wb, wb.sheet(0)).as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// XMATCH
// ---------------------------------------------------------------------------

TEST(BuiltinsXMatch, ExactNumeric) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  const Value v = EvalSourceIn("=XMATCH(20, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsXMatch, ArrayLookupValuePreservesShapeAndErrors) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(20.0));
  wb.sheet(0).set_cell_value(1, 3, Value::error(ErrorCode::Div0));
  wb.sheet(0).set_cell_value(2, 3, Value::number(99.0));
  const Value vertical = EvalSourceIn("=XMATCH(D1:D3,A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(vertical.is_array()) << vertical.debug_to_string();
  EXPECT_EQ(vertical.as_array_rows(), 3U);
  EXPECT_EQ(vertical.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(vertical.as_array_cells()[0].as_number(), 2.0);
  ASSERT_TRUE(vertical.as_array_cells()[1].is_error());
  EXPECT_EQ(vertical.as_array_cells()[1].as_error(), ErrorCode::Div0);
  ASSERT_TRUE(vertical.as_array_cells()[2].is_error());
  EXPECT_EQ(vertical.as_array_cells()[2].as_error(), ErrorCode::NA);

  const Value literal = EvalSourceIn("=XMATCH({30;10},A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(literal.is_array()) << literal.debug_to_string();
  EXPECT_EQ(literal.as_array_rows(), 2U);
  EXPECT_EQ(literal.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(literal.as_array_cells()[0].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(literal.as_array_cells()[1].as_number(), 1.0);
}

TEST(BuiltinsXMatch, ExactTextCaseInsensitive) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("Apple"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("Banana"));
  const Value v = EvalSourceIn("=XMATCH(\"APPLE\", A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsXMatch, SmallerBetweenPicksBelow) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  const Value v = EvalSourceIn("=XMATCH(25, A1:A3, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsXMatch, LargerBetweenPicksAbove) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  const Value v = EvalSourceIn("=XMATCH(25, A1:A3, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsXMatch, WildcardMode) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("banana"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("apple"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("apricot"));
  const Value v = EvalSourceIn("=XMATCH(\"ap*\", A1:A3, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsXMatch, ReverseSearchDuplicates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(30.0));
  const Value fwd = EvalSourceIn("=XMATCH(20, A1:A4, 0, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(fwd.is_number());
  EXPECT_DOUBLE_EQ(fwd.as_number(), 2.0);
  const Value rev = EvalSourceIn("=XMATCH(20, A1:A4, 0, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(rev.is_number());
  EXPECT_DOUBLE_EQ(rev.as_number(), 3.0);
}

TEST(BuiltinsXMatch, OmittedModesUseDefaultAndExplicitReverseFindsLastDuplicate) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  const Value default_search = EvalSourceIn("=XMATCH(20, A1:A3,,)", wb, wb.sheet(0));
  ASSERT_TRUE(default_search.is_number()) << default_search.debug_to_string();
  EXPECT_DOUBLE_EQ(default_search.as_number(), 1.0);
  const Value reverse_search = EvalSourceIn("=XMATCH(20, A1:A3,,-1)", wb, wb.sheet(0));
  ASSERT_TRUE(reverse_search.is_number()) << reverse_search.debug_to_string();
  EXPECT_DOUBLE_EQ(reverse_search.as_number(), 2.0);
}

TEST(BuiltinsXMatch, BinaryAscExactHit) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(40.0));
  const Value v = EvalSourceIn("=XMATCH(30, A1:A4, 0, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsXMatch, BinaryDescExactHit) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(40.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=XMATCH(30, A1:A4, 0, -2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsXMatch, NoMatchIsNa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  const Value v = EvalSourceIn("=XMATCH(99, A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsXMatch, InvalidMatchModeIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  EXPECT_EQ(EvalSourceIn("=XMATCH(10, A1:A1, 7)", wb, wb.sheet(0)).as_error(), ErrorCode::Value);
}

TEST(BuiltinsXMatch, InvalidSearchModeIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  EXPECT_EQ(EvalSourceIn("=XMATCH(10, A1:A1, 0, 3)", wb, wb.sheet(0)).as_error(), ErrorCode::Value);
}

TEST(BuiltinsXMatch, TwoDimensionalArrayIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  EXPECT_EQ(EvalSourceIn("=XMATCH(1, A1:B2)", wb, wb.sheet(0)).as_error(), ErrorCode::Value);
}

TEST(BuiltinsXMatch, CrossSheet) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(0, 0, Value::text("alpha"));
  wb.sheet(1).set_cell_value(1, 0, Value::text("beta"));
  wb.sheet(1).set_cell_value(2, 0, Value::text("gamma"));
  const Value v = EvalSourceIn("=XMATCH(\"gamma\", Data!A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsXMatch, ArityTooFewIsValueError) {
  EXPECT_EQ(EvalSource("=XMATCH(1)").as_error(), ErrorCode::Value);
}

TEST(BuiltinsXMatch, LookupValueErrorPropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  EXPECT_EQ(EvalSourceIn("=XMATCH(#DIV/0!, A1:A1)", wb, wb.sheet(0)).as_error(), ErrorCode::Div0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
