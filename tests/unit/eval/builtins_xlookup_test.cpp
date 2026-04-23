// Copyright 2026 libraz. Licensed under the MIT License.
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
