// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the lazy-dispatched lookup / reference functions
// CHOOSE, INDEX, and MATCH. All three are routed through the lazy table in
// `tree_walker.cpp` because they need AST-level inspection of their
// arguments (range shape for INDEX / MATCH; per-branch short-circuit for
// CHOOSE).

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
// MATCH
// ---------------------------------------------------------------------------

TEST(BuiltinsMatch, ExactNumericFirstHit) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(20.0));  // duplicate, should be ignored
  const Value v = EvalSourceIn("=MATCH(20, A1:A4, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactTextCaseInsensitive) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("apple"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("Banana"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("CHERRY"));
  const Value v = EvalSourceIn("=MATCH(\"banana\", A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactWildcardStar) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("Banana"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("Apple"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("Apricot"));
  const Value v = EvalSourceIn("=MATCH(\"A*\", A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactWildcardQuestion) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("cab"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("ab"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("abc"));
  // "?b" matches exactly one byte before 'b' -> "ab" doesn't match (no
  // leading byte) but "cb" would; we want "?b" to match a 2-byte string
  // whose 2nd char is 'b'. Only "ab" is length 2 with b at index 1, which
  // matches "?b" (? eats "a"). "cab" is 3 bytes so no match.
  const Value v = EvalSourceIn("=MATCH(\"?b\", A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactWildcardEscape) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("foo"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("*"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("bar"));
  // "~*" is the literal asterisk.
  const Value v = EvalSourceIn("=MATCH(\"~*\", A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactNoMatchIsNa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=MATCH(99, A1:A2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, AscendingExactHit) {
  // Ascending array, target exactly matches a cell.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=MATCH(5, A1:A3, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, AscendingBetweenCells) {
  // Target falls between cells: return the position whose value is
  // largest but still <= target.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=MATCH(7, A1:A3, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, AscendingAllGreaterIsNa) {
  // Every cell is strictly greater than the target -> #N/A.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(20.0));
  const Value v = EvalSourceIn("=MATCH(1, A1:A3, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, DescendingSmallestGreaterOrEqual) {
  // Descending array: return largest position whose value is >= target.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=MATCH(15, A1:A3, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, DescendingAllLessIsNa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=MATCH(10, A1:A3, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, TwoDimensionalRangeIsNa) {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 2; ++r) {
    for (std::uint32_t c = 0; c < 2; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 2 + c)));
    }
  }
  const Value v = EvalSourceIn("=MATCH(1, A1:B2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, OmittedMatchTypeDefaultsToAscending) {
  // Default match_type is 1; omitting it should behave identically.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=MATCH(7, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, LookupValueErrorPropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=MATCH(#DIV/0!, A1:A1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMatch, SingleCellRefAsLookupArray) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value hit = EvalSourceIn("=MATCH(42, A1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(hit.is_number());
  EXPECT_DOUBLE_EQ(hit.as_number(), 1.0);
  const Value miss = EvalSourceIn("=MATCH(99, A1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(miss.is_error());
  EXPECT_EQ(miss.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, CrossSheetLookupArray) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(0, 0, Value::text("alpha"));
  wb.sheet(1).set_cell_value(1, 0, Value::text("beta"));
  wb.sheet(1).set_cell_value(2, 0, Value::text("gamma"));
  const Value v = EvalSourceIn("=MATCH(\"beta\", Data!A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, InvalidMatchTypeIsNa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=MATCH(1, A1:A1, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, WrongArityIsValueError) {
  EXPECT_EQ(EvalSource("=MATCH(1)").as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// INDEX
// ---------------------------------------------------------------------------

TEST(BuiltinsIndex, ColumnRangeRowSelects) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(40.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(50.0));
  const Value v = EvalSourceIn("=INDEX(A1:A5, 3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsIndex, RowRangeSoleIndexPicksColumn) {
  // 1-D row vector A1:E1 with two args (sole index = column).
  Workbook wb = Workbook::create();
  for (std::uint32_t c = 0; c < 5; ++c) {
    wb.sheet(0).set_cell_value(0, c, Value::number(10.0 * static_cast<double>(c + 1)));
  }
  const Value v = EvalSourceIn("=INDEX(A1:E1, 4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 40.0);
}

TEST(BuiltinsIndex, TwoDimensionalRowAndColumn) {
  Workbook wb = Workbook::create();
  // Populate A1:C3 with distinct values 1..9 row-major.
  for (std::uint32_t r = 0; r < 3; ++r) {
    for (std::uint32_t c = 0; c < 3; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 3 + c + 1)));
    }
  }
  const Value v = EvalSourceIn("=INDEX(A1:C3, 2, 3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);  // row 2, col 3 -> cell (1,2) -> 1*3+2+1 = 6
}

TEST(BuiltinsIndex, OutOfBoundsRowIsRefError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=INDEX(A1:A2, 5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndex, OutOfBoundsColumnIsRefError) {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 2; ++r) {
    for (std::uint32_t c = 0; c < 2; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(1.0));
    }
  }
  const Value v = EvalSourceIn("=INDEX(A1:B2, 1, 5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndex, NegativeIndexIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=INDEX(A1:A1, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIndex, ZeroRowOn2DRangeIsValueError) {
  // Accepted divergence: Excel 365 would spill the whole column; we don't
  // have scalar spill results yet and document that in the impl.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 2; ++r) {
    for (std::uint32_t c = 0; c < 2; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(1.0));
    }
  }
  const Value v = EvalSourceIn("=INDEX(A1:B2, 0, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIndex, SingleCellRefAsArray) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("hello"));
  const Value hit = EvalSourceIn("=INDEX(A1, 1, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(hit.is_text());
  EXPECT_EQ(hit.as_text(), "hello");
  // Out of bounds on a single-cell Ref.
  const Value miss = EvalSourceIn("=INDEX(A1, 2, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(miss.is_error());
  EXPECT_EQ(miss.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsIndex, CrossSheetTwoDimensionalRange) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  for (std::uint32_t r = 0; r < 3; ++r) {
    for (std::uint32_t c = 0; c < 3; ++c) {
      wb.sheet(1).set_cell_value(r, c, Value::number(static_cast<double>((r + 1) * 10 + (c + 1))));
    }
  }
  const Value v = EvalSourceIn("=INDEX(Data!A1:C3, 3, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 32.0);
}

TEST(BuiltinsIndex, NonCoercibleIndexIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=INDEX(A1:A1, \"not a number\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIndex, WrongArityIsValueError) {
  EXPECT_EQ(EvalSource("=INDEX()").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=INDEX(1)").as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// CHOOSE
// ---------------------------------------------------------------------------

TEST(BuiltinsChoose, BasicSelect) {
  const Value v = EvalSource("=CHOOSE(2, \"a\", \"b\", \"c\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsChoose, FractionalIndexTruncates) {
  // 2.9 -> floor to 2, selects "b" (not rounding to 3).
  const Value v = EvalSource("=CHOOSE(2.9, \"a\", \"b\", \"c\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsChoose, OutOfRangeZeroIsValueError) {
  const Value v = EvalSource("=CHOOSE(0, \"a\", \"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChoose, OutOfRangeTooLargeIsValueError) {
  const Value v = EvalSource("=CHOOSE(4, \"a\", \"b\", \"c\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChoose, IndexErrorPropagates) {
  const Value v = EvalSource("=CHOOSE(#DIV/0!, \"a\", \"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsChoose, ChosenArgErrorPropagates) {
  const Value v = EvalSource("=CHOOSE(2, 1, #N/A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsChoose, ChosenArgIsCellRef) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(11.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(22.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(33.0));
  const Value v = EvalSourceIn("=CHOOSE(3, A1, A2, A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 33.0);
}

TEST(BuiltinsChoose, LargeArity) {
  // Ten values; pick the 7th.
  const Value v = EvalSource("=CHOOSE(7, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsChoose, UnselectedArgNotEvaluated) {
  // Non-chosen arg contains a #DIV/0! literal that would propagate if
  // evaluated. CHOOSE must short-circuit so the result is "a".
  const Value v = EvalSource("=CHOOSE(1, \"a\", #DIV/0!)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a");
}

TEST(BuiltinsChoose, ZeroArityIsValueError) {
  EXPECT_EQ(EvalSource("=CHOOSE()").as_error(), ErrorCode::Value);
  // Just an index without any values is also invalid.
  EXPECT_EQ(EvalSource("=CHOOSE(1)").as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
