// Copyright 2026 libraz. Licensed under the MIT License.
//
// Evaluator tests for the spilled-range `=A1#` operator.
//
// The parser produces `NodeKind::SpillRef`; `eval_node` resolves it via
// `Sheet::spill_region_at_anchor` and returns a `Value::Array`. Unlike
// SEQUENCE (which produces a Value::Array via `dispatch_array_result`),
// SpillRef reads an *already-committed* spill — these tests pre-seed the
// sheet's spill table directly and then evaluate the formula.

#include <cstdint>
#include <string_view>
#include <vector>

#include "cell.h"
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

// Parses `src`, evaluates it under `ctx`, returns the resulting `Value`.
// Both arenas live on the caller's stack so any text / array payloads
// remain readable for assertions.
Value EvalUnder(std::string_view src, Arena* parse_arena, Arena* eval_arena, const EvalContext& ctx) {
  parser::Parser p(src, *parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, *eval_arena, default_registry(), ctx);
}

TEST(SpillRefEval, ResolvesCommittedSpill) {
  // Pre-commit a 3x1 spill at A1 with cells [10, 20, 30]; `=A1#` should
  // return a Value::Array with the same shape and cells.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells{Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  const Value* read_cells = v.as_array_cells();
  EXPECT_EQ(read_cells[0], Value::number(10.0));
  EXPECT_EQ(read_cells[1], Value::number(20.0));
  EXPECT_EQ(read_cells[2], Value::number(30.0));
}

TEST(SpillRefEval, NoSpillReturnsRef) {
  // No spill anchored at A1 -> `#REF!`. Mac Excel behaves this way: the `#`
  // operator only resolves to a real spill region; querying a cell that
  // never spilled is an explicit reference error.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(SpillRefEval, SumOverSpillRef) {
  // SUM consumes a SpillRef like any other range. Pre-commit [10, 20, 30]
  // at A1 and call `=SUM(A1#)`; result is 60.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells{Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SUM(A1#)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 60.0);
}

TEST(SpillRefEval, SpillRefArithmeticParsesAndEvaluates) {
  // `=A1#+B1#` over two 1x1 spills: the parser accepts the form and
  // evaluation reaches the BinaryOp. The scalar arithmetic path does not
  // yet implement array-context broadcasting, so the result is currently
  // `#VALUE!` (Array operand fails `coerce_to_number`). Mac Excel would
  // spill this to a single-cell array of 50; closing that gap is a
  // follow-up that needs array-aware BinaryOp dispatch — tracked outside
  // this test. What we lock in here is the parser shape and the absence
  // of any crash / wrong-error path.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 1U, 1U, std::vector<Value>{Value::number(42.0)}));
  ASSERT_TRUE(sheet.commit_spill(0U, 1U, 1U, 1U, std::vector<Value>{Value::number(8.0)}));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#+B1#", &parse_arena, &eval_arena, ctx);
  if (v.is_number()) {
    EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
  } else if (v.is_array()) {
    ASSERT_EQ(v.as_array_rows(), 1U);
    ASSERT_EQ(v.as_array_cols(), 1U);
    EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 50.0);
  } else {
    ASSERT_TRUE(v.is_error());
    EXPECT_EQ(v.as_error(), ErrorCode::Value);
  }
}

TEST(SpillRefEval, SpillRefAnchorOnPhantomReturnsRef) {
  // Pre-commit a 3x1 spill at A1; the cells at A2 / A3 are phantoms of the
  // anchor. `=A2#` queries a phantom address (not the anchor) and must
  // return `#REF!` because only anchors carry the spill region.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells{Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A2#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(SpillRefEval, UnboundContextReturnsName) {
  // Without a current sheet bound to the EvalContext there is nothing to
  // query; SpillRef surfaces `#NAME?` to mirror the corresponding Ref
  // resolution path.
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#", &parse_arena, &eval_arena, EvalContext{});
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(SpillRefEval, QualifiedSpillRef) {
  // `=Sheet2!A1#` resolves the anchor against Sheet2's spill table even
  // though the EvalContext's `current_sheet_` is Sheet1.
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  Sheet& s1 = wb.sheet(0);
  Sheet& s2 = wb.sheet(1);
  ASSERT_TRUE(s2.commit_spill(0U, 0U, 2U, 1U, std::vector<Value>{Value::number(7.0), Value::number(11.0)}));

  EvalState state;
  const EvalContext ctx(wb, s1, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=Sheet2!A1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0], Value::number(7.0));
  EXPECT_EQ(v.as_array_cells()[1], Value::number(11.0));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
