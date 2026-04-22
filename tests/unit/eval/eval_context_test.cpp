// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `EvalContext::resolve_ref`. These tests exercise the
// semantics table documented on the header directly — constructing
// `parser::Reference` objects by hand so the parser, arena, and AST are not
// on the critical path.

#include "eval/eval_context.h"

#include <cstdint>
#include <string>

#include "gtest/gtest.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Builds a local A1 reference at (row, col). Defaults match the parser's
// output for a plain, unqualified, not-absolute A1 lexeme.
parser::Reference MakeLocalRef(std::uint32_t row, std::uint32_t col) {
  parser::Reference ref;
  ref.row = row;
  ref.col = col;
  return ref;
}

TEST(EvalContextResolve, NoBoundSheet_ReturnsName) {
  const EvalContext ctx;
  const Value v = ctx.resolve_ref(MakeLocalRef(0, 0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(EvalContextResolve, MissingCell_ReturnsBlank) {
  Sheet sheet("Sheet1");
  const EvalContext ctx(sheet);
  const Value v = ctx.resolve_ref(MakeLocalRef(0, 0));
  EXPECT_TRUE(v.is_blank());
}

TEST(EvalContextResolve, LiteralNumberCell_ReturnsNumber) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(42.0));
  const EvalContext ctx(sheet);
  const Value v = ctx.resolve_ref(MakeLocalRef(0, 0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(EvalContextResolve, LiteralTextCell_ReturnsText) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(1, 2, Value::text("hello"));
  const EvalContext ctx(sheet);
  const Value v = ctx.resolve_ref(MakeLocalRef(1, 2));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello");
}

TEST(EvalContextResolve, LiteralBoolCell_ReturnsBool) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(3, 4, Value::boolean(true));
  const EvalContext ctx(sheet);
  const Value v = ctx.resolve_ref(MakeLocalRef(3, 4));
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(EvalContextResolve, FormulaCellWithBlankCache_ReturnsBlank) {
  Sheet sheet("Sheet1");
  // Formula cells store blank in cached_value until the evaluator populates
  // them (a later increment); this resolve_ref hands back the cache verbatim.
  sheet.set_cell_formula(0, 0, "=1+1");
  const EvalContext ctx(sheet);
  const Value v = ctx.resolve_ref(MakeLocalRef(0, 0));
  EXPECT_TRUE(v.is_blank());
}

TEST(EvalContextResolve, CrossSheet_ReturnsRef) {
  Sheet sheet("Sheet1");
  const EvalContext ctx(sheet);
  parser::Reference ref = MakeLocalRef(0, 0);
  ref.sheet = std::string_view("Other");
  const Value v = ctx.resolve_ref(ref);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(EvalContextResolve, FullColumn_ReturnsValue) {
  Sheet sheet("Sheet1");
  const EvalContext ctx(sheet);
  parser::Reference ref = MakeLocalRef(0, 0);
  ref.is_full_col = true;
  const Value v = ctx.resolve_ref(ref);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(EvalContextResolve, FullRow_ReturnsValue) {
  Sheet sheet("Sheet1");
  const EvalContext ctx(sheet);
  parser::Reference ref = MakeLocalRef(0, 0);
  ref.is_full_row = true;
  const Value v = ctx.resolve_ref(ref);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(EvalContextResolve, RowAtMax_ReturnsRef) {
  Sheet sheet("Sheet1");
  const EvalContext ctx(sheet);
  const Value v = ctx.resolve_ref(MakeLocalRef(Sheet::kMaxRows, 0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(EvalContextResolve, ColAtMax_ReturnsRef) {
  Sheet sheet("Sheet1");
  const EvalContext ctx(sheet);
  const Value v = ctx.resolve_ref(MakeLocalRef(0, Sheet::kMaxCols));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(EvalContextResolve, CurrentSheetAccessor_ReflectsBinding) {
  const EvalContext unbound;
  EXPECT_EQ(unbound.current_sheet(), nullptr);

  Sheet sheet("Sheet1");
  const EvalContext bound(sheet);
  EXPECT_EQ(bound.current_sheet(), &sheet);
}

// ---------------------------------------------------------------------------
// Recursive overload (`resolve_ref` + EvalState)
// ---------------------------------------------------------------------------

// Resolves `(row, col)` on `sheet` with the supplied recursive `state`,
// using a fresh-per-call arena to keep text payloads alive for the caller.
Value ResolveWithState(const Sheet& sheet, EvalState& state,
                       std::uint32_t row, std::uint32_t col) {
  static thread_local Arena arena;
  arena.reset();
  const eval::EvalContext ctx(sheet, state);
  const parser::Reference ref = MakeLocalRef(row, col);
  return ctx.resolve_ref(ref, arena, eval::default_registry());
}

TEST(EvalContextRecursive, FormulaCellResolvesToResult) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=1+2");
  EvalState state;
  const Value v = ResolveWithState(sheet, state, 0, 0);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(EvalContextRecursive, FormulaReferencingLiteralCell) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(5.0));
  sheet.set_cell_formula(1, 0, "=A1+1");
  EvalState state;
  const Value v = ResolveWithState(sheet, state, 1, 0);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(EvalContextRecursive, FormulaChain) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(5.0));
  sheet.set_cell_formula(1, 0, "=A1+1");
  sheet.set_cell_formula(2, 0, "=A2*2");
  EvalState state;
  const Value v = ResolveWithState(sheet, state, 2, 0);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

TEST(EvalContextRecursive, DirectCycleReturnsRef) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=A1");
  EvalState state;
  const Value v = ResolveWithState(sheet, state, 0, 0);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(EvalContextRecursive, IndirectCycleReturnsRef) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=A2");
  sheet.set_cell_formula(1, 0, "=A1");
  EvalState state;
  const Value v = ResolveWithState(sheet, state, 0, 0);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(EvalContextRecursive, ThreeCellCycle) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=A2");
  sheet.set_cell_formula(1, 0, "=A3");
  sheet.set_cell_formula(2, 0, "=A1");
  EvalState state;
  const Value v = ResolveWithState(sheet, state, 0, 0);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(EvalContextRecursive, FormulaDiv0Propagates) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=1/0");
  EvalState state;
  const Value v = ResolveWithState(sheet, state, 0, 0);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(EvalContextRecursive, FormulaParseErrorReturnsName) {
  Sheet sheet("Sheet1");
  // Panic-mode recovery substitutes an ErrorPlaceholder; the evaluator
  // converts that to #NAME? at evaluation time.
  sheet.set_cell_formula(0, 0, "=???");
  EvalState state;
  const Value v = ResolveWithState(sheet, state, 0, 0);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(EvalContextRecursive, MemoizationReusesResult) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=42");
  EvalState state;
  // First call populates the memo.
  const Value v1 = ResolveWithState(sheet, state, 0, 0);
  ASSERT_TRUE(v1.is_number());
  EXPECT_EQ(v1.as_number(), 42.0);
  const Value* memoised = state.lookup_memo(&sheet, 0, 0);
  ASSERT_NE(memoised, nullptr);
  EXPECT_EQ(memoised->as_number(), 42.0);

  // Second call hits the memo and still returns the same result.
  const Value v2 = ResolveWithState(sheet, state, 0, 0);
  ASSERT_TRUE(v2.is_number());
  EXPECT_EQ(v2.as_number(), 42.0);
}

TEST(EvalContextRecursive, WithoutStateFormulaCellStillReturnsCachedValue) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=1+1");
  // 3-arg overload with state_ == nullptr must agree with the 1-arg form:
  // formula cell -> cached_value (blank by construction).
  Arena arena;
  const eval::EvalContext ctx(sheet);
  const Value v = ctx.resolve_ref(MakeLocalRef(0, 0), arena, eval::default_registry());
  EXPECT_TRUE(v.is_blank());
}

TEST(EvalContextRecursive, LiteralCellViaRecursiveOverload) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(7.0));
  EvalState state;
  const Value v = ResolveWithState(sheet, state, 0, 0);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
  // Literals are not pushed onto the stack and not memoised.
  EXPECT_EQ(state.depth(), 0U);
  EXPECT_EQ(state.lookup_memo(&sheet, 0, 0), nullptr);
}

TEST(EvalContextRecursive, NestedFormulaPreservesStackBalance) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(5.0));
  sheet.set_cell_formula(1, 0, "=A1+1");
  sheet.set_cell_formula(2, 0, "=A2*2");
  EvalState state;
  const Value v = ResolveWithState(sheet, state, 2, 0);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
  EXPECT_EQ(state.depth(), 0U);
}

// ---------------------------------------------------------------------------
// expand_range - rectangle flattening used by aggregator functions
// ---------------------------------------------------------------------------

// Helper: invoke `expand_range` on a literal-cell-only sheet (no EvalState
// needed). Uses a fresh-per-call arena to keep text payloads readable.
Expected<std::vector<Value>, ErrorCode> ExpandRange(const Sheet& sheet,
                                                    const parser::Reference& lhs,
                                                    const parser::Reference& rhs) {
  static thread_local Arena arena;
  arena.reset();
  const eval::EvalContext ctx(sheet);
  return ctx.expand_range(lhs, rhs, arena, eval::default_registry());
}

TEST(EvalContextExpandRange, UnboundContext_ReturnsName) {
  const eval::EvalContext ctx;
  static thread_local Arena arena;
  arena.reset();
  auto result = ctx.expand_range(MakeLocalRef(0, 0), MakeLocalRef(2, 0), arena,
                                 eval::default_registry());
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error(), ErrorCode::Name);
}

TEST(EvalContextExpandRange, SingleCell_ReturnsOne) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(42.0));
  auto result = ExpandRange(sheet, MakeLocalRef(0, 0), MakeLocalRef(0, 0));
  ASSERT_TRUE(result);
  ASSERT_EQ(result.value().size(), 1u);
  ASSERT_TRUE(result.value()[0].is_number());
  EXPECT_EQ(result.value()[0].as_number(), 42.0);
}

TEST(EvalContextExpandRange, ThreeCellColumn_ReturnsThreeInRowMajor) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(1, 0, Value::number(2.0));
  sheet.set_cell_value(2, 0, Value::number(3.0));
  auto result = ExpandRange(sheet, MakeLocalRef(0, 0), MakeLocalRef(2, 0));
  ASSERT_TRUE(result);
  ASSERT_EQ(result.value().size(), 3u);
  EXPECT_EQ(result.value()[0].as_number(), 1.0);
  EXPECT_EQ(result.value()[1].as_number(), 2.0);
  EXPECT_EQ(result.value()[2].as_number(), 3.0);
}

TEST(EvalContextExpandRange, TwoByTwo_RowMajorOrder) {
  Sheet sheet("Sheet1");
  // A1=1, B1=2, A2=3, B2=4. Row-major expansion yields [A1, B1, A2, B2].
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(0, 1, Value::number(2.0));
  sheet.set_cell_value(1, 0, Value::number(3.0));
  sheet.set_cell_value(1, 1, Value::number(4.0));
  auto result = ExpandRange(sheet, MakeLocalRef(0, 0), MakeLocalRef(1, 1));
  ASSERT_TRUE(result);
  ASSERT_EQ(result.value().size(), 4u);
  EXPECT_EQ(result.value()[0].as_number(), 1.0);
  EXPECT_EQ(result.value()[1].as_number(), 2.0);
  EXPECT_EQ(result.value()[2].as_number(), 3.0);
  EXPECT_EQ(result.value()[3].as_number(), 4.0);
}

TEST(EvalContextExpandRange, ReversedEndpoints_Normalised) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(1, 0, Value::number(2.0));
  sheet.set_cell_value(2, 0, Value::number(3.0));
  // A3:A1 describes the same rectangle as A1:A3.
  auto forward = ExpandRange(sheet, MakeLocalRef(0, 0), MakeLocalRef(2, 0));
  auto reversed = ExpandRange(sheet, MakeLocalRef(2, 0), MakeLocalRef(0, 0));
  ASSERT_TRUE(forward);
  ASSERT_TRUE(reversed);
  ASSERT_EQ(reversed.value().size(), forward.value().size());
  for (std::size_t i = 0; i < forward.value().size(); ++i) {
    EXPECT_EQ(reversed.value()[i].as_number(), forward.value()[i].as_number()) << "i=" << i;
  }
}

TEST(EvalContextExpandRange, BlankCellsBecomeBlank) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(1.0));
  // A2 is intentionally untouched.
  sheet.set_cell_value(2, 0, Value::number(3.0));
  auto result = ExpandRange(sheet, MakeLocalRef(0, 0), MakeLocalRef(2, 0));
  ASSERT_TRUE(result);
  ASSERT_EQ(result.value().size(), 3u);
  EXPECT_TRUE(result.value()[0].is_number());
  EXPECT_TRUE(result.value()[1].is_blank());
  EXPECT_TRUE(result.value()[2].is_number());
}

TEST(EvalContextExpandRange, CrossSheetLhs_ReturnsRef) {
  Sheet sheet("Sheet1");
  parser::Reference lhs = MakeLocalRef(0, 0);
  lhs.sheet = std::string_view("Other");
  auto result = ExpandRange(sheet, lhs, MakeLocalRef(2, 0));
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error(), ErrorCode::Ref);
}

TEST(EvalContextExpandRange, CrossSheetRhs_ReturnsRef) {
  Sheet sheet("Sheet1");
  parser::Reference rhs = MakeLocalRef(2, 0);
  rhs.sheet = std::string_view("Other");
  auto result = ExpandRange(sheet, MakeLocalRef(0, 0), rhs);
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error(), ErrorCode::Ref);
}

TEST(EvalContextExpandRange, FullColumn_ReturnsValue) {
  Sheet sheet("Sheet1");
  parser::Reference lhs = MakeLocalRef(0, 0);
  lhs.is_full_col = true;
  parser::Reference rhs = MakeLocalRef(0, 0);
  rhs.is_full_col = true;
  auto result = ExpandRange(sheet, lhs, rhs);
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error(), ErrorCode::Value);
}

TEST(EvalContextExpandRange, FullRow_ReturnsValue) {
  Sheet sheet("Sheet1");
  parser::Reference lhs = MakeLocalRef(0, 0);
  lhs.is_full_row = true;
  parser::Reference rhs = MakeLocalRef(0, 0);
  rhs.is_full_row = true;
  auto result = ExpandRange(sheet, lhs, rhs);
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error(), ErrorCode::Value);
}

TEST(EvalContextExpandRange, OutOfBoundsLhs_ReturnsRef) {
  Sheet sheet("Sheet1");
  auto result = ExpandRange(sheet, MakeLocalRef(Sheet::kMaxRows, 0), MakeLocalRef(0, 0));
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error(), ErrorCode::Ref);
}

TEST(EvalContextExpandRange, OutOfBoundsRhs_ReturnsRef) {
  Sheet sheet("Sheet1");
  auto result = ExpandRange(sheet, MakeLocalRef(0, 0), MakeLocalRef(0, Sheet::kMaxCols));
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error(), ErrorCode::Ref);
}

TEST(EvalContextExpandRange, LargeButValidRectangle_Reserves) {
  Sheet sheet("Sheet1");
  // Populate a 10x10 numeric grid so the probe cells are not blank. Values
  // encode (row, col) as `row * 10 + col` for easy ordering verification.
  for (std::uint32_t r = 0; r < 10; ++r) {
    for (std::uint32_t c = 0; c < 10; ++c) {
      sheet.set_cell_value(r, c, Value::number(static_cast<double>(r * 10 + c)));
    }
  }
  auto result = ExpandRange(sheet, MakeLocalRef(0, 0), MakeLocalRef(9, 9));
  ASSERT_TRUE(result);
  ASSERT_EQ(result.value().size(), 100u);
  // Row-major ordering probes: index = row * width + col, width = 10.
  EXPECT_EQ(result.value()[0].as_number(), 0.0);    // (0,0)
  EXPECT_EQ(result.value()[9].as_number(), 9.0);    // (0,9)
  EXPECT_EQ(result.value()[10].as_number(), 10.0);  // (1,0)
  EXPECT_EQ(result.value()[55].as_number(), 55.0);  // (5,5)
  EXPECT_EQ(result.value()[99].as_number(), 99.0);  // (9,9)
}

TEST(EvalContextExpandRange, FormulaCellInRange_IsEvaluated) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(10.0));
  sheet.set_cell_formula(1, 0, "=A1+1");
  sheet.set_cell_value(2, 0, Value::number(20.0));
  EvalState state;
  static thread_local Arena arena;
  arena.reset();
  const eval::EvalContext ctx(sheet, state);
  auto result = ctx.expand_range(MakeLocalRef(0, 0), MakeLocalRef(2, 0), arena,
                                 eval::default_registry());
  ASSERT_TRUE(result);
  ASSERT_EQ(result.value().size(), 3u);
  EXPECT_EQ(result.value()[0].as_number(), 10.0);
  EXPECT_EQ(result.value()[1].as_number(), 11.0);
  EXPECT_EQ(result.value()[2].as_number(), 20.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
