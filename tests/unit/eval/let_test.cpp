//
// Evaluator tests for the `LET` special form and bare `NameRef` lookup.
// Covers sequential binding, shadowing, error flow, nested LETs, case
// insensitivity, and the base-case `#NAME?` when a name is unbound.

#include <string_view>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "test_eval_helpers.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

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

Value EvalSourceWithHost(std::string_view src, ExcelHost host) {
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
  const EvalContext ctx = test::host_context(host);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// Scalar binding shapes
// ---------------------------------------------------------------------------

TEST(EvalLet, SingleBindingArithmetic) {
  const Value v = EvalSource("=LET(x, 10, x+5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 15.0);
}

TEST(EvalLet, SequentialBindingsUsePriorScope) {
  // y is defined in terms of x, which is already in scope.
  const Value v = EvalSource("=LET(x, 10, y, x*2, y+1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 21.0);
}

TEST(EvalLet, BodyReferencesFirstBinding) {
  const Value v = EvalSource("=LET(x, 7, x)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(EvalLet, NumericConstantBinding) {
  const Value v = EvalSource("=LET(x, 0, x)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(EvalLet, TextValueBinding) {
  const Value v = EvalSource("=LET(g, \"hello\", UPPER(g))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "HELLO");
}

TEST(EvalLet, BooleanValueBinding) {
  const Value v = EvalSource("=LET(b, TRUE, IF(b, 1, 0))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(EvalLet, FunctionCallOnBoundValue) {
  const Value v = EvalSource("=LET(x, 3, MAX(x, 5))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(EvalHostProfile, WinHostKeepsModernExcel365FunctionsAvailable) {
  Value v = EvalSourceWithHost("=LET(x, 1, x)", ExcelHost::kWin365);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);

  v = EvalSourceWithHost("=LAMBDA(x, x+1)(5)", ExcelHost::kWin365);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);

  v = EvalSourceWithHost("=SEQUENCE(3)", ExcelHost::kWin365);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);

  v = EvalSourceWithHost("=VALUETOTEXT(1)", ExcelHost::kWin365);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1");

  v = EvalSourceWithHost("=UNIQUE({1;1;2})", ExcelHost::kWin365);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);

  v = EvalSourceWithHost("=TEXTSPLIT(\"a,b\", \",\")", ExcelHost::kWin365);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_cols(), 2U);
}

// ---------------------------------------------------------------------------
// Nested LETs
// ---------------------------------------------------------------------------

TEST(EvalLet, NestedLetInBindingExpression) {
  // Inner LET produces 25; outer body adds 1.
  const Value v = EvalSource("=LET(x, LET(y, 5, y*y), x+1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 26.0);
}

TEST(EvalLet, NestedLetInBody) {
  const Value v = EvalSource("=LET(x, 2, LET(y, x+1, y*y))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 9.0);
}

// ---------------------------------------------------------------------------
// Shadowing
// ---------------------------------------------------------------------------

TEST(EvalLet, LaterBindingShadowsEarlier) {
  // Second `x` binds (outer x) + 10 = 11. Body reads the shadowed value.
  const Value v = EvalSource("=LET(x, 1, x, x+10, x)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 11.0);
}

TEST(EvalLet, InnerLetDoesNotLeakOuterScope) {
  // Outer x is shadowed inside the inner LET's body but recovers after.
  const Value v = EvalSource("=LET(x, 1, LET(x, 99, x) + x)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 100.0);
}

// ---------------------------------------------------------------------------
// Case sensitivity
// ---------------------------------------------------------------------------

TEST(EvalLet, NameLookupIsCaseInsensitive) {
  const Value v = EvalSource("=LET(Foo, 3, FOO+1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(EvalLet, ShadowingIsCaseInsensitive) {
  // FOO and foo are the same identifier; the second binding shadows.
  const Value v = EvalSource("=LET(FOO, 1, foo, 9, Foo)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 9.0);
}

// ---------------------------------------------------------------------------
// Error flow
// ---------------------------------------------------------------------------

TEST(EvalLet, BindingErrorIsCatchableInBody) {
  // LET does not short-circuit on a binding initialiser that evaluates to
  // an error -- the error becomes the binding's value and IFERROR in the
  // body can catch it.
  const Value v = EvalSource("=LET(x, 1/0, IFERROR(x, 99))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 99.0);
}

TEST(EvalLet, BindingErrorPropagatesWhenUncaught) {
  const Value v = EvalSource("=LET(x, 1/0, x+1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(EvalLet, ErrorInBodyPropagates) {
  const Value v = EvalSource("=LET(x, 5, x/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Name-resolution boundary
// ---------------------------------------------------------------------------

TEST(EvalLet, UnboundNameRefIsNameError) {
  // Bare identifier at the top level (no LET in scope, no workbook-level
  // name resolution wired) resolves to #NAME?.
  const Value v = EvalSource("=foo");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(EvalLet, NameOutOfLetScopeIsNameError) {
  // `y` is never bound; the inner reference fails.
  const Value v = EvalSource("=LET(x, 1, y)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(EvalLet, ForwardReferenceBeforeBindingIsNameError) {
  // The body of `x`'s initialiser runs BEFORE y is bound, so `y` is
  // unresolved and the binding value becomes #NAME?. Body then uses x
  // which carries that error forward.
  const Value v = EvalSource("=LET(x, y, y, 2, x)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// Reference-returning initialisers
// ---------------------------------------------------------------------------

// A1:A10 = 1..10, B1:B10 = 10..100, C1:C10 = 100..1000, D1 = "abc".
Workbook ReferenceFixture() {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 10; ++r) {
    const double n = static_cast<double>(r + 1U);
    wb.sheet(0).set_cell_value(r, 0, Value::number(n));
    wb.sheet(0).set_cell_value(r, 1, Value::number(n * 10.0));
    wb.sheet(0).set_cell_value(r, 2, Value::number(n * 100.0));
  }
  wb.sheet(0).set_cell_value(0, 3, Value::text("abc"));
  return wb;
}

Value EvalRef(std::string_view src) {
  const Workbook wb = ReferenceFixture();
  return formulon::test::EvalSourceAt(src, wb, wb.sheet(0), 19, 7);
}

void ExpectNumber(std::string_view src, double expected) {
  const Value v = EvalRef(src);
  ASSERT_TRUE(v.is_number()) << src;
  EXPECT_EQ(v.as_number(), expected) << src;
}

TEST(EvalLetReference, RowOfIndexBinding) {
  ExpectNumber("=LET(r,INDEX(A1:A10,3),ROW(r))", 3.0);
}

TEST(EvalLetReference, ColumnOfIndexColumnBinding) {
  ExpectNumber("=LET(r,INDEX(A1:B10,0,2),COLUMN(r))", 2.0);
}

TEST(EvalLetReference, RowsAndColumnsOfIndexColumnBinding) {
  ExpectNumber("=LET(r,INDEX(A1:B10,0,2),ROWS(r))", 10.0);
  ExpectNumber("=LET(r,INDEX(A1:B10,0,2),COLUMNS(r))", 1.0);
}

TEST(EvalLetReference, IsrefAndAreasOfIndexBinding) {
  const Value isref = EvalRef("=LET(r,INDEX(A1:A10,3),ISREF(r))");
  ASSERT_TRUE(isref.is_boolean());
  EXPECT_TRUE(isref.as_boolean());
  ExpectNumber("=LET(r,INDEX(A1:A10,3),AREAS(r))", 1.0);
}

TEST(EvalLetReference, CellAddressOfXlookupBinding) {
  const Value v = EvalRef("=LET(r,XLOOKUP(3,A1:A10,B1:B10),CELL(\"address\",r))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "$B$3");
}

TEST(EvalLetReference, RangeEndpointFromIndexBinding) {
  ExpectNumber("=LET(r,INDEX(A1:A10,3),SUM(A1:r))", 6.0);
}

TEST(EvalLetReference, IfsAndSwitchBindings) {
  ExpectNumber("=LET(r,IFS(TRUE,B5),ROW(r))", 5.0);
  ExpectNumber("=LET(r,SWITCH(1,1,C3),COLUMN(r))", 3.0);
}

TEST(EvalLetReference, ShadowedBindingResolvesAgainstOuterScope) {
  ExpectNumber("=LET(x,A1:A10,LET(x,INDEX(x,3),ROW(x)))", 3.0);
}

TEST(EvalLetReference, IndexOverArrayLiteralIsNotReference) {
  const Value v = EvalRef("=LET(r,INDEX({1,2,3},2),ISREF(r))");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(EvalLetReference, CellRowOfPlainRefBinding) {
  ExpectNumber("=LET(r,A5,CELL(\"row\",r))", 5.0);
}

TEST(EvalLetReference, MultiCellReferenceCallBindingsAggregate) {
  ExpectNumber("=LET(r,OFFSET(A1,0,0,3,1),SUM(r))", 6.0);
  ExpectNumber("=LET(r,CHOOSE(2,A1:A2,A1:A3),SUM(r))", 6.0);
  ExpectNumber("=LET(r,IF(TRUE,A1:A4,B1:B4),SUM(r))", 10.0);
  ExpectNumber("=LET(r,INDIRECT(\"A1:A5\"),SUM(r))", 15.0);
}

TEST(EvalLetReference, SingleCellTextBindingAggregatesAsReference) {
  // A one-cell reference returned by a call stays a reference, so SUM skips
  // the text in it exactly as `SUM(OFFSET(D1,0,0))` does.
  ExpectNumber("=SUM(OFFSET(D1,0,0))", 0.0);
  ExpectNumber("=LET(r,OFFSET(D1,0,0),SUM(r))", 0.0);
  ExpectNumber("=LET(r,CHOOSE(1,D1),SUM(r))", 0.0);
  ExpectNumber("=LET(r,IF(TRUE,D1),SUM(r))", 0.0);
  ExpectNumber("=LET(r,INDIRECT(\"D1\"),SUM(r))", 0.0);
  ExpectNumber("=LET(r,INDEX(D1:D2,1),SUM(r))", 0.0);
}

TEST(EvalLetReference, SingleCellBindingValueIsScalar) {
  ExpectNumber("=LET(r,INDEX(A1:A10,3),r+1)", 4.0);
  ExpectNumber("=LET(r,OFFSET(A1,4,0),r*2)", 10.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
