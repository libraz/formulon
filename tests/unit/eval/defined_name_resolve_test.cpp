//
// Tests for workbook / sheet-scoped defined-name resolution in the
// tree-walk evaluator. Each case builds a workbook with `set_defined_names`,
// parses a formula, and evaluates it against a sheet-bound `EvalContext`.
// The resolution rules mirror `dep_extractor.cpp` (sheet scope wins over
// workbook scope; sheet-scoped names are invisible from other sheets).

#include "eval/defined_name_resolve.h"

#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it against `ctx`. Aborts the test on parse
// failure so callers can assert on the result shape directly.
Value EvalOrDie(std::string_view src, Arena& arena, const EvalContext& ctx) {
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, arena, default_registry(), ctx);
}

// (a) Workbook-scoped constant name: =A1*Rate with Rate=0.1.
TEST(DefinedNameResolve, WorkbookScopeConstant) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(200.0));  // A1
  wb.set_defined_names({io::DefinedName{"Rate", "0.1", -1, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=A1*Rate", a, ctx);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

// (b) Sheet scope wins over workbook scope for the same name.
TEST(DefinedNameResolve, SheetScopeOverridesWorkbookScope) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");  // sheet index 0
  wb.set_defined_names({
      io::DefinedName{"Rate", "0.1", -1, false, ""},  // workbook scope
      io::DefinedName{"Rate", "0.2", 0, false, ""},   // Sheet1 scope
  });
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=Rate", a, ctx);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 0.2);
}

// (c) A sheet-scoped name is invisible from a different sheet (matching the
// dep_extractor rule); resolution falls through to #NAME?.
TEST(DefinedNameResolve, SheetScopedNameInvisibleFromOtherSheet) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");                                                // index 0
  Sheet& s2 = wb.add_sheet("Sheet2");                                    // index 1
  wb.set_defined_names({io::DefinedName{"Local", "42", 0, false, ""}});  // Sheet1 scope
  EvalState state;
  EvalContext ctx(wb, s2, state);
  Arena a;
  const Value v = EvalOrDie("=Local", a, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// (c') The same sheet-scoped name resolves on its owning sheet.
TEST(DefinedNameResolve, SheetScopedNameVisibleOnOwningSheet) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");  // index 0
  wb.add_sheet("Sheet2");
  // Fetch the sheet reference only after every add_sheet: add_sheet may
  // reallocate the internal sheets vector and invalidate earlier references.
  const Sheet& s1 = wb.sheet(0);
  wb.set_defined_names({io::DefinedName{"Local", "42", 0, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s1, state);
  Arena a;
  const Value v = EvalOrDie("=Local", a, ctx);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

// (d) Reference-type definition (=Sheet1!$A$1) evaluates through the cell and
// follows the cell's value on re-evaluation.
TEST(DefinedNameResolve, ReferenceDefinitionFollowsCellValue) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(5.0));  // A1
  wb.set_defined_names({io::DefinedName{"Ref", "Sheet1!$A$1", -1, false, ""}});
  {
    EvalState state;
    EvalContext ctx(wb, s, state);
    Arena a;
    const Value v = EvalOrDie("=Ref*2", a, ctx);
    ASSERT_TRUE(v.is_number()) << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
  }
  // Change the underlying cell and re-evaluate: the name tracks it.
  s.set_cell_value(0U, 0U, Value::number(7.0));
  {
    EvalState state;
    EvalContext ctx(wb, s, state);
    Arena a;
    const Value v = EvalOrDie("=Ref*2", a, ctx);
    ASSERT_TRUE(v.is_number()) << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), 14.0);
  }
}

// (e) Formula-type definition (=A1*2) evaluates against the current sheet.
TEST(DefinedNameResolve, FormulaDefinitionEvaluates) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(6.0));  // A1
  wb.set_defined_names({io::DefinedName{"Doubled", "A1*2", -1, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=Doubled+1", a, ctx);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 13.0);
}

// (f) An undefined name is #NAME?.
TEST(DefinedNameResolve, UndefinedNameIsNameError) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=Nonexistent", a, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// (g) A directly circular definition surfaces #REF! without hanging.
TEST(DefinedNameResolve, CircularDefinitionDoesNotHang) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  wb.set_defined_names({io::DefinedName{"Loop", "Loop+1", -1, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=Loop", a, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// (g') A mutually recursive pair (A=B, B=A) also terminates with #REF!.
TEST(DefinedNameResolve, MutualCycleDoesNotHang) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  wb.set_defined_names({
      io::DefinedName{"AName", "BName", -1, false, ""},
      io::DefinedName{"BName", "AName", -1, false, ""},
  });
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=AName", a, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// (h) A LET binding shadows a defined name of the same identifier.
TEST(DefinedNameResolve, LetBindingShadowsDefinedName) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  wb.set_defined_names({io::DefinedName{"Rate", "0.1", -1, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=LET(Rate, 2, Rate)", a, ctx);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

// A nested defined name (name references another name) resolves transitively.
TEST(DefinedNameResolve, NestedNameResolvesTransitively) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  wb.set_defined_names({
      io::DefinedName{"Base", "10", -1, false, ""},
      io::DefinedName{"Derived", "Base*3", -1, false, ""},
  });
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=Derived", a, ctx);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

// A workbook-scoped range name expands to its full shape inside a range-aware
// aggregator instead of collapsing to its top-left cell.
TEST(DefinedNameResolve, RangeNameExpandsInAggregators) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(10.0));  // A1
  s.set_cell_value(1U, 0U, Value::number(20.0));  // A2
  s.set_cell_value(2U, 0U, Value::number(30.0));  // A3
  s.set_cell_value(3U, 0U, Value::number(40.0));  // A4
  s.set_cell_value(4U, 0U, Value::number(50.0));  // A5
  wb.set_defined_names({io::DefinedName{"MyRange", "Sheet1!$A$1:$A$5", -1, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value sum = EvalOrDie("=SUM(MyRange)", a, ctx);
  ASSERT_TRUE(sum.is_number()) << sum.debug_to_string();
  EXPECT_DOUBLE_EQ(sum.as_number(), 150.0);
  const Value count = EvalOrDie("=COUNT(MyRange)", a, ctx);
  ASSERT_TRUE(count.is_number()) << count.debug_to_string();
  EXPECT_DOUBLE_EQ(count.as_number(), 5.0);
  const Value avg = EvalOrDie("=AVERAGE(MyRange)", a, ctx);
  ASSERT_TRUE(avg.is_number()) << avg.debug_to_string();
  EXPECT_DOUBLE_EQ(avg.as_number(), 30.0);
}

// A sheet-scoped range name expands the same way on its owning sheet.
TEST(DefinedNameResolve, SheetScopedRangeNameExpands) {
  Workbook wb = Workbook::create_empty();
  Sheet& s1 = wb.add_sheet("Sheet1");             // index 0
  s1.set_cell_value(0U, 0U, Value::number(1.0));  // A1
  s1.set_cell_value(1U, 0U, Value::number(2.0));  // A2
  s1.set_cell_value(2U, 0U, Value::number(3.0));  // A3
  wb.set_defined_names({io::DefinedName{"LocalRange", "Sheet1!$A$1:$A$3", 0, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s1, state);
  Arena a;
  const Value sum = EvalOrDie("=SUM(LocalRange)", a, ctx);
  ASSERT_TRUE(sum.is_number()) << sum.debug_to_string();
  EXPECT_DOUBLE_EQ(sum.as_number(), 6.0);
}

// Regression guard: a single-cell reference name stays scalar (does not become
// a 1x1 array via the range-shaped path).
TEST(DefinedNameResolve, SingleCellNameStaysScalar) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(7.0));  // A1
  wb.set_defined_names({io::DefinedName{"Cell", "Sheet1!$A$1", -1, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=Cell", a, ctx);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

// A workbook-defined LAMBDA is callable through the ordinary `Name(args)`
// syntax. The authored name and call lexeme are matched case-insensitively,
// and the body executes in the caller's workbook context.
TEST(DefinedNameResolve, WorkbookNamedLambdaCallIsCaseInsensitive) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(10.0));  // A1
  wb.set_defined_names({io::DefinedName{"AddA1", "LAMBDA(x,x+A1)", -1, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=adda1(5)", a, ctx);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(DefinedNameResolve, SheetNamedLambdaShadowsWorkbookLambda) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  const Sheet& s1 = wb.sheet(0);
  const Sheet& s2 = wb.sheet(1);
  wb.set_defined_names({
      io::DefinedName{"Adder", "LAMBDA(x,x+1)", -1, false, ""},
      io::DefinedName{"ADDER", "LAMBDA(x,x+10)", 0, false, ""},
  });

  EvalState state1;
  EvalContext ctx1(wb, s1, state1);
  Arena a1;
  const Value local = EvalOrDie("=Adder(2)", a1, ctx1);
  ASSERT_TRUE(local.is_number()) << local.debug_to_string();
  EXPECT_DOUBLE_EQ(local.as_number(), 12.0);

  EvalState state2;
  EvalContext ctx2(wb, s2, state2);
  Arena a2;
  const Value global = EvalOrDie("=aDdEr(2)", a2, ctx2);
  ASSERT_TRUE(global.is_number()) << global.debug_to_string();
  EXPECT_DOUBLE_EQ(global.as_number(), 3.0);
}

TEST(DefinedNameResolve, LexicalLambdaShadowsDefinedLambda) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  wb.set_defined_names({io::DefinedName{"f", "LAMBDA(x,x+100)", -1, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=LET(F,LAMBDA(x,x+1),f(2))", a, ctx);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(DefinedNameResolve, NamedLambdaNonLambdaAndArityErrors) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  wb.set_defined_names({
      io::DefinedName{"Constant", "7", -1, false, ""},
      io::DefinedName{"Pair", "LAMBDA(x,y,x+y)", -1, false, ""},
  });
  EvalState state;
  EvalContext ctx(wb, s, state);

  Arena non_lambda_arena;
  const Value non_lambda = EvalOrDie("=Constant(2)", non_lambda_arena, ctx);
  ASSERT_TRUE(non_lambda.is_error());
  EXPECT_EQ(non_lambda.as_error(), ErrorCode::Value);

  Arena arity_arena;
  const Value wrong_arity = EvalOrDie("=Pair(1)", arity_arena, ctx);
  ASSERT_TRUE(wrong_arity.is_error());
  EXPECT_EQ(wrong_arity.as_error(), ErrorCode::Value);
}

TEST(DefinedNameResolve, NamedLambdaParseFailureIsNameError) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  wb.set_defined_names({io::DefinedName{"Broken", "LAMBDA(x,x+)", -1, false, ""}});
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=Broken(1)", a, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(DefinedNameResolve, NamedLambdaSupportsNamedRecursion) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  wb.set_defined_names({
      io::DefinedName{"Fact", "LAMBDA(n,IF(n<=1,1,n*Fact(n-1)))", -1, false, ""},
  });
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=fact(5)", a, ctx);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 120.0);
}

TEST(DefinedNameResolve, NamedLambdaRunawayRecursionHitsCalcCap) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  wb.set_defined_names({
      io::DefinedName{"LoopFn", "LAMBDA(n,LoopFn(n+1))", -1, false, ""},
  });
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = EvalOrDie("=loopfn(0)", a, ctx);
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
