//
// Direct VM tests. Each case parses a small formula, runs both the
// tree-walker and the bytecode VM, and asserts the result matches. The
// tree-walker remains the source of truth: a divergence in any test below
// is a VM bug.
//
// Tests that fabricate malformed bytecode by hand exercise the VM-fault
// surface (`kVmEmptyBytecode` / `kVmStackUnderflow`).

#include "eval/vm.h"

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/bytecode.h"
#include "eval/compiler.h"
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
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// Parse + compile + execute. Returns the VM's result as a Value. Aborts the
// test on parse / compile failure so callers can ASSERT_TRUE on the result
// shape directly.
Value RunVmOrDie(std::string_view src, Arena& arena) {
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  auto bc = compile(*root, arena);
  EXPECT_TRUE(bc.has_value()) << "compile failed for: " << src;
  if (!bc.has_value()) {
    return Value::error(ErrorCode::Name);
  }
  auto out = execute(bc.value(), arena, default_registry(), test::mac_context());
  EXPECT_TRUE(out.has_value()) << "VM error for: " << src;
  if (!out.has_value()) {
    return Value::error(ErrorCode::Value);
  }
  return out.value();
}

Value RunVmWithCtx(std::string_view src, Arena& arena, const EvalContext& ctx) {
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  auto bc = compile(*root, arena);
  EXPECT_TRUE(bc.has_value()) << "compile failed for: " << src;
  if (!bc.has_value()) {
    return Value::error(ErrorCode::Name);
  }
  auto out = execute(bc.value(), arena, default_registry(), ctx);
  EXPECT_TRUE(out.has_value()) << "VM error for: " << src;
  if (!out.has_value()) {
    return Value::error(ErrorCode::Value);
  }
  return out.value();
}

Value RunTreeOrDie(std::string_view src, Arena& arena) {
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, arena, default_registry(), test::mac_context());
}

Value RunTreeWithCtx(std::string_view src, Arena& arena, const EvalContext& ctx) {
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, arena, default_registry(), ctx);
}

bool ValuesAgreeBitExact(const Value& a, const Value& b) {
  if (a.kind() != b.kind()) {
    return false;
  }
  switch (a.kind()) {
    case ValueKind::Blank:
      return true;
    case ValueKind::Number:
      // Bit-exact equality (NaN / Inf payloads matter for parity).
      return std::memcmp(&a, &b, sizeof(Value)) == 0 ? true : (a.as_number() == b.as_number());
    case ValueKind::Bool:
      return a.as_boolean() == b.as_boolean();
    case ValueKind::Error:
      return a.as_error() == b.as_error();
    case ValueKind::Text:
      return a.as_text() == b.as_text();
    case ValueKind::Array: {
      const ArrayValue* la = a.as_array();
      const ArrayValue* ra = b.as_array();
      return la->rows == ra->rows && la->cols == ra->cols;
    }
    default:
      return false;
  }
}

// ---------------------------------------------------------------------------
// Scalar literals
// ---------------------------------------------------------------------------

TEST(Vm, NumberLiteral) {
  Arena a;
  const Value v = RunVmOrDie("=42", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

TEST(Vm, BoolLiteralTrue) {
  Arena a;
  const Value v = RunVmOrDie("=TRUE", a);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(Vm, BoolLiteralFalse) {
  Arena a;
  const Value v = RunVmOrDie("=FALSE", a);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(Vm, TextLiteral) {
  Arena a;
  const Value v = RunVmOrDie("=\"hello\"", a);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello");
}

TEST(Vm, ErrorLiteral) {
  Arena a;
  const Value v = RunVmOrDie("=#DIV/0!", a);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Arithmetic and comparison
// ---------------------------------------------------------------------------

TEST(Vm, ArithmeticAdd) {
  Arena a;
  const Value v = RunVmOrDie("=1+2", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(Vm, ArithmeticChain) {
  Arena a;
  const Value v = RunVmOrDie("=1+2*3-4/2", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(Vm, ArithmeticPow) {
  Arena a;
  const Value v = RunVmOrDie("=2^10", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1024.0);
}

TEST(Vm, ArithmeticDivByZero) {
  Arena a;
  const Value v = RunVmOrDie("=1/0", a);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(Vm, UnaryMinus) {
  Arena a;
  const Value v = RunVmOrDie("=-5", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -5.0);
}

TEST(Vm, UnaryPercent) {
  Arena a;
  const Value v = RunVmOrDie("=50%", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(Vm, ComparisonEq) {
  Arena a;
  const Value v = RunVmOrDie("=3=3", a);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(Vm, ComparisonLt) {
  Arena a;
  const Value v = RunVmOrDie("=2<3", a);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(Vm, ComparisonNotEq) {
  Arena a;
  const Value v = RunVmOrDie("=1<>2", a);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(Vm, ConcatBinop) {
  Arena a;
  const Value v = RunVmOrDie("=\"hello \" & \"world\"", a);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello world");
}

TEST(Vm, ConcatNumberAndText) {
  Arena a;
  const Value v = RunVmOrDie("=\"=\" & 42", a);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "=42");
}

// ---------------------------------------------------------------------------
// Function calls (eager dispatch)
// ---------------------------------------------------------------------------

TEST(Vm, CallSumScalars) {
  Arena a;
  const Value v = RunVmOrDie("=SUM(1,2,3)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(Vm, CallAbs) {
  Arena a;
  const Value v = RunVmOrDie("=ABS(-7)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(Vm, CallPower) {
  Arena a;
  const Value v = RunVmOrDie("=POWER(2,8)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 256.0);
}

TEST(Vm, CallUnknownNameError) {
  Arena a;
  const Value v = RunVmOrDie("=NOSUCHFN(1,2)", a);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(Vm, CallErrorPropagates) {
  Arena a;
  const Value v = RunVmOrDie("=ABS(1/0)", a);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Lazy IF / IFERROR / IFNA
// ---------------------------------------------------------------------------

TEST(Vm, IfTrueBranch) {
  Arena a;
  const Value v = RunVmOrDie("=IF(TRUE, 10, 20)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(Vm, IfFalseBranch) {
  Arena a;
  const Value v = RunVmOrDie("=IF(FALSE, 10, 20)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

TEST(Vm, IfErrorPrimaryWins) {
  Arena a;
  const Value v = RunVmOrDie("=IFERROR(42, 99)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

TEST(Vm, IfErrorTriggersFallback) {
  Arena a;
  const Value v = RunVmOrDie("=IFERROR(1/0, 99)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 99.0);
}

TEST(Vm, IfNaPropagatesNonNa) {
  Arena a;
  const Value v = RunVmOrDie("=IFNA(1/0, 99)", a);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(Vm, IfNaCatchesNa) {
  Arena a;
  const Value v = RunVmOrDie("=IFNA(NA(), 99)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 99.0);
}

TEST(Vm, IfNestedShortCircuit) {
  Arena a;
  // Tree-walker comparison ensures nested IF + arithmetic agree.
  const Value tree = RunTreeOrDie("=IF(1<2, IF(3<4, 10, 20), 30)", a);
  Arena a2;
  const Value vm = RunVmOrDie("=IF(1<2, IF(3<4, 10, 20), 30)", a2);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm));
}

TEST(Vm, IfTextConditionIsValueError) {
  // A non-numeric text IF condition is not boolean-coercible; Excel (and the
  // tree-walker) surface #VALUE! instead of taking the THEN branch. The VM
  // routes the condition through the same `coerce_to_bool` helper, so the two
  // paths must agree.
  Arena a;
  const Value vm = RunVmOrDie("=IF(\"hello\", 10, 20)", a);
  ASSERT_TRUE(vm.is_error());
  EXPECT_EQ(vm.as_error(), ErrorCode::Value);
  Arena a2;
  const Value tree = RunTreeOrDie("=IF(\"hello\", 10, 20)", a2);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm));
}

TEST(Vm, IfTextTrueConditionTakesThenBranch) {
  // The literal "TRUE" is the one text form `coerce_to_bool` accepts; both
  // evaluators take the THEN branch.
  Arena a;
  const Value vm = RunVmOrDie("=IF(\"TRUE\", 10, 20)", a);
  ASSERT_TRUE(vm.is_number());
  EXPECT_DOUBLE_EQ(vm.as_number(), 10.0);
  Arena a2;
  const Value tree = RunTreeOrDie("=IF(\"TRUE\", 10, 20)", a2);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm));
}

TEST(Vm, RunawayRecursiveLambdaReturnsCalcNotCrash) {
  // A self-recursive LAMBDA (Y-combinator style: the lambda takes itself as a
  // parameter and applies itself unconditionally) would overflow the native
  // stack without a depth guard. The VM caps lambda-call recursion at the
  // same depth as the tree-walker and surfaces #CALC! rather than crashing.
  constexpr const char* kSrc = "=LET(g, LAMBDA(self, n, self(self, n+1)), g(g, 0))";
  Arena a;
  const Value vm = RunVmOrDie(kSrc, a);
  ASSERT_TRUE(vm.is_error());
  EXPECT_EQ(vm.as_error(), ErrorCode::Calc);
  Arena a2;
  const Value tree = RunTreeOrDie(kSrc, a2);
  ASSERT_TRUE(tree.is_error());
  EXPECT_EQ(tree.as_error(), ErrorCode::Calc);
}

// ---------------------------------------------------------------------------
// Array literals + LET + LAMBDA
// ---------------------------------------------------------------------------

TEST(Vm, ArrayLiteralSum) {
  Arena a;
  const Value v = RunVmOrDie("=SUM({1,2,3})", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(Vm, LetSimple) {
  Arena a;
  const Value v = RunVmOrDie("=LET(x, 10, x+5)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(Vm, LetMultipleBindings) {
  Arena a;
  const Value v = RunVmOrDie("=LET(x, 10, y, 20, x+y)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(Vm, LetShadowing) {
  Arena a;
  const Value v = RunVmOrDie("=LET(x, 1, x, x+10, x)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 11.0);
}

TEST(Vm, LetLookupIsCaseInsensitive) {
  // The body references `X` (uppercase) but binds `x`; Excel folds ASCII
  // case on name resolution, so this must resolve rather than surface
  // `#NAME?`. Mirrors the tree-walker's case-insensitive NameEnv.
  Arena a;
  const Value v = RunVmOrDie("=LET(x, 10, X+5)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(Vm, LambdaParamShadowsOuterLetByNesting) {
  // The inner lambda parameter `x` shadows the outer LET `x` inside the
  // lambda body, so calling with 5 yields 5+100=105. A kind-split lookup
  // that always resolves LET before lambda params would wrongly read the
  // outer LET value (1) and return 101.
  Arena a;
  const Value v = RunVmOrDie("=LET(x, 1, LAMBDA(x, x+100)(5))", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 105.0);
}

TEST(Vm, LambdaParamLookupIsCaseInsensitive) {
  Arena a;
  const Value v = RunVmOrDie("=LAMBDA(x, X+1)(5)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(Vm, LambdaIife) {
  Arena a;
  const Value v = RunVmOrDie("=LAMBDA(x, x+1)(5)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(Vm, LambdaTwoArgs) {
  Arena a;
  const Value v = RunVmOrDie("=LAMBDA(x, y, x*y+1)(3, 4)", a);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 13.0);
}

// ---------------------------------------------------------------------------
// Refs (single-cell) via in-memory Sheet
// ---------------------------------------------------------------------------

TEST(Vm, RefResolvesScalarCell) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(7.0));  // A1 = 7
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = RunVmWithCtx("=A1", a, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(Vm, RefArithmetic) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(2.0));  // A1 = 2
  s.set_cell_value(1U, 0U, Value::number(5.0));  // A2 = 5
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value v = RunVmWithCtx("=A1*A2+1", a, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 11.0);
}

// ---------------------------------------------------------------------------
// Many string constants: stable `string_storage` borrow contract
// ---------------------------------------------------------------------------

TEST(Vm, ManyStringConstantsDoNotDangle) {
  // `ByteCode::string_storage` backs the `string_view`s that Text constants
  // borrow. It must keep every element address stable no matter how many
  // literals are interned; concatenating far more than any small-reserve
  // heuristic would forces the pool to grow well past its initial size. A
  // dangling borrow here would surface as a corrupted / crashing result.
  const std::string src =
      "=CONCAT(\"s0\",\"s1\",\"s2\",\"s3\",\"s4\",\"s5\",\"s6\",\"s7\",\"s8\",\"s9\","
      "\"s10\",\"s11\",\"s12\",\"s13\",\"s14\",\"s15\",\"s16\",\"s17\",\"s18\",\"s19\")";
  Arena a;
  const Value v = RunVmOrDie(src, a);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "s0s1s2s3s4s5s6s7s8s9s10s11s12s13s14s15s16s17s18s19");
}

TEST(Vm, ManyStringConstantsConcatAgreesWithTreeWalker) {
  // A second high-cardinality constant-pool path cross-checked against the
  // tree-walker: the two evaluators must agree on the concatenation.
  const std::string src =
      "=CONCAT(\"s0\",\"s1\",\"s2\",\"s3\",\"s4\",\"s5\",\"s6\",\"s7\",\"s8\",\"s9\","
      "\"s10\",\"s11\",\"s12\",\"s13\",\"s14\",\"s15\",\"s16\",\"s17\",\"s18\",\"s19\")";
  Arena a;
  const Value vm = RunVmOrDie(src, a);
  Arena a2;
  const Value tree = RunTreeOrDie(src, a2);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm)) << "vm=" << vm.debug_to_string();
}

// ---------------------------------------------------------------------------
// Range arguments: SUM(A1:B2) must aggregate every cell, not collapse
// ---------------------------------------------------------------------------

TEST(Vm, SumOverRectangularRange) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));  // A1 = 1
  s.set_cell_value(0U, 1U, Value::number(2.0));  // B1 = 2
  s.set_cell_value(1U, 0U, Value::number(3.0));  // A2 = 3
  s.set_cell_value(1U, 1U, Value::number(4.0));  // B2 = 4
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value vm = RunVmWithCtx("=SUM(A1:B2)", a, ctx);
  ASSERT_TRUE(vm.is_number());
  EXPECT_DOUBLE_EQ(vm.as_number(), 10.0);
  Arena a2;
  const Value tree = RunTreeWithCtx("=SUM(A1:B2)", a2, ctx);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm));
}

TEST(Vm, SumOverColumnRange) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(10.0));  // A1
  s.set_cell_value(1U, 0U, Value::number(20.0));  // A2
  s.set_cell_value(2U, 0U, Value::number(30.0));  // A3
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value vm = RunVmWithCtx("=SUM(A1:A3)", a, ctx);
  ASSERT_TRUE(vm.is_number());
  EXPECT_DOUBLE_EQ(vm.as_number(), 60.0);
}

TEST(Vm, RangeAverageAgreesWithTreeWalker) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(2.0));   // A1
  s.set_cell_value(0U, 1U, Value::number(4.0));   // B1
  s.set_cell_value(1U, 0U, Value::number(6.0));   // A2
  s.set_cell_value(1U, 1U, Value::number(12.0));  // B2
  EvalState state;
  EvalContext ctx(wb, s, state);
  Arena a;
  const Value vm = RunVmWithCtx("=AVERAGE(A1:B2)", a, ctx);
  Arena a2;
  const Value tree = RunTreeWithCtx("=AVERAGE(A1:B2)", a2, ctx);
  ASSERT_TRUE(vm.is_number());
  EXPECT_DOUBLE_EQ(vm.as_number(), 6.0);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm));
}

// ---------------------------------------------------------------------------
// Recursive LAMBDA + inner LET: per-invocation slot isolation
// ---------------------------------------------------------------------------

TEST(Vm, RecursiveLambdaWithInnerLetDoesNotClobberSlots) {
  // Self-passing recursion (Excel's LAMBDA recursion idiom). The inner
  // `LET(t, n, ...)` reuses the same LET slot on every recursion depth, and
  // `t` is read AFTER the recursive `self(self, n-1)` call. Without
  // per-invocation slot isolation the deeper call would clobber `t`, yielding
  // a wrong product; the correct factorial of 5 is 120.
  constexpr const char* kSrc = "=LET(f, LAMBDA(self, n, IF(n<=1, 1, LET(t, n, self(self, n-1) * t))), f(f, 5))";
  Arena a;
  const Value vm = RunVmOrDie(kSrc, a);
  ASSERT_TRUE(vm.is_number()) << "vm=" << vm.debug_to_string();
  EXPECT_DOUBLE_EQ(vm.as_number(), 120.0);
  Arena a2;
  const Value tree = RunTreeOrDie(kSrc, a2);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm));
}

TEST(Vm, RecursiveLambdaSumWithInnerLetAgreesWithTreeWalker) {
  // Recursive summation 5+4+3+2+1 = 15 with an inner LET binding read after
  // the recursive self-call. Cross-checks the slot-isolation fix against the
  // tree-walker.
  constexpr const char* kSrc = "=LET(f, LAMBDA(self, n, IF(n<=0, 0, LET(t, n, self(self, n-1) + t))), f(f, 5))";
  Arena a;
  const Value vm = RunVmOrDie(kSrc, a);
  ASSERT_TRUE(vm.is_number()) << "vm=" << vm.debug_to_string();
  EXPECT_DOUBLE_EQ(vm.as_number(), 15.0);
  Arena a2;
  const Value tree = RunTreeOrDie(kSrc, a2);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm));
}

// ---------------------------------------------------------------------------
// Implicit intersection (@ operator) parity with the tree-walker
// ---------------------------------------------------------------------------

TEST(Vm, ImplicitIntersectionColumnProjectsFormulaRow) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(10.0));  // A1
  s.set_cell_value(1U, 0U, Value::number(20.0));  // A2
  s.set_cell_value(2U, 0U, Value::number(30.0));  // A3
  EvalState state;
  // Formula cell anchored on the A2 row: a single-column range projects it.
  const EvalContext ctx = EvalContext(wb, s, state).with_formula_cell(1U, 5U);
  Arena a;
  const Value vm = RunVmWithCtx("=@A1:A3", a, ctx);
  ASSERT_TRUE(vm.is_number()) << vm.debug_to_string();
  EXPECT_DOUBLE_EQ(vm.as_number(), 20.0);
  Arena a2;
  const Value tree = RunTreeWithCtx("=@A1:A3", a2, ctx);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm));
}

TEST(Vm, ImplicitIntersectionOutsideRangeIsValueError) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(10.0));
  s.set_cell_value(1U, 0U, Value::number(20.0));
  s.set_cell_value(2U, 0U, Value::number(30.0));
  EvalState state;
  // Formula row 5 is outside A1:A3 -> #VALUE!.
  const EvalContext ctx = EvalContext(wb, s, state).with_formula_cell(5U, 5U);
  Arena a;
  const Value vm = RunVmWithCtx("=@A1:A3", a, ctx);
  ASSERT_TRUE(vm.is_error());
  EXPECT_EQ(vm.as_error(), ErrorCode::Value);
  Arena a2;
  const Value tree = RunTreeWithCtx("=@A1:A3", a2, ctx);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm));
}

TEST(Vm, ImplicitIntersection2DProjectsIntersectionCell) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));  // A1
  s.set_cell_value(0U, 1U, Value::number(2.0));  // B1
  s.set_cell_value(1U, 0U, Value::number(3.0));  // A2
  s.set_cell_value(1U, 1U, Value::number(4.0));  // B2
  EvalState state;
  // Formula cell at (row 1, col 1) = B2: the 2D range intersects to that cell.
  const EvalContext ctx = EvalContext(wb, s, state).with_formula_cell(1U, 1U);
  Arena a;
  const Value vm = RunVmWithCtx("=@A1:B2", a, ctx);
  ASSERT_TRUE(vm.is_number()) << vm.debug_to_string();
  EXPECT_DOUBLE_EQ(vm.as_number(), 4.0);
  Arena a2;
  const Value tree = RunTreeWithCtx("=@A1:B2", a2, ctx);
  EXPECT_TRUE(ValuesAgreeBitExact(tree, vm));
}

// ---------------------------------------------------------------------------
// VM fault surface
// ---------------------------------------------------------------------------

TEST(Vm, EmptyBytecodeReturnsError) {
  ByteCode bc;  // no code
  Arena a;
  auto out = execute(bc, a, default_registry(), test::mac_context());
  ASSERT_FALSE(out.has_value());
  EXPECT_EQ(out.error().code, FormulonErrorCode::kVmEmptyBytecode);
}

TEST(Vm, StackUnderflowReturnsError) {
  // Hand-roll a bytecode body that pops without pushing: a Return on empty
  // stack. The compiler never emits this, but the VM must surface the
  // underflow rather than crash.
  ByteCode bc;
  Instruction ret{};
  ret.op = OpCode::Return;
  bc.code.push_back(ret);
  bc.source_pos.push_back(0U);
  Arena a;
  auto out = execute(bc, a, default_registry(), test::mac_context());
  ASSERT_FALSE(out.has_value());
  EXPECT_EQ(out.error().code, FormulonErrorCode::kVmStackUnderflow);
}

TEST(Vm, UnknownOpcodeReturnsErrorInsteadOfHanging) {
  // A real opcode always advances `pc` itself; an unrecognised opcode would
  // leave `pc` unchanged and spin the dispatch loop forever. The default
  // switch arm must fail fast instead. Hand-roll a body whose sole
  // instruction carries an out-of-range opcode value.
  ByteCode bc;
  Instruction bad{};
  bad.op = static_cast<OpCode>(0xFEU);  // not a defined OpCode
  bc.code.push_back(bad);
  bc.source_pos.push_back(0U);
  Arena a;
  auto out = execute(bc, a, default_registry(), test::mac_context());
  ASSERT_FALSE(out.has_value());
  EXPECT_EQ(out.error().code, FormulonErrorCode::kVmInvalidOpcode);
}

TEST(Vm, BinaryOpUnderflowReturnsError) {
  // A single LoadConst followed by BinaryOp pops 2 with only 1 on the stack.
  ByteCode bc;
  bc.constants.push_back(Value::number(1.0));
  Instruction ld{};
  ld.op = OpCode::LoadConst;
  ld.a = 0U;
  bc.code.push_back(ld);
  bc.source_pos.push_back(0U);
  Instruction bo{};
  bo.op = OpCode::BinaryOp;
  bo.a = static_cast<std::uint32_t>(parser::BinOp::Add);
  bc.code.push_back(bo);
  bc.source_pos.push_back(0U);
  Arena a;
  auto out = execute(bc, a, default_registry(), test::mac_context());
  ASSERT_FALSE(out.has_value());
  EXPECT_EQ(out.error().code, FormulonErrorCode::kVmStackUnderflow);
}

TEST(Vm, ResultMatchesTreeWalker_ScalarSum) {
  Arena a1;
  Arena a2;
  EXPECT_TRUE(ValuesAgreeBitExact(RunTreeOrDie("=1+2*3", a1), RunVmOrDie("=1+2*3", a2)));
}

TEST(Vm, ResultMatchesTreeWalker_Function) {
  Arena a1;
  Arena a2;
  EXPECT_TRUE(ValuesAgreeBitExact(RunTreeOrDie("=ABS(-7)+SQRT(16)", a1), RunVmOrDie("=ABS(-7)+SQRT(16)", a2)));
}

TEST(Vm, ResultMatchesTreeWalker_Concat) {
  Arena a1;
  Arena a2;
  EXPECT_TRUE(
      ValuesAgreeBitExact(RunTreeOrDie("=\"a\" & \"b\" & \"c\"", a1), RunVmOrDie("=\"a\" & \"b\" & \"c\"", a2)));
}

TEST(Vm, ResultMatchesTreeWalker_LetLambdaCompose) {
  Arena a1;
  Arena a2;
  EXPECT_TRUE(ValuesAgreeBitExact(RunTreeOrDie("=LET(f, LAMBDA(x, x*x), f(7))", a1),
                                  RunVmOrDie("=LET(f, LAMBDA(x, x*x), f(7))", a2)));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
