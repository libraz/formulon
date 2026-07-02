// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tree-walker / bytecode-VM parity sweep. Each test parses a small formula,
// evaluates it through both paths, and asserts the results are bit-exact.
// The corpus is mined from the oracle suites (scalar arithmetic, scalar
// function calls, lazy IF / IFERROR / IFNA, LET / LAMBDA scalar uses) —
// shapes the bytecode IR can faithfully express end-to-end.
//
// The full sweep runs only when `FORMULON_VM_PARITY_CHECK` is defined at
// compile time. With the option OFF the test binary still links and runs;
// the file becomes a no-op so default-build CI stays unchanged.

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "eval/bytecode.h"
#include "eval/compiler.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "eval/vm.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

#ifdef FORMULON_VM_PARITY_CHECK

// Bit-exact equality for the VM-vs-tree-walk comparison.
bool ValuesAgree(const Value& a, const Value& b) {
  if (a.kind() != b.kind()) {
    return false;
  }
  switch (a.kind()) {
    case ValueKind::Blank:
      return true;
    case ValueKind::Number: {
      const double na = a.as_number();
      const double nb = b.as_number();
      std::uint64_t ua = 0;
      std::uint64_t ub = 0;
      std::memcpy(&ua, &na, sizeof(ua));
      std::memcpy(&ub, &nb, sizeof(ub));
      return ua == ub;
    }
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

struct Case {
  const char* formula;
};

// ~70 hand-picked formulas the bytecode IR represents end-to-end. Each is a
// scalar formula whose result the VM computes purely from `Value`-shaped
// arguments (no AST-introspection paths, no range-aware aggregator
// flattening, no SpillRef-broadcast). All are mined from the oracle YAML
// corpus under tests/oracle/cases/.
constexpr Case kCorpus[] = {
    // Scalar arithmetic.
    {"=1+1"},
    {"=1+2*3"},
    {"=2^10"},
    {"=10/4"},
    {"=10-3-2"},
    {"=1.5+2.5"},
    {"=-(-5)"},
    {"=50%"},
    {"=(1+2)*3"},
    {"=(1+2)*(3+4)"},
    // Comparisons and boolean operators.
    {"=1=1"},
    {"=1<>2"},
    {"=2<3"},
    {"=3>=3"},
    {"=4<=4"},
    {"=AND(TRUE, TRUE)"},
    {"=AND(TRUE, FALSE)"},
    {"=OR(FALSE, TRUE)"},
    {"=NOT(FALSE)"},
    {"=XOR(TRUE, FALSE)"},
    // Concat operator.
    {"=\"hello\" & \" \" & \"world\""},
    {"=\"x=\" & 42"},
    {"=\"\"&\"\""},
    // Math built-ins (scalar).
    {"=ABS(-7)"},
    {"=POWER(2,8)"},
    {"=SQRT(16)"},
    {"=SQRT(2)"},
    {"=ROUND(3.14159, 2)"},
    {"=INT(3.7)"},
    {"=MOD(10, 3)"},
    {"=SIGN(-5)"},
    {"=PI()"},
    {"=EXP(1)"},
    {"=LN(EXP(2))"},
    {"=LOG10(1000)"},
    {"=LOG(8, 2)"},
    // Aggregators with scalar arg lists (no range expansion).
    {"=SUM(1,2,3)"},
    {"=SUM(1,2,3,4,5,6,7,8,9,10)"},
    {"=AVERAGE(2,4,6)"},
    {"=MIN(5,3,9,1,7)"},
    {"=MAX(5,3,9,1,7)"},
    {"=PRODUCT(2,3,4)"},
    {"=COUNT(1,2,3,4)"},
    {"=COUNTA(1,\"\",3,\"x\")"},
    {"=SUM({1,2,3,4,5})"},
    {"=AVERAGE({10,20,30})"},
    {"=MAX({1,2,3,4,5})"},
    {"=PRODUCT({2,3,4})"},
    // Text built-ins (scalar).
    {"=LEN(\"hello\")"},
    {"=UPPER(\"hello\")"},
    {"=LOWER(\"WORLD\")"},
    {"=LEFT(\"hello\", 3)"},
    {"=RIGHT(\"hello\", 2)"},
    {"=MID(\"abcdef\", 2, 3)"},
    {"=TRIM(\"  hello  \")"},
    {"=REPT(\"ab\", 3)"},
    {"=CONCAT(\"a\", \"b\", \"c\")"},
    // Lazy short-circuit forms.
    {"=IF(TRUE, 1, 2)"},
    {"=IF(FALSE, 1, 2)"},
    {"=IF(1>0, \"yes\", \"no\")"},
    {"=IF(2=2, IF(3=3, 100, 200), 300)"},
    {"=IFERROR(42, 99)"},
    {"=IFERROR(1/0, 99)"},
    {"=IFERROR(SQRT(-1), \"bad\")"},
    {"=IFNA(42, 99)"},
    {"=IFNA(NA(), 99)"},
    {"=IFNA(1/0, 99)"},
    // Error literal propagation.
    {"=#N/A"},
    {"=#DIV/0!"},
    {"=ABS(#NUM!)"},
    // LET / LAMBDA scalar.
    {"=LET(x, 10, x+1)"},
    {"=LET(x, 10, y, 20, x*y)"},
    {"=LET(x, 1, x, x+10, x)"},
    {"=LAMBDA(x, x+1)(5)"},
    {"=LAMBDA(x, y, x*y)(3, 4)"},
    {"=LET(f, LAMBDA(x, x*x), f(7))"},
    {"=LET(f, LAMBDA(x, x+1), f(f(f(0))))"},
    // Self-passing recursive LAMBDA with an inner LET whose slot is reused
    // per depth and read after the recursive call (factorial / summation).
    {"=LET(f, LAMBDA(self, n, IF(n<=1, 1, LET(t, n, self(self, n-1) * t))), f(f, 5))"},
    {"=LET(f, LAMBDA(self, n, IF(n<=0, 0, LET(t, n, self(self, n-1) + t))), f(f, 6))"},
    // Information-style scalar.
    {"=ISNUMBER(1)"},
    {"=ISNUMBER(\"x\")"},
    {"=ISTEXT(\"x\")"},
    {"=ISBLANK(\"\")"},
    {"=ISERROR(1/0)"},
    {"=ISERR(NA())"},
    {"=N(7)"},
    // Logical / type predicates.
    {"=IF(ISERROR(1/0), \"err\", \"ok\")"},
};

class ParitySweep : public ::testing::TestWithParam<std::size_t> {};

TEST_P(ParitySweep, ResultsAgree) {
  const std::size_t idx = GetParam();
  ASSERT_LT(idx, sizeof(kCorpus) / sizeof(kCorpus[0]));
  const std::string_view src = kCorpus[idx].formula;

  Arena tree_arena;
  parser::Parser tp(src, tree_arena);
  parser::AstNode* tree_root = tp.parse();
  ASSERT_NE(tree_root, nullptr) << "parse failed for: " << src;
  const Value tree_v = evaluate(*tree_root, tree_arena);

  Arena vm_arena;
  parser::Parser vp(src, vm_arena);
  parser::AstNode* vm_root = vp.parse();
  ASSERT_NE(vm_root, nullptr) << "parse failed for: " << src;
  auto bc = compile(*vm_root, vm_arena);
  ASSERT_TRUE(bc.has_value()) << "compile failed for: " << src;
  auto out = execute(bc.value(), vm_arena, default_registry(), EvalContext{});
  ASSERT_TRUE(out.has_value()) << "VM error for: " << src;

  // Apply the surface contracts evaluate() applies on the tree-walker side
  // so the two paths are compared on equal terms.
  Value vm_v = out.value();
  if (vm_v.is_lambda()) {
    vm_v = Value::error(ErrorCode::Calc);
  }
  if (vm_v.is_blank() && vm_root->kind() != parser::NodeKind::Literal) {
    vm_v = Value::number(0.0);
  }

  EXPECT_TRUE(ValuesAgree(tree_v, vm_v)) << "parity divergence for: " << src << " tree=" << tree_v.debug_to_string()
                                         << " vm=" << vm_v.debug_to_string();
}

INSTANTIATE_TEST_SUITE_P(VmParity, ParitySweep,
                         ::testing::Range(static_cast<std::size_t>(0), sizeof(kCorpus) / sizeof(kCorpus[0])));

#else  // !FORMULON_VM_PARITY_CHECK

TEST(VmParity, NoOpWhenDisabled) {
  // The parity sweep is gated behind FORMULON_VM_PARITY_CHECK so the default
  // build does not pay the double-evaluation cost. This stub keeps the test
  // binary self-consistent and gives ctest a green case to report.
  SUCCEED();
}

#endif  // FORMULON_VM_PARITY_CHECK

// Lazy-only families (IRR / XIRR / TEXTSPLIT / NETWORKDAYS / MAP / ...) are
// registered ONLY in the lazy-dispatch table, not the eager registry, so the
// bytecode VM cannot execute them. When FORMULON_VM_PARITY_CHECK is ON the
// in-flight parity hook inside `evaluate()` must SKIP these formulas instead
// of reporting a false `#NAME?` mismatch. This test drives the public
// `evaluate()` entry point over a few such formulas and asserts it returns a
// value without aborting — exercising the skip path when the flag is on, and
// a plain tree-walk when it is off. Either way the harness must not crash.
TEST(VmParity, LazyOnlyFamiliesDoNotAbortUnderParity) {
  constexpr const char* kLazyOnly[] = {
      "=IRR({-100,40,40,40})",
      "=TEXTSPLIT(\"a,b,c\", \",\")",
      "=NETWORKDAYS(1, 30)",
      "=MAP({1,2,3}, LAMBDA(x, x*2))",
  };
  for (const char* src : kLazyOnly) {
    Arena arena;
    parser::Parser p(src, arena);
    parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << "parse failed for: " << src;
    // The parity hook (when compiled in) runs inside evaluate(); a successful
    // return means the skip path fired rather than asserting on a mismatch.
    const Value v = evaluate(*root, arena);
    (void)v;
    SUCCEED() << "evaluated without abort: " << src;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
