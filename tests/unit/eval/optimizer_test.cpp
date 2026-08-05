//
// Tests for the bytecode optimiser (`eval::optimize`).
//
// The tests are organised by pass:
//   * Constant folding: every `BinOp` (except the deliberately-skipped
//     `Concat`) and every `UnaryOp`, plus the no-op cases (non-folding
//     sequences are preserved exactly), error propagation through fold,
//     and chained folds.
//   * Range canonicalisation: `A1:A1` collapses to `A1`, `B5:A1`
//     reorders to `A1:B5`, and the no-op `A1:B5` stays put.
//   * Branch hoisting: the IF skeleton is recognised as a missed
//     opportunity but the bytecode is unchanged.
//   * Stats counters: every counter increments only when the
//     corresponding rewrite happened.
//   * Sweep: a corpus of hand-rolled formulas, compiled with and
//     without the optimiser, executed through the VM, asserting raw-bit
//     `Number` equality. This is the local parity check for this
//     bundle.

#include "eval/optimizer.h"

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/bytecode.h"
#include "eval/compiler.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/vm.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Returns the parser AST root or aborts the test on failure.
parser::AstNode* ParseOrDie(Arena& arena, std::string_view src) {
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  EXPECT_TRUE(p.errors().empty()) << "unexpected parse errors for: " << src;
  return root;
}

// Compiles `src` (which must include the leading `=`) and returns the raw
// bytecode without optimisation. Aborts on error.
ByteCode CompileOrDie(Arena& arena, std::string_view src) {
  parser::AstNode* root = ParseOrDie(arena, src);
  if (root == nullptr) {
    return {};
  }
  auto bc = compile(*root, arena);
  EXPECT_TRUE(bc.has_value()) << "compile failed for: " << src;
  if (!bc.has_value()) {
    return {};
  }
  return std::move(bc.value());
}

ByteCode OptimizeOrDie(Arena& arena, ByteCode bc, OptimizerStats* stats = nullptr) {
  auto out = optimize(std::move(bc), arena, stats);
  EXPECT_TRUE(out.has_value()) << "optimize failed";
  if (!out.has_value()) {
    return {};
  }
  return std::move(out.value());
}

// Executes `bc` against the default function registry and an empty
// `EvalContext`. Aborts on a VM-level fault.
Value ExecuteOrDie(Arena& arena, const ByteCode& bc) {
  auto out = execute(bc, arena, default_registry(), EvalContext{});
  EXPECT_TRUE(out.has_value());
  if (!out.has_value()) {
    return Value::error(ErrorCode::Value);
  }
  return out.value();
}

// Compares two `Value`s for raw-bit equality on `Number`, identity on
// other scalar variants. Mirrors the parity harness in
// `tree_walker.cpp`.
bool ValuesAgreeBitExact(const Value& a, const Value& b) {
  if (a.kind() != b.kind()) {
    return false;
  }
  switch (a.kind()) {
    case ValueKind::Number: {
      const double x = a.as_number();
      const double y = b.as_number();
      std::uint64_t ux = 0;
      std::uint64_t uy = 0;
      std::memcpy(&ux, &x, sizeof(ux));
      std::memcpy(&uy, &y, sizeof(uy));
      return ux == uy;
    }
    case ValueKind::Bool:
      return a.as_boolean() == b.as_boolean();
    case ValueKind::Error:
      return a.as_error() == b.as_error();
    case ValueKind::Text:
      return a.as_text() == b.as_text();
    case ValueKind::Blank:
      return true;
    default:
      return false;
  }
}

// Returns the sequence of opcodes for shape assertions.
std::vector<OpCode> Opcodes(const ByteCode& bc) {
  std::vector<OpCode> out;
  out.reserve(bc.code.size());
  for (const auto& ins : bc.code) {
    out.push_back(ins.op);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Pass 1: constant folding -- arithmetic
// ---------------------------------------------------------------------------

TEST(OptimizerFold, AddTwoNumbers) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=1+2"));
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[1].op, OpCode::Return);
  // The folded constant is appended; bc.constants[0] / [1] are the
  // pre-fold operands (not removed -- pool entries are append-only).
  ASSERT_GE(bc.constants.size(), 1u);
  const Value v = bc.constants[bc.code[0].a];
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(OptimizerFold, SubTwoNumbers) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=5-3"));
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_DOUBLE_EQ(bc.constants[bc.code[0].a].as_number(), 2.0);
}

TEST(OptimizerFold, MulTwoNumbers) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=4*6"));
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_DOUBLE_EQ(bc.constants[bc.code[0].a].as_number(), 24.0);
}

TEST(OptimizerFold, DivTwoNumbers) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=10/4"));
  EXPECT_DOUBLE_EQ(bc.constants[bc.code[0].a].as_number(), 2.5);
}

TEST(OptimizerFold, DivByZeroFoldsToErrorValue) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=1/0"));
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  const Value v = bc.constants[bc.code[0].a];
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(OptimizerFold, PowFolds) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=2^10"));
  EXPECT_DOUBLE_EQ(bc.constants[bc.code[0].a].as_number(), 1024.0);
}

// ---------------------------------------------------------------------------
// Pass 1: constant folding -- comparison
// ---------------------------------------------------------------------------

TEST(OptimizerFold, EqFolds) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=3=3"));
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_TRUE(bc.constants[bc.code[0].a].as_boolean());
}

TEST(OptimizerFold, NotEqFolds) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=3<>4"));
  EXPECT_TRUE(bc.constants[bc.code[0].a].as_boolean());
}

TEST(OptimizerFold, LtFolds) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=2<5"));
  EXPECT_TRUE(bc.constants[bc.code[0].a].as_boolean());
}

TEST(OptimizerFold, LtEqFolds) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=5<=5"));
  EXPECT_TRUE(bc.constants[bc.code[0].a].as_boolean());
}

TEST(OptimizerFold, GtFolds) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=10>3"));
  EXPECT_TRUE(bc.constants[bc.code[0].a].as_boolean());
}

TEST(OptimizerFold, GtEqFolds) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=3>=3"));
  EXPECT_TRUE(bc.constants[bc.code[0].a].as_boolean());
}

// ---------------------------------------------------------------------------
// Pass 1: constant folding -- unary
// ---------------------------------------------------------------------------

TEST(OptimizerFold, UnaryMinusFolds) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=-5"));
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_DOUBLE_EQ(bc.constants[bc.code[0].a].as_number(), -5.0);
}

TEST(OptimizerFold, UnaryPlusFolds) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=+7"));
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_DOUBLE_EQ(bc.constants[bc.code[0].a].as_number(), 7.0);
}

TEST(OptimizerFold, UnaryPercentFolds) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=50%"));
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_DOUBLE_EQ(bc.constants[bc.code[0].a].as_number(), 0.5);
}

// ---------------------------------------------------------------------------
// Pass 1: constant folding -- error propagation and chained folds
// ---------------------------------------------------------------------------

TEST(OptimizerFold, ErrorOperandShortCircuits) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=#N/A+1"));
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.constants[bc.code[0].a].as_error(), ErrorCode::NA);
}

TEST(OptimizerFold, ChainedArithmeticFoldsAllTheWay) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=1+2+3+4"));
  // Without fold: 4 LoadConst, 3 BinaryOp, 1 Return = 8 ins.
  // With fold: 1 LoadConst, 1 Return = 2 ins.
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[1].op, OpCode::Return);
  EXPECT_DOUBLE_EQ(bc.constants[bc.code[0].a].as_number(), 10.0);
}

TEST(OptimizerFold, ChainedMixedArithmetic) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=2*3+4*5"));
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_DOUBLE_EQ(bc.constants[bc.code[0].a].as_number(), 26.0);
}

// ---------------------------------------------------------------------------
// Pass 1: constant folding -- non-fold cases preserve the stream
// ---------------------------------------------------------------------------

TEST(OptimizerFold, RefOperandIsNotFolded) {
  Arena a;
  ByteCode bc_raw = CompileOrDie(a, "=A1+1");
  const std::vector<OpCode> raw_ops = Opcodes(bc_raw);
  ByteCode bc = OptimizeOrDie(a, std::move(bc_raw));
  EXPECT_EQ(Opcodes(bc), raw_ops);
}

TEST(OptimizerFold, CallResultIsNotFolded) {
  Arena a;
  ByteCode bc_raw = CompileOrDie(a, "=SUM(1,2)+3");
  const std::vector<OpCode> raw_ops = Opcodes(bc_raw);
  ByteCode bc = OptimizeOrDie(a, std::move(bc_raw));
  EXPECT_EQ(Opcodes(bc), raw_ops);
}

TEST(OptimizerFold, ConcatIsSkipped) {
  // The optimiser deliberately skips Concat to keep text storage
  // self-contained without an arena dependency. The VM still produces
  // the correct result; the optimiser just leaves the IR alone.
  Arena a;
  ByteCode bc_raw = CompileOrDie(a, "=\"hello\"&\"world\"");
  const std::vector<OpCode> raw_ops = Opcodes(bc_raw);
  ByteCode bc = OptimizeOrDie(a, std::move(bc_raw));
  EXPECT_EQ(Opcodes(bc), raw_ops);
}

TEST(OptimizerFold, IfBodyConstantsFoldButBranchKept) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=IF(A1, 1+2, 3+4)"));
  // The IF lowering is:
  //   pc 0: LoadRef A1
  //   pc 1: JumpIfFalse 6
  //   pc 2-4: LoadConst 1; LoadConst 2; BinaryOp Add
  //   pc 5: Jump 9
  //   pc 6-8: LoadConst 3; LoadConst 4; BinaryOp Add
  //   pc 9: Return
  // The THEN branch (pcs 2..4) folds to a single LoadConst because
  // none of those input pcs are jump targets. The ELSE branch's first
  // instruction (pc 6) IS the JumpIfFalse target, so a single-pass
  // fold cannot collapse (6, 7, 8) without orphaning the inbound jump.
  // The optimiser correctly rejects that fold; a future fixed-point
  // pass could remap the target after fold-1 and fold the else branch
  // too, but that is out of scope for this bundle.
  //
  // Resulting shape (8 instructions):
  //   LoadRef ; JumpIfFalse ; LoadConst (folded) ; Jump ;
  //   LoadConst ; LoadConst ; BinaryOp ; Return
  ASSERT_EQ(bc.code.size(), 8u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadRef);
  EXPECT_EQ(bc.code[1].op, OpCode::JumpIfFalse);
  EXPECT_EQ(bc.code[2].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[3].op, OpCode::Jump);
  EXPECT_EQ(bc.code[4].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[5].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[6].op, OpCode::BinaryOp);
  EXPECT_EQ(bc.code[7].op, OpCode::Return);
  // JumpIfFalse must still target the start of the else branch.
  EXPECT_EQ(bc.code[1].a, 4u);
  // Jump must target Return.
  EXPECT_EQ(bc.code[3].a, 7u);
  // The folded THEN-branch constant is correct.
  EXPECT_DOUBLE_EQ(bc.constants[bc.code[2].a].as_number(), 3.0);
}

TEST(OptimizerFold, IfWithoutElseBranchPreservesShape) {
  Arena a;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=IF(A1, 1+2)"));
  // 6 instructions: LoadRef, JumpIfFalse, LoadConst (folded), Jump,
  // LoadConst (synthetic FALSE), Return.
  ASSERT_EQ(bc.code.size(), 6u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadRef);
  EXPECT_EQ(bc.code[1].op, OpCode::JumpIfFalse);
  EXPECT_EQ(bc.code[2].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[3].op, OpCode::Jump);
  EXPECT_EQ(bc.code[4].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[5].op, OpCode::Return);
}

TEST(OptimizerFold, EmptyBodyPreservedExceptForReturn) {
  // Smallest possible body: just a LoadConst + Return. No fold possible.
  Arena a;
  ByteCode bc_raw = CompileOrDie(a, "=42");
  const std::vector<OpCode> raw_ops = Opcodes(bc_raw);
  ByteCode bc = OptimizeOrDie(a, std::move(bc_raw));
  EXPECT_EQ(Opcodes(bc), raw_ops);
}

// ---------------------------------------------------------------------------
// Pass 3: range canonicalisation
// ---------------------------------------------------------------------------

TEST(OptimizerRange, NormalRangePreserved) {
  Arena a;
  ByteCode bc_raw = CompileOrDie(a, "=A1:B5");
  const std::vector<OpCode> raw_ops = Opcodes(bc_raw);
  ByteCode bc = OptimizeOrDie(a, std::move(bc_raw));
  EXPECT_EQ(Opcodes(bc), raw_ops);
}

TEST(OptimizerRange, ReversedRangeIsSwapped) {
  Arena a;
  // B5:A1 should be canonicalised so that the resulting LoadRefs
  // address A1 first, B5 second.
  ByteCode bc_raw = CompileOrDie(a, "=B5:A1");
  // Capture the pre-canon order for comparison.
  ASSERT_EQ(bc_raw.refs.size(), 2u);
  const auto pre_lhs = bc_raw.refs[bc_raw.code[0].a];
  const auto pre_rhs = bc_raw.refs[bc_raw.code[1].a];
  EXPECT_EQ(pre_lhs.col, 1u);  // B
  EXPECT_EQ(pre_rhs.col, 0u);  // A
  OptimizerStats stats;
  ByteCode bc = OptimizeOrDie(a, std::move(bc_raw), &stats);
  ASSERT_EQ(bc.code.size(), 4u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadRef);
  EXPECT_EQ(bc.code[1].op, OpCode::LoadRef);
  EXPECT_EQ(bc.code[2].op, OpCode::LoadRange);
  // After canonicalisation, the first LoadRef must point at the
  // lexicographically-earlier endpoint (A1).
  const auto post_lhs = bc.refs[bc.code[0].a];
  const auto post_rhs = bc.refs[bc.code[1].a];
  EXPECT_EQ(post_lhs.col, 0u);
  EXPECT_EQ(post_lhs.row, 0u);
  EXPECT_EQ(post_rhs.col, 1u);
  EXPECT_EQ(post_rhs.row, 4u);
  EXPECT_EQ(stats.ranges_canonicalized, 1u);
}

TEST(OptimizerRange, DegenerateRangeCollapses) {
  Arena a;
  // A1:A1 collapses to a single LoadRef A1 (the LoadRange and the
  // duplicate LoadRef are both removed).
  ByteCode bc_raw = CompileOrDie(a, "=A1:A1");
  // Sanity: pre-canon shape is `LoadRef; LoadRef; LoadRange; Return`.
  ASSERT_EQ(bc_raw.code.size(), 4u);
  EXPECT_EQ(bc_raw.code[0].op, OpCode::LoadRef);
  EXPECT_EQ(bc_raw.code[1].op, OpCode::LoadRef);
  EXPECT_EQ(bc_raw.code[2].op, OpCode::LoadRange);
  OptimizerStats stats;
  ByteCode bc = OptimizeOrDie(a, std::move(bc_raw), &stats);
  ASSERT_EQ(bc.code.size(), 2u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadRef);
  EXPECT_EQ(bc.code[1].op, OpCode::Return);
  EXPECT_EQ(stats.ranges_canonicalized, 1u);
}

// ---------------------------------------------------------------------------
// Pass 4: branch-hoist skeleton (no-op rewrite, missed-opportunity counter)
// ---------------------------------------------------------------------------

TEST(OptimizerBranchHoist, IfWithTwoConstantArmsIsRecognisedAsOpportunity) {
  // After fold, an IF over two literal constants becomes:
  //   <cond>; JumpIfFalse Lfalse; LoadConst K1; Jump Lend;
  //     LoadConst K2; Return
  // which is exactly the branch-hoist skeleton.
  Arena a;
  OptimizerStats stats;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=IF(A1, 1, 2)"), &stats);
  EXPECT_GE(stats.branch_hoist_opportunities, 1u);
  EXPECT_EQ(stats.branches_hoisted, 0u);
  // The bytecode is unchanged (skeleton only).
  EXPECT_EQ(bc.code[1].op, OpCode::JumpIfFalse);
  EXPECT_EQ(bc.code[3].op, OpCode::Jump);
}

TEST(OptimizerBranchHoist, IfWithRefArmsIsNotAnOpportunity) {
  Arena a;
  OptimizerStats stats;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=IF(A1, B1, C1)"), &stats);
  EXPECT_EQ(stats.branch_hoist_opportunities, 0u);
  EXPECT_EQ(stats.branches_hoisted, 0u);
  (void)bc;
}

// ---------------------------------------------------------------------------
// Stats counters
// ---------------------------------------------------------------------------

TEST(OptimizerStatsCounters, FoldIncrementsConstantsFolded) {
  Arena a;
  OptimizerStats stats;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=1+2+3+4"), &stats);
  // Three folds for 4 chained additions.
  EXPECT_EQ(stats.constants_folded, 3u);
  EXPECT_EQ(stats.names_inlined, 0u);
  EXPECT_EQ(stats.ranges_canonicalized, 0u);
  EXPECT_EQ(stats.branches_hoisted, 0u);
  (void)bc;
}

TEST(OptimizerStatsCounters, NamesInlinedIsAlwaysZeroForNow) {
  Arena a;
  OptimizerStats stats;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=MyName"), &stats);
  EXPECT_EQ(stats.names_inlined, 0u);
  (void)bc;
}

TEST(OptimizerStatsCounters, BranchesHoistedIsAlwaysZeroForNow) {
  Arena a;
  OptimizerStats stats;
  ByteCode bc = OptimizeOrDie(a, CompileOrDie(a, "=IF(A1, 1, 2)"), &stats);
  EXPECT_EQ(stats.branches_hoisted, 0u);
  (void)bc;
}

// ---------------------------------------------------------------------------
// Sweep: compile + optimise + run, raw-bit equality versus pre-optimise.
// This is the local parity check for Bundle 5.3.
// ---------------------------------------------------------------------------

struct SweepCase {
  std::string_view src;
};

TEST(OptimizerSweep, RawBitEqualityAcrossRepresentativeCorpus) {
  const std::vector<SweepCase> corpus = {
      // Pure-literal arithmetic: everything folds.
      {"=1+2"},
      {"=10-3"},
      {"=4*5"},
      {"=10/4"},
      {"=2^10"},
      {"=1+2*3-4/2"},
      {"=(1+2)*(3-4)"},
      {"=-7"},
      {"=+3.14"},
      {"=50%"},
      // Comparison folds.
      {"=3=3"},
      {"=3<>4"},
      {"=2<5"},
      {"=5<=5"},
      {"=10>3"},
      {"=3>=3"},
      // Error fold-through.
      {"=#N/A+1"},
      {"=1/0"},
      // Mixed literal + ref (partial fold).
      {"=A1+1"},
      {"=2*A1+3"},
      {"=A1*A2"},
      // Concat (skipped).
      {"=\"hello\"&\"world\""},
      {"=\"x\"&1"},
      // Function calls.
      {"=SUM(1,2,3)"},
      {"=SUM(1+2, 3+4)"},
      {"=ABS(-5)"},
      // IF lowering.
      {"=IF(TRUE, 1+2, 3+4)"},
      {"=IF(FALSE, 1+2, 3+4)"},
      {"=IF(A1, 1, 2)"},
      // Range canonicalisation.
      {"=A1:A1"},
      {"=B5:A1"},
      {"=A1:B5"},
  };

  std::uint32_t total_constants_folded = 0;
  std::uint32_t total_ranges_canonicalized = 0;
  std::uint32_t total_branch_hoist_opportunities = 0;

  for (const auto& tc : corpus) {
    Arena a_raw;
    ByteCode bc_raw = CompileOrDie(a_raw, tc.src);
    const Value v_raw = ExecuteOrDie(a_raw, bc_raw);

    Arena a_opt;
    ByteCode bc_raw_for_opt = CompileOrDie(a_opt, tc.src);
    OptimizerStats stats;
    ByteCode bc_opt = OptimizeOrDie(a_opt, std::move(bc_raw_for_opt), &stats);
    const Value v_opt = ExecuteOrDie(a_opt, bc_opt);

    EXPECT_TRUE(ValuesAgreeBitExact(v_raw, v_opt))
        << "divergence on src='" << tc.src << "': raw=" << v_raw.debug_to_string()
        << " opt=" << v_opt.debug_to_string();

    total_constants_folded += stats.constants_folded;
    total_ranges_canonicalized += stats.ranges_canonicalized;
    total_branch_hoist_opportunities += stats.branch_hoist_opportunities;
  }
  // Sanity: the optimiser is doing real work on this corpus.
  EXPECT_GT(total_constants_folded, 0u);
  EXPECT_GT(total_ranges_canonicalized, 0u);
  // Print the totals so the test report can be cross-checked.
  std::fprintf(stdout, "[OptimizerSweep] folds=%u canons=%u branch_opps=%u\n", total_constants_folded,
               total_ranges_canonicalized, total_branch_hoist_opportunities);
}

// ---------------------------------------------------------------------------
// compile_and_optimize: convenience wrapper sanity check.
// ---------------------------------------------------------------------------

TEST(OptimizerCompileAndOptimize, ProducesFoldedShape) {
  Arena a;
  parser::AstNode* root = ParseOrDie(a, "=1+2+3");
  ASSERT_NE(root, nullptr);
  auto bc = compile_and_optimize(*root, a);
  ASSERT_TRUE(bc.has_value());
  ASSERT_EQ(bc.value().code.size(), 2u);
  EXPECT_EQ(bc.value().code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.value().code[1].op, OpCode::Return);
  EXPECT_DOUBLE_EQ(bc.value().constants[bc.value().code[0].a].as_number(), 6.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
