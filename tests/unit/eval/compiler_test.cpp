// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lowering tests for `eval::compile()`. Each test parses a small formula
// string, runs it through the compiler, and asserts the emitted opcode
// sequence (or pool contents) matches the expected shape.
//
// The tests intentionally inspect the bytecode at a coarse granularity
// (opcode + key operand fields) rather than exact `(a, b)` tuples for
// every instruction, because operand encoding for some opcodes (e.g.
// `LoadStructRef`'s packed `(modifier, col_idx)`) is documented as
// "private to the compiler / VM agreement" and may evolve as Bundle 5.2
// lands.

#include "eval/compiler.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/bytecode.h"
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

// Parses `src` (which must include the leading `=` per the parser's
// contract) and returns the AST root. Aborts the test if parsing fails.
parser::AstNode* ParseOrDie(Arena& arena, std::string_view src) {
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  EXPECT_TRUE(p.errors().empty()) << "unexpected parse errors for: " << src;
  return root;
}

// Convenience wrapper: parse + compile, abort on either failure.
ByteCode CompileOrDie(std::string_view src) {
  Arena a;
  parser::AstNode* root = ParseOrDie(a, src);
  auto bc = compile(*root, a);
  EXPECT_TRUE(bc.has_value()) << "compile failed for: " << src << " err=" << (bc.has_value() ? "" : bc.error().message);
  if (!bc.has_value()) {
    return {};
  }
  return std::move(bc.value());
}

// Returns the sequence of opcodes in `bc.code` for a quick shape match.
std::vector<OpCode> Opcodes(const ByteCode& bc) {
  std::vector<OpCode> out;
  out.reserve(bc.code.size());
  for (const auto& ins : bc.code) {
    out.push_back(ins.op);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Scalar literals
// ---------------------------------------------------------------------------

TEST(CompilerScalar, NumberLiteralEmitsLoadConst) {
  ByteCode bc = CompileOrDie("=42");
  ASSERT_EQ(bc.code.size(), 2u);  // LoadConst, Return
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[0].a, 0u);
  EXPECT_EQ(bc.code[1].op, OpCode::Return);
  ASSERT_EQ(bc.constants.size(), 1u);
  EXPECT_TRUE(bc.constants[0].is_number());
  EXPECT_DOUBLE_EQ(bc.constants[0].as_number(), 42.0);
}

TEST(CompilerScalar, BoolLiteralEmitsLoadConst) {
  ByteCode bc = CompileOrDie("=TRUE");
  ASSERT_GE(bc.code.size(), 1u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  ASSERT_EQ(bc.constants.size(), 1u);
  EXPECT_TRUE(bc.constants[0].is_boolean());
  EXPECT_TRUE(bc.constants[0].as_boolean());
}

TEST(CompilerScalar, StringLiteralInternsIntoStringStorage) {
  ByteCode bc = CompileOrDie("=\"hello\"");
  ASSERT_EQ(bc.constants.size(), 1u);
  EXPECT_TRUE(bc.constants[0].is_text());
  EXPECT_EQ(bc.constants[0].as_text(), "hello");
  // Text payload must borrow from `string_storage`, not from the parser arena.
  ASSERT_FALSE(bc.string_storage.empty());
  EXPECT_EQ(bc.string_storage[0], "hello");
}

TEST(CompilerScalar, ErrorLiteralEmitsLoadConstWithErrorValue) {
  ByteCode bc = CompileOrDie("=#DIV/0!");
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  ASSERT_EQ(bc.constants.size(), 1u);
  EXPECT_TRUE(bc.constants[0].is_error());
  EXPECT_EQ(bc.constants[0].as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// References
// ---------------------------------------------------------------------------

TEST(CompilerRef, CellRefEmitsLoadRef) {
  ByteCode bc = CompileOrDie("=A1");
  EXPECT_EQ(bc.code[0].op, OpCode::LoadRef);
  ASSERT_EQ(bc.refs.size(), 1u);
  EXPECT_EQ(bc.refs[0].col, 0u);
  EXPECT_EQ(bc.refs[0].row, 0u);
}

TEST(CompilerRef, RangeEmitsTwoLoadRefsAndLoadRangeOp) {
  ByteCode bc = CompileOrDie("=A1:B2");
  // Two endpoint LoadRef + LoadRange + Return.
  ASSERT_GE(bc.code.size(), 4u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadRef);
  EXPECT_EQ(bc.code[1].op, OpCode::LoadRef);
  EXPECT_EQ(bc.code[2].op, OpCode::LoadRange);
  EXPECT_EQ(bc.code[3].op, OpCode::Return);
  EXPECT_EQ(bc.refs.size(), 2u);
}

TEST(CompilerRef, NameRefEmitsLoadName) {
  ByteCode bc = CompileOrDie("=MyName");
  EXPECT_EQ(bc.code[0].op, OpCode::LoadName);
  ASSERT_EQ(bc.names.size(), 1u);
  EXPECT_EQ(bc.names[0], "MyName");
}

TEST(CompilerRef, SpillRefEmitsLoadSpillRef) {
  ByteCode bc = CompileOrDie("=A1#");
  EXPECT_EQ(bc.code[0].op, OpCode::LoadSpillRef);
  ASSERT_EQ(bc.refs.size(), 1u);
}

// ---------------------------------------------------------------------------
// Calls / operators
// ---------------------------------------------------------------------------

TEST(CompilerCall, SumOneArgEmitsLoadConstAndCall) {
  ByteCode bc = CompileOrDie("=SUM(1)");
  // LoadConst 1 ; Call SUM, arity=1 ; Return
  ASSERT_EQ(bc.code.size(), 3u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[1].op, OpCode::Call);
  EXPECT_EQ(bc.code[1].b, 1u);  // arity
  EXPECT_EQ(bc.code[2].op, OpCode::Return);
  ASSERT_EQ(bc.names.size(), 1u);
  EXPECT_EQ(bc.names[0], "SUM");
}

TEST(CompilerCall, MultipleArgsCompileLeftToRight) {
  ByteCode bc = CompileOrDie("=SUM(1, 2, 3)");
  ASSERT_EQ(bc.code.size(), 5u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[1].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[2].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[3].op, OpCode::Call);
  EXPECT_EQ(bc.code[3].b, 3u);
  EXPECT_EQ(bc.code[4].op, OpCode::Return);
  ASSERT_EQ(bc.constants.size(), 3u);
  EXPECT_DOUBLE_EQ(bc.constants[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(bc.constants[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(bc.constants[2].as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// Binary operators (one test per BinOp variant + concat)
// ---------------------------------------------------------------------------

struct BinaryCase {
  std::string_view src;
  parser::BinOp op;
  OpCode expected_op;
};

class CompilerBinaryParam : public ::testing::TestWithParam<BinaryCase> {};

TEST_P(CompilerBinaryParam, EmitsExpectedOp) {
  const auto& tc = GetParam();
  ByteCode bc = CompileOrDie(tc.src);
  ASSERT_GE(bc.code.size(), 3u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[1].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[2].op, tc.expected_op);
  if (tc.expected_op == OpCode::BinaryOp) {
    EXPECT_EQ(bc.code[2].a, static_cast<std::uint32_t>(tc.op));
  }
}

INSTANTIATE_TEST_SUITE_P(AllBinOps, CompilerBinaryParam,
                         ::testing::Values(BinaryCase{"=1+2", parser::BinOp::Add, OpCode::BinaryOp},
                                           BinaryCase{"=1-2", parser::BinOp::Sub, OpCode::BinaryOp},
                                           BinaryCase{"=1*2", parser::BinOp::Mul, OpCode::BinaryOp},
                                           BinaryCase{"=1/2", parser::BinOp::Div, OpCode::BinaryOp},
                                           BinaryCase{"=2^3", parser::BinOp::Pow, OpCode::BinaryOp},
                                           BinaryCase{"=\"a\"&\"b\"", parser::BinOp::Concat, OpCode::Concat},
                                           BinaryCase{"=1=2", parser::BinOp::Eq, OpCode::BinaryOp},
                                           BinaryCase{"=1<>2", parser::BinOp::NotEq, OpCode::BinaryOp},
                                           BinaryCase{"=1<2", parser::BinOp::Lt, OpCode::BinaryOp},
                                           BinaryCase{"=1<=2", parser::BinOp::LtEq, OpCode::BinaryOp},
                                           BinaryCase{"=1>2", parser::BinOp::Gt, OpCode::BinaryOp},
                                           BinaryCase{"=1>=2", parser::BinOp::GtEq, OpCode::BinaryOp}));

// ---------------------------------------------------------------------------
// Unary operators
// ---------------------------------------------------------------------------

TEST(CompilerUnary, NegateEmitsUnaryOp) {
  ByteCode bc = CompileOrDie("=-5");
  // LoadConst 5 ; UnaryOp Minus ; Return
  ASSERT_EQ(bc.code.size(), 3u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[1].op, OpCode::UnaryOp);
  EXPECT_EQ(bc.code[1].a, static_cast<std::uint32_t>(parser::UnaryOp::Minus));
}

TEST(CompilerUnary, PercentEmitsUnaryOp) {
  ByteCode bc = CompileOrDie("=50%");
  ASSERT_GE(bc.code.size(), 3u);
  EXPECT_EQ(bc.code[1].op, OpCode::UnaryOp);
  EXPECT_EQ(bc.code[1].a, static_cast<std::uint32_t>(parser::UnaryOp::Percent));
}

TEST(CompilerUnary, PlusEmitsUnaryOp) {
  ByteCode bc = CompileOrDie("=+5");
  ASSERT_GE(bc.code.size(), 3u);
  EXPECT_EQ(bc.code[1].op, OpCode::UnaryOp);
  EXPECT_EQ(bc.code[1].a, static_cast<std::uint32_t>(parser::UnaryOp::Plus));
}

// ---------------------------------------------------------------------------
// IF / IFERROR / IFNA short-circuit lowering
// ---------------------------------------------------------------------------

TEST(CompilerLazy, IfWithThreeArgsEmitsBranchPair) {
  ByteCode bc = CompileOrDie("=IF(TRUE, 1, 2)");
  // Expected layout:
  //   0: LoadConst TRUE
  //   1: JumpIfFalse -> Lfalse
  //   2: LoadConst 1
  //   3: Jump -> Lend
  //   4: LoadConst 2     (Lfalse)
  //   5: Return          (Lend)
  ASSERT_EQ(bc.code.size(), 6u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[1].op, OpCode::JumpIfFalse);
  EXPECT_EQ(bc.code[2].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[3].op, OpCode::Jump);
  EXPECT_EQ(bc.code[4].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[5].op, OpCode::Return);
  // JumpIfFalse should land on the "else" arm (index 4).
  EXPECT_EQ(bc.code[1].a, 4u);
  // Jump should land on the join point past the else (index 5).
  EXPECT_EQ(bc.code[3].a, 5u);
  // No CALL emitted: short-circuit form uses branches, not the registry.
  for (const auto& ins : bc.code) {
    EXPECT_NE(ins.op, OpCode::Call);
  }
}

TEST(CompilerLazy, IfWithoutElseSynthesisesFalse) {
  ByteCode bc = CompileOrDie("=IF(A1, 1)");
  // The second arm is a synthesised LoadConst FALSE.
  ASSERT_GE(bc.code.size(), 6u);
  // Last constant must be a boolean false.
  bool found_false = false;
  for (const auto& v : bc.constants) {
    if (v.is_boolean() && !v.as_boolean()) {
      found_false = true;
      break;
    }
  }
  EXPECT_TRUE(found_false);
}

TEST(CompilerLazy, IfErrorEmitsCallNotBranch) {
  // IFERROR uses the named-call lowering path because its short-circuit
  // semantics depend on whether the first argument is an *error* value, not
  // on a boolean coercion. The bytecode therefore preserves the call shape.
  ByteCode bc = CompileOrDie("=IFERROR(A1, 0)");
  bool has_call = false;
  bool has_branch = false;
  for (const auto& ins : bc.code) {
    if (ins.op == OpCode::Call) {
      has_call = true;
    }
    if (ins.op == OpCode::JumpIfFalse) {
      has_branch = true;
    }
  }
  EXPECT_TRUE(has_call);
  EXPECT_FALSE(has_branch);
  ASSERT_FALSE(bc.names.empty());
  EXPECT_EQ(bc.names.back(), "IFERROR");
}

TEST(CompilerLazy, IfNaIsCaseInsensitive) {
  ByteCode bc = CompileOrDie("=ifna(A1, 0)");
  bool has_call = false;
  for (const auto& ins : bc.code) {
    if (ins.op == OpCode::Call) {
      has_call = true;
    }
  }
  EXPECT_TRUE(has_call);
}

TEST(CompilerLazy, IfLowercaseStillBranches) {
  ByteCode bc = CompileOrDie("=if(TRUE, 1, 2)");
  bool has_branch = false;
  for (const auto& ins : bc.code) {
    if (ins.op == OpCode::JumpIfFalse) {
      has_branch = true;
    }
  }
  EXPECT_TRUE(has_branch);
}

// ---------------------------------------------------------------------------
// Array literals
// ---------------------------------------------------------------------------

TEST(CompilerArray, RowVectorEmitsMakeArray) {
  ByteCode bc = CompileOrDie("={1,2,3}");
  // LoadConst x 3 ; MakeArray rows=1 cols=3 ; Return
  ASSERT_EQ(bc.code.size(), 5u);
  EXPECT_EQ(bc.code[0].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[1].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[2].op, OpCode::LoadConst);
  EXPECT_EQ(bc.code[3].op, OpCode::MakeArray);
  EXPECT_EQ(bc.code[3].a, 1u);
  EXPECT_EQ(bc.code[3].b, 3u);
  EXPECT_EQ(bc.code[4].op, OpCode::Return);
}

TEST(CompilerArray, TwoDArrayEmitsRowMajorMakeArray) {
  ByteCode bc = CompileOrDie("={1,2;3,4}");
  // 4 LoadConsts (row-major) + MakeArray 2 2 + Return
  ASSERT_EQ(bc.code.size(), 6u);
  EXPECT_EQ(bc.code[4].op, OpCode::MakeArray);
  EXPECT_EQ(bc.code[4].a, 2u);
  EXPECT_EQ(bc.code[4].b, 2u);
}

// ---------------------------------------------------------------------------
// LET / LAMBDA / LambdaCall
// ---------------------------------------------------------------------------

TEST(CompilerLet, BindingAllocatesSlotAndStores) {
  ByteCode bc = CompileOrDie("=LET(x, 1, x+2)");
  // Expected:
  //   LoadConst 1
  //   StoreLet 0
  //   LoadLet 0
  //   LoadConst 2
  //   BinaryOp Add
  //   Return
  auto ops = Opcodes(bc);
  ASSERT_EQ(ops.size(), 6u);
  EXPECT_EQ(ops[0], OpCode::LoadConst);
  EXPECT_EQ(ops[1], OpCode::StoreLet);
  EXPECT_EQ(ops[2], OpCode::LoadLet);
  EXPECT_EQ(ops[3], OpCode::LoadConst);
  EXPECT_EQ(ops[4], OpCode::BinaryOp);
  EXPECT_EQ(ops[5], OpCode::Return);
  EXPECT_EQ(bc.code[1].a, 0u);  // StoreLet slot
  EXPECT_EQ(bc.code[2].a, 0u);  // LoadLet same slot
}

TEST(CompilerLet, NestedBindingsUseDistinctSlots) {
  ByteCode bc = CompileOrDie("=LET(x, 1, y, 2, x+y)");
  // Two slots allocated; the body's LoadLet ops must reference 0 and 1.
  std::vector<std::uint32_t> store_slots;
  std::vector<std::uint32_t> load_slots;
  for (const auto& ins : bc.code) {
    if (ins.op == OpCode::StoreLet) {
      store_slots.push_back(ins.a);
    }
    if (ins.op == OpCode::LoadLet) {
      load_slots.push_back(ins.a);
    }
  }
  ASSERT_EQ(store_slots.size(), 2u);
  EXPECT_EQ(store_slots[0], 0u);
  EXPECT_EQ(store_slots[1], 1u);
  ASSERT_EQ(load_slots.size(), 2u);
  EXPECT_EQ(load_slots[0], 0u);
  EXPECT_EQ(load_slots[1], 1u);
}

TEST(CompilerLambda, BodyEmitsMakeLambdaJumpReturn) {
  ByteCode bc = CompileOrDie("=LAMBDA(x, x+1)");
  // Expected:
  //   MakeLambda
  //   Jump <past body>
  //   LoadLambdaArg 0
  //   LoadConst 1
  //   BinaryOp Add
  //   Return        (end of lambda body)
  //   Return        (end of outer formula)
  auto ops = Opcodes(bc);
  ASSERT_GE(ops.size(), 7u);
  EXPECT_EQ(ops[0], OpCode::MakeLambda);
  EXPECT_EQ(ops[1], OpCode::Jump);
  EXPECT_EQ(ops[2], OpCode::LoadLambdaArg);
  EXPECT_EQ(bc.code[2].a, 0u);  // arg slot 0
  EXPECT_EQ(ops[3], OpCode::LoadConst);
  EXPECT_EQ(ops[4], OpCode::BinaryOp);
  EXPECT_EQ(ops[5], OpCode::Return);
  EXPECT_EQ(ops[6], OpCode::Return);
  // The Jump should skip past the body to land on the final Return.
  EXPECT_EQ(bc.code[1].a, 6u);
}

TEST(CompilerLambdaCall, EmitsCallLambda) {
  ByteCode bc = CompileOrDie("=LAMBDA(x, x+1)(5)");
  bool has_call_lambda = false;
  for (const auto& ins : bc.code) {
    if (ins.op == OpCode::CallLambda) {
      has_call_lambda = true;
      EXPECT_EQ(ins.a, 1u);  // arity
    }
  }
  EXPECT_TRUE(has_call_lambda);
}

// ---------------------------------------------------------------------------
// ErrorPlaceholder rejection
// ---------------------------------------------------------------------------

TEST(CompilerError, ErrorPlaceholderIsRejectedWithSpecificCode) {
  // Construct a synthetic placeholder directly (bypassing the parser) so the
  // test exercises the compiler's error-handling path deterministically.
  Arena a;
  parser::AstNode* placeholder = parser::make_error_placeholder(a);
  ASSERT_NE(placeholder, nullptr);
  auto result = compile(*placeholder, a);
  ASSERT_FALSE(result.has_value());
  EXPECT_EQ(result.error().code, FormulonErrorCode::kVmUnsupportedNode);
}

// ---------------------------------------------------------------------------
// Determinism
// ---------------------------------------------------------------------------

TEST(CompilerDeterminism, SameSourceCompilesToEqualBytecode) {
  Arena a1;
  Arena a2;
  parser::AstNode* r1 = ParseOrDie(a1, "=SUM(1, 2, 3) + 4");
  parser::AstNode* r2 = ParseOrDie(a2, "=SUM(1, 2, 3) + 4");
  auto bc1 = compile(*r1, a1);
  auto bc2 = compile(*r2, a2);
  ASSERT_TRUE(bc1.has_value());
  ASSERT_TRUE(bc2.has_value());
  EXPECT_TRUE(bytecode_shapes_equal(bc1.value(), bc2.value()));
}

// ---------------------------------------------------------------------------
// Source-position map parallel to instruction stream
// ---------------------------------------------------------------------------

TEST(CompilerSourcePos, ParallelToInstructionStream) {
  ByteCode bc = CompileOrDie("=1+2");
  EXPECT_EQ(bc.source_pos.size(), bc.code.size());
}

}  // namespace
}  // namespace eval
}  // namespace formulon
