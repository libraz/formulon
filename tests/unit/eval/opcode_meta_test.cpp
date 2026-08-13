//
// Consistency tests for the per-opcode metadata table.
//
// The metadata table is a single source of truth for the VM's dispatch /
// validation paths, the compiler's emit-side assertions, and any future
// disassembler. These tests pin the table's invariants so adding a new
// opcode without extending the table (or with a mis-classified entry)
// surfaces as a test failure rather than as a silent VM crash.

#include "eval/opcode_meta.h"

#include <cstdint>
#include <string_view>

#include "eval/bytecode.h"
#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace {

// Sanity: the table has one row per defined opcode value, no gaps.
TEST(OpcodeMeta, TableCoversEveryDefinedOpcode) {
  EXPECT_EQ(kOpcodeMeta.size(), static_cast<std::size_t>(OpCode::Halt) + 1U);
  for (std::uint8_t v = 0; v <= static_cast<std::uint8_t>(OpCode::Halt); ++v) {
    const auto& meta = opcode_meta(static_cast<OpCode>(v));
    ASSERT_NE(meta.name, nullptr) << "opcode " << static_cast<int>(v) << " missing name";
    EXPECT_NE(std::string_view(meta.name), std::string_view("")) << "opcode " << static_cast<int>(v) << " empty name";
  }
}

// The metadata `name` field must match the existing `opcode_name()` helper
// in `bytecode.h`. Two sources of truth, one shape: keep them in lockstep.
TEST(OpcodeMeta, NameMatchesBytecodeMnemonic) {
  for (std::uint8_t v = 0; v <= static_cast<std::uint8_t>(OpCode::Halt); ++v) {
    const auto op = static_cast<OpCode>(v);
    EXPECT_STREQ(opcode_meta(op).name, opcode_name(op));
  }
}

// Classification spot-checks: every opcode is partitioned into exactly one
// class and the partition matches what the dispatcher actually does.
TEST(OpcodeMeta, LoadOpcodesAreClassifiedLoad) {
  EXPECT_EQ(opcode_meta(OpCode::LoadConst).cls, OpClass::Load);
  EXPECT_EQ(opcode_meta(OpCode::LoadRef).cls, OpClass::Load);
  EXPECT_EQ(opcode_meta(OpCode::LoadName).cls, OpClass::Load);
  EXPECT_EQ(opcode_meta(OpCode::LoadStructRef).cls, OpClass::Load);
  EXPECT_EQ(opcode_meta(OpCode::LoadSpillRef).cls, OpClass::Load);
  EXPECT_EQ(opcode_meta(OpCode::LoadLet).cls, OpClass::Load);
  EXPECT_EQ(opcode_meta(OpCode::LoadLambdaArg).cls, OpClass::Load);
  // The convenience predicate must agree.
  EXPECT_TRUE(is_load(OpCode::LoadConst));
  EXPECT_TRUE(is_load(OpCode::LoadLambdaArg));
  EXPECT_FALSE(is_load(OpCode::StoreLet));
  EXPECT_FALSE(is_load(OpCode::Jump));
}

TEST(OpcodeMeta, StoreLetIsClassifiedStore) {
  EXPECT_EQ(opcode_meta(OpCode::StoreLet).cls, OpClass::Store);
  // Currently the only Store-classified opcode; if a new one is added the
  // partition update should be visible here.
  std::size_t store_count = 0;
  for (std::uint8_t v = 0; v <= static_cast<std::uint8_t>(OpCode::Halt); ++v) {
    if (opcode_meta(static_cast<OpCode>(v)).cls == OpClass::Store) {
      ++store_count;
    }
  }
  EXPECT_EQ(store_count, 1U);
}

TEST(OpcodeMeta, JumpOpcodesAreClassifiedJump) {
  EXPECT_EQ(opcode_meta(OpCode::Jump).cls, OpClass::Jump);
  EXPECT_EQ(opcode_meta(OpCode::JumpIfFalse).cls, OpClass::Jump);
  EXPECT_TRUE(is_jump(OpCode::Jump));
  EXPECT_TRUE(is_jump(OpCode::JumpIfFalse));
  EXPECT_FALSE(is_jump(OpCode::Call));
  EXPECT_FALSE(is_jump(OpCode::LoadConst));
  // Jump targets always live in operand A.
  EXPECT_EQ(opcode_meta(OpCode::Jump).a, OperandA::Target);
  EXPECT_EQ(opcode_meta(OpCode::JumpIfFalse).a, OperandA::Target);
}

TEST(OpcodeMeta, CallOpcodesAreClassifiedCall) {
  EXPECT_EQ(opcode_meta(OpCode::Call).cls, OpClass::Call);
  EXPECT_EQ(opcode_meta(OpCode::CallLambda).cls, OpClass::Call);
  EXPECT_TRUE(is_call(OpCode::Call));
  EXPECT_TRUE(is_call(OpCode::CallLambda));
  EXPECT_FALSE(is_call(OpCode::Jump));
  // `Call` packs the name index in A and the arity in B; `CallLambda`
  // carries only the arity in A.
  EXPECT_EQ(opcode_meta(OpCode::Call).a, OperandA::NamesIndex);
  EXPECT_EQ(opcode_meta(OpCode::Call).b, OperandB::InlineCount);
  EXPECT_EQ(opcode_meta(OpCode::CallLambda).a, OperandA::InlineCount);
  EXPECT_EQ(opcode_meta(OpCode::CallLambda).b, OperandB::None);
}

TEST(OpcodeMeta, TerminateOpcodesAreClassifiedTerminate) {
  EXPECT_EQ(opcode_meta(OpCode::Return).cls, OpClass::Terminate);
  EXPECT_EQ(opcode_meta(OpCode::Halt).cls, OpClass::Terminate);
  EXPECT_TRUE(is_terminate(OpCode::Return));
  EXPECT_TRUE(is_terminate(OpCode::Halt));
  EXPECT_FALSE(is_terminate(OpCode::Jump));
  EXPECT_FALSE(is_terminate(OpCode::Call));
  // Neither uses operand fields.
  EXPECT_EQ(opcode_meta(OpCode::Return).a, OperandA::None);
  EXPECT_EQ(opcode_meta(OpCode::Return).b, OperandB::None);
  EXPECT_EQ(opcode_meta(OpCode::Halt).a, OperandA::None);
  EXPECT_EQ(opcode_meta(OpCode::Halt).b, OperandB::None);
}

// Pool-index semantics must match what `compiler.cpp` emits for each
// opcode; if the compiler is changed to encode operands differently the
// metadata must move with it.
TEST(OpcodeMeta, PoolReferencesMatchCompilerEncoding) {
  EXPECT_EQ(opcode_meta(OpCode::LoadConst).a, OperandA::ConstantsIndex);
  EXPECT_EQ(opcode_meta(OpCode::LoadRef).a, OperandA::RefsIndex);
  EXPECT_EQ(opcode_meta(OpCode::LoadSpillRef).a, OperandA::RefsIndex);
  EXPECT_EQ(opcode_meta(OpCode::LoadName).a, OperandA::NamesIndex);
  // Structured-ref packs the column-name index plus a modifier into `b`.
  EXPECT_EQ(opcode_meta(OpCode::LoadStructRef).a, OperandA::NamesIndex);
  EXPECT_EQ(opcode_meta(OpCode::LoadStructRef).b, OperandB::Packed);
}

// Inline-tag operands carry small parser enums (BinOp / UnaryOp), not
// indices. The compiler casts the enum directly into `Instruction::a`.
TEST(OpcodeMeta, OpTagOperandsCarryParserEnums) {
  EXPECT_EQ(opcode_meta(OpCode::BinaryOp).a, OperandA::InlineOpTag);
  EXPECT_EQ(opcode_meta(OpCode::UnaryOp).a, OperandA::InlineOpTag);
  // Concat is a shorthand for `BinaryOp(BinOp::Concat)` so it has no
  // operand of its own.
  EXPECT_EQ(opcode_meta(OpCode::Concat).a, OperandA::None);
  EXPECT_EQ(opcode_meta(OpCode::Concat).b, OperandB::None);
}

// `MakeArray` is the canonical (rows, cols) inline-count pair; this is
// the only opcode where both operand fields carry plain counts.
TEST(OpcodeMeta, MakeArrayCarriesRowsAndCols) {
  EXPECT_EQ(opcode_meta(OpCode::MakeArray).a, OperandA::InlineCount);
  EXPECT_EQ(opcode_meta(OpCode::MakeArray).b, OperandB::InlineCount);
  EXPECT_EQ(opcode_meta(OpCode::MakeArray).cls, OpClass::Combine);
}

// Slot operands are LET / lambda-arg numbers. Distinct from pool indices
// so disassembler tooling can render `slot=3` rather than `name[3]`.
TEST(OpcodeMeta, SlotOperandsAreNotPoolIndices) {
  EXPECT_EQ(opcode_meta(OpCode::LoadLet).a, OperandA::Slot);
  EXPECT_EQ(opcode_meta(OpCode::StoreLet).a, OperandA::Slot);
  EXPECT_EQ(opcode_meta(OpCode::LoadLambdaArg).a, OperandA::Slot);
}

// The convenience predicates form a coarse but exhaustive partition: every
// opcode lights up exactly one of {load, store, combine, call, jump,
// terminate}. This is the property that disassemblers / optimisers rely
// on, so pin it here.
TEST(OpcodeMeta, EveryOpcodeIsInExactlyOneClass) {
  for (std::uint8_t v = 0; v <= static_cast<std::uint8_t>(OpCode::Halt); ++v) {
    const auto op = static_cast<OpCode>(v);
    const auto cls = opcode_meta(op).cls;
    int hits = 0;
    if (cls == OpClass::Load)
      ++hits;
    if (cls == OpClass::Store)
      ++hits;
    if (cls == OpClass::Combine)
      ++hits;
    if (cls == OpClass::Call)
      ++hits;
    if (cls == OpClass::Jump)
      ++hits;
    if (cls == OpClass::Terminate)
      ++hits;
    EXPECT_EQ(hits, 1) << "opcode " << opcode_meta(op).name << " is in " << hits << " classes (expected 1)";
  }
}

// The lookup function is constexpr so the compiler can fold an
// `opcode_meta(LiteralOpCode).cls` call into a constant. Pin the
// constexpr-ness here so a future refactor that drops `constexpr`
// surfaces immediately.
TEST(OpcodeMeta, AccessorIsUsableInConstantExpressions) {
  constexpr auto kLoadConstClass = opcode_meta(OpCode::LoadConst).cls;
  constexpr auto kJumpClass = opcode_meta(OpCode::Jump).cls;
  constexpr bool kReturnTerminates = is_terminate(OpCode::Return);
  EXPECT_EQ(kLoadConstClass, OpClass::Load);
  EXPECT_EQ(kJumpClass, OpClass::Jump);
  EXPECT_TRUE(kReturnTerminates);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
