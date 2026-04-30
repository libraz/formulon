// Copyright 2026 libraz. Licensed under the MIT License.
//
// Shape / round-trip tests for `eval::ByteCode` and `eval::Instruction`.
// These tests target the IR header in isolation: they construct
// `Instruction` and `ByteCode` instances by hand, asserting field
// invariants and `bytecode_shapes_equal()` behaviour. The companion
// `compiler_test.cpp` covers AST -> ByteCode lowering proper.

#include "eval/bytecode.h"

#include <cstdint>
#include <string>
#include <type_traits>

#include "gtest/gtest.h"
#include "parser/reference.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

TEST(InstructionLayout, IsTriviallyCopyable) {
  static_assert(std::is_trivially_copyable_v<Instruction>);
  Instruction i{};
  EXPECT_EQ(i.op, OpCode::Halt);
  EXPECT_EQ(i.a, 0u);
  EXPECT_EQ(i.b, 0u);
  EXPECT_EQ(i.pad, 0u);
  EXPECT_EQ(i.flags, 0u);
}

TEST(InstructionLayout, MaxAOperandIs24Bits) {
  EXPECT_EQ(Instruction::kMaxA, (1u << 24) - 1u);
}

TEST(InstructionLayout, AssignsFieldsCorrectly) {
  Instruction i{};
  i.op = OpCode::LoadConst;
  i.a = 0x123456u;
  i.b = 0xDEADBEEFu;
  EXPECT_EQ(i.op, OpCode::LoadConst);
  EXPECT_EQ(i.a, 0x123456u);
  EXPECT_EQ(i.b, 0xDEADBEEFu);
}

TEST(InstructionLayout, OpcodeNamesCoverAllVariants) {
  // Spot-check a handful and then sweep every defined opcode to ensure the
  // helper does not return the fallback "?" for any in-range value.
  EXPECT_STREQ(opcode_name(OpCode::LoadConst), "LoadConst");
  EXPECT_STREQ(opcode_name(OpCode::Halt), "Halt");
  for (std::uint8_t v = 0; v <= static_cast<std::uint8_t>(OpCode::Halt); ++v) {
    const char* name = opcode_name(static_cast<OpCode>(v));
    ASSERT_NE(name, nullptr);
    EXPECT_STRNE(name, "?") << "opcode " << static_cast<int>(v) << " missing from opcode_name()";
  }
}

TEST(ByteCodeShape, EmptyInstancesCompareEqual) {
  ByteCode a;
  ByteCode b;
  EXPECT_TRUE(bytecode_shapes_equal(a, b));
}

TEST(ByteCodeShape, DistinctInstructionsCompareUnequal) {
  ByteCode a;
  Instruction i{};
  i.op = OpCode::LoadConst;
  i.a = 7;
  a.code.push_back(i);

  ByteCode b;
  EXPECT_FALSE(bytecode_shapes_equal(a, b));

  Instruction j = i;
  j.a = 8;
  b.code.push_back(j);
  EXPECT_FALSE(bytecode_shapes_equal(a, b));
}

TEST(ByteCodeShape, IdenticalConstantsAreEqual) {
  ByteCode a;
  ByteCode b;
  a.constants.push_back(Value::number(42.0));
  b.constants.push_back(Value::number(42.0));
  EXPECT_TRUE(bytecode_shapes_equal(a, b));

  a.constants.push_back(Value::boolean(true));
  EXPECT_FALSE(bytecode_shapes_equal(a, b));

  b.constants.push_back(Value::boolean(false));
  EXPECT_FALSE(bytecode_shapes_equal(a, b));

  // Replacing element with equal payload restores equality.
  b.constants[1] = Value::boolean(true);
  EXPECT_TRUE(bytecode_shapes_equal(a, b));
}

TEST(ByteCodeShape, NamesPoolMatters) {
  ByteCode a;
  ByteCode b;
  a.names.emplace_back("SUM");
  b.names.emplace_back("SUM");
  EXPECT_TRUE(bytecode_shapes_equal(a, b));
  b.names[0] = "SUMX";
  EXPECT_FALSE(bytecode_shapes_equal(a, b));
}

TEST(ByteCodeShape, RefsPoolMatters) {
  ByteCode a;
  ByteCode b;
  parser::Reference r;
  r.col = 1;
  r.row = 2;
  a.refs.push_back(r);
  b.refs.push_back(r);
  EXPECT_TRUE(bytecode_shapes_equal(a, b));
  b.refs[0].row = 3;
  EXPECT_FALSE(bytecode_shapes_equal(a, b));
}

TEST(ByteCodeShape, SourcePosIsExcludedFromEquality) {
  // Two ByteCodes that differ only in their source_pos map should still be
  // considered equal: source attribution is intentionally not part of the
  // shape contract.
  ByteCode a;
  ByteCode b;
  a.source_pos = {10, 20, 30};
  b.source_pos = {1, 2, 3};
  EXPECT_TRUE(bytecode_shapes_equal(a, b));
}

TEST(ByteCodeShape, StringStorageIsPartOfEquality) {
  ByteCode a;
  ByteCode b;
  a.string_storage.emplace_back("hello");
  b.string_storage.emplace_back("hello");
  EXPECT_TRUE(bytecode_shapes_equal(a, b));
  b.string_storage[0] = "HELLO";
  EXPECT_FALSE(bytecode_shapes_equal(a, b));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
