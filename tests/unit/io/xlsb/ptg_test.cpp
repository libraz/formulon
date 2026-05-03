// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for the MS-XLSB Ptg dispatch table.

#include "io/xlsb/ptg.h"

#include <cstdint>

#include "gtest/gtest.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

TEST(XlsbPtg, TableIsSortedByBaseByte) {
  // Compile-time asserted in the .cpp; assert again at runtime so an
  // accidental edit that survives the compile-time check (e.g.
  // duplicate base bytes) still trips the test suite.
  for (std::size_t i = 1; i < kPtgInfoCount; ++i) {
    EXPECT_LT(kPtgInfoTable[i - 1].base_byte, kPtgInfoTable[i].base_byte)
        << "row " << i << " out of order: 0x" << std::hex << static_cast<int>(kPtgInfoTable[i - 1].base_byte)
        << " >= 0x" << static_cast<int>(kPtgInfoTable[i].base_byte);
  }
}

TEST(XlsbPtg, LookupRecognisesRepresentativeOperands) {
  // PtgInt (0x1E) — fixed integer literal.
  const PtgInfo* p = lookup_ptg(0x1E);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(p->kind, PtgKind::Int);
  EXPECT_STREQ(p->name, "Int");

  // PtgNum (0x1F) — IEEE 754 double.
  p = lookup_ptg(0x1F);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(p->kind, PtgKind::Num);

  // PtgAttr (0x19) — function-call attribute.
  p = lookup_ptg(0x19);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(p->kind, PtgKind::Attr);

  // PtgErr (0x1C) — error literal.
  p = lookup_ptg(0x1C);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(p->kind, PtgKind::Err);
}

TEST(XlsbPtg, LookupRecognisesClassMarkedRefAndFunc) {
  // PtgRef base byte (0x24).
  const PtgInfo* ref = lookup_ptg(0x24);
  ASSERT_NE(ref, nullptr);
  EXPECT_EQ(ref->kind, PtgKind::Ref);
  EXPECT_TRUE(is_class_marked(ref->kind));

  // PtgFunc base byte (0x21).
  const PtgInfo* func = lookup_ptg(0x21);
  ASSERT_NE(func, nullptr);
  EXPECT_EQ(func->kind, PtgKind::Func);
  EXPECT_TRUE(is_class_marked(func->kind));

  // PtgFuncVar base byte (0x22).
  const PtgInfo* fvar = lookup_ptg(0x22);
  ASSERT_NE(fvar, nullptr);
  EXPECT_EQ(fvar->kind, PtgKind::FuncVar);
}

TEST(XlsbPtg, LookupRecognisesExtensionPtgIfError) {
  const PtgInfo* p = lookup_ptg(0xEA);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(p->kind, PtgKind::IfError);
  EXPECT_FALSE(is_class_marked(p->kind));
}

TEST(XlsbPtg, LookupReturnsNullForUnknownByte) {
  EXPECT_EQ(lookup_ptg(0x00), nullptr);
  EXPECT_EQ(lookup_ptg(0x1A), nullptr);  // reserved
  EXPECT_EQ(lookup_ptg(0x1B), nullptr);  // reserved
  EXPECT_EQ(lookup_ptg(0x80), nullptr);
  EXPECT_EQ(lookup_ptg(0xFF), nullptr);
}

TEST(XlsbPtg, ClassFromByteDecodesTrioCorrectly) {
  // PtgRef base (0x24) = Reference class.
  EXPECT_EQ(class_from_byte(0x24), PtgClass::Reference);
  // PtgRef value-class (0x44) = Value.
  EXPECT_EQ(class_from_byte(0x44), PtgClass::Value);
  // PtgRef array-class (0x64) = Array.
  EXPECT_EQ(class_from_byte(0x64), PtgClass::Array);

  // PtgFunc trio.
  EXPECT_EQ(class_from_byte(0x21), PtgClass::Reference);
  EXPECT_EQ(class_from_byte(0x41), PtgClass::Value);
  EXPECT_EQ(class_from_byte(0x61), PtgClass::Array);
}

TEST(XlsbPtg, LookupFromWireStripsClassBits) {
  // 0x44 (PtgRef value-class) should resolve to the same row as 0x24.
  const PtgInfo* a = lookup_ptg_from_wire(0x24);
  const PtgInfo* b = lookup_ptg_from_wire(0x44);
  const PtgInfo* c = lookup_ptg_from_wire(0x64);
  ASSERT_NE(a, nullptr);
  EXPECT_EQ(a, b);
  EXPECT_EQ(a, c);
  EXPECT_EQ(a->kind, PtgKind::Ref);

  // PtgFunc trio.
  EXPECT_EQ(lookup_ptg_from_wire(0x21), lookup_ptg_from_wire(0x41));
  EXPECT_EQ(lookup_ptg_from_wire(0x21), lookup_ptg_from_wire(0x61));

  // Unclassed byte (0x1E PtgInt) is not class-stripped.
  const PtgInfo* p = lookup_ptg_from_wire(0x1E);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(p->kind, PtgKind::Int);
}

TEST(XlsbPtg, IsClassMarkedDistinguishesOperatorsFromOperands) {
  // Class-marked operands.
  EXPECT_TRUE(is_class_marked(PtgKind::Ref));
  EXPECT_TRUE(is_class_marked(PtgKind::Area));
  EXPECT_TRUE(is_class_marked(PtgKind::Func));
  EXPECT_TRUE(is_class_marked(PtgKind::Array));

  // Non-class-marked operators / literals.
  EXPECT_FALSE(is_class_marked(PtgKind::Int));
  EXPECT_FALSE(is_class_marked(PtgKind::Num));
  EXPECT_FALSE(is_class_marked(PtgKind::Add));
  EXPECT_FALSE(is_class_marked(PtgKind::Attr));
  EXPECT_FALSE(is_class_marked(PtgKind::IfError));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
