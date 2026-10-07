#include "eval/tail_array.h"

#include <gtest/gtest.h>

#include "eval/array_alloc.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

double num(const Value& v) {
  return v.as_number();
}

TEST(TailArray, ReadsHeadAndTailOnRowAxis) {
  Arena arena;
  // 6x2, head 2: head rows (1,2),(3,4); tail row (9,8).
  const Value cells[] = {Value::number(1), Value::number(2), Value::number(3), Value::number(4)};
  const Value tail[] = {Value::number(9), Value::number(8)};
  const TailArray* ta = make_tail_array(arena, 6, 2, 2, TailAxis::kRows, cells, tail, true);
  ASSERT_NE(ta, nullptr);
  EXPECT_EQ(num(tail_array_at(*ta, 0, 0)), 1);
  EXPECT_EQ(num(tail_array_at(*ta, 1, 1)), 4);
  EXPECT_EQ(num(tail_array_at(*ta, 2, 0)), 9);
  EXPECT_EQ(num(tail_array_at(*ta, 5, 1)), 8);
}

TEST(TailArray, ReadsHeadAndTailOnColAxis) {
  Arena arena;
  // 2x5, head 2: rows (1,2),(3,4); tail column (9,8).
  const Value cells[] = {Value::number(1), Value::number(2), Value::number(3), Value::number(4)};
  const Value tail[] = {Value::number(9), Value::number(8)};
  const TailArray* ta = make_tail_array(arena, 2, 5, 2, TailAxis::kCols, cells, tail, false);
  ASSERT_NE(ta, nullptr);
  EXPECT_FALSE(ta->from_reference);
  EXPECT_EQ(num(tail_array_at(*ta, 0, 1)), 2);
  EXPECT_EQ(num(tail_array_at(*ta, 1, 0)), 3);
  EXPECT_EQ(num(tail_array_at(*ta, 0, 2)), 9);
  EXPECT_EQ(num(tail_array_at(*ta, 1, 4)), 8);
}

TEST(TailArray, ZeroHeadReadsOnlyTail) {
  Arena arena;
  const Value tail[] = {Value::number(7)};
  const TailArray* ta = make_tail_array(arena, 1048576, 1, 0, TailAxis::kRows, nullptr, tail, true);
  ASSERT_NE(ta, nullptr);
  EXPECT_EQ(num(tail_array_at(*ta, 0, 0)), 7);
  EXPECT_EQ(num(tail_array_at(*ta, 1048575, 0)), 7);
}

TEST(TailArray, DensifyRowAxis) {
  Arena arena;
  const Value cells[] = {Value::number(1), Value::number(2), Value::number(3), Value::number(4)};
  const Value tail[] = {Value::number(9), Value::number(8)};
  Shaped s;
  s.tail_array = make_tail_array(arena, 6, 2, 2, TailAxis::kRows, cells, tail, true);
  const Value v = densify(s, arena);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 6U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  const Value* d = v.as_array_cells();
  const double expect[] = {1, 2, 3, 4, 9, 8, 9, 8, 9, 8, 9, 8};
  for (int i = 0; i < 12; ++i) {
    EXPECT_EQ(num(d[i]), expect[i]) << i;
  }
}

TEST(TailArray, DensifyColAxis) {
  Arena arena;
  const Value cells[] = {Value::number(1), Value::number(2), Value::number(3), Value::number(4)};
  const Value tail[] = {Value::number(9), Value::number(8)};
  Shaped s;
  s.tail_array = make_tail_array(arena, 2, 4, 2, TailAxis::kCols, cells, tail, false);
  const Value v = densify(s, arena);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  const Value* d = v.as_array_cells();
  const double expect[] = {1, 2, 9, 9, 3, 4, 8, 8};
  for (int i = 0; i < 8; ++i) {
    EXPECT_EQ(num(d[i]), expect[i]) << i;
  }
}

TEST(TailArray, DensifyPassesPlainValueThrough) {
  Arena arena;
  Shaped s;
  s.value = Value::number(42);
  const Value v = densify(s, arena);
  ASSERT_FALSE(v.is_array());
  EXPECT_EQ(num(v), 42);
}

TEST(TailArray, DensifyOverCapIsNum) {
  Arena arena;
  const Value tail[] = {Value::number(1), Value::number(1)};
  // 1048576 x 16 cells exceeds kMaxDerivedArrayCells; nothing that size is allocated.
  Shaped s;
  s.tail_array = make_tail_array(arena, 1048576, 16, 0, TailAxis::kRows, nullptr, tail, false);
  ASSERT_NE(s.tail_array, nullptr);
  ASSERT_GT(1048576ULL * 16ULL, kMaxDerivedArrayCells);
  const Value v = densify(s, arena);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(TailArray, ShapedDefaultsToBlankWithoutTail) {
  const Shaped s;
  EXPECT_TRUE(s.value.is_blank());
  EXPECT_EQ(s.tail_array, nullptr);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
