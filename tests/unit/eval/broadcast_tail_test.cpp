#include <gtest/gtest.h>

#include <cstdint>

#include "eval/array_alloc.h"
#include "eval/tail_array.h"
#include "eval/tree_walker/broadcast.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

constexpr std::uint32_t kGridRows = 1048576U;
constexpr std::uint32_t kGridCols = 16384U;

// 1048576x1 column with a dense head of `n` numbers and a blank tail.
Shaped column_head(Arena& arena, const double* vals, std::uint32_t n) {
  Value* cells = arena.create_array<Value>(n);
  for (std::uint32_t i = 0; i < n; ++i) {
    cells[i] = Value::number(vals[i]);
  }
  Value* tail = arena.create_array<Value>(1);
  tail[0] = Value::blank(BlankGridProjection::kReferenceGridZero);
  Shaped s;
  s.tail_array = make_tail_array(arena, kGridRows, 1U, n, TailAxis::kRows, cells, tail, true);
  return s;
}

Shaped dense(Arena& arena, std::uint32_t rows, std::uint32_t cols, const double* vals) {
  Value* buf = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buf, kMaxDerivedArrayCells);
  for (std::uint32_t i = 0; i < rows * cols; ++i) {
    buf[i] = Value::number(vals[i]);
  }
  return Shaped{Value::array(arr), nullptr};
}

Shaped scalar(Value v) {
  return Shaped{v, nullptr};
}

const TailArray& as_tail(const Shaped& s) {
  EXPECT_NE(s.tail_array, nullptr);
  return *s.tail_array;
}

TEST(BroadcastTail, TailPlusScalar) {
  Arena arena;
  const double v[] = {10, 20, 30, 40, 50};
  const Shaped r = broadcast_binop(parser::BinOp::Add, column_head(arena, v, 5), scalar(Value::number(0)), arena);
  const TailArray& t = as_tail(r);
  EXPECT_EQ(t.rows, kGridRows);
  EXPECT_EQ(t.cols, 1U);
  EXPECT_EQ(t.head, 5U);
  EXPECT_FALSE(t.from_reference);
  for (std::uint32_t i = 0; i < 5; ++i) {
    EXPECT_EQ(tail_array_at(t, i, 0).as_number(), v[i]);
  }
  EXPECT_EQ(tail_array_at(t, 5, 0).as_number(), 0);
  EXPECT_EQ(tail_array_at(t, kGridRows - 1, 0).as_number(), 0);
}

TEST(BroadcastTail, ScalarOnLeftKeepsOperandOrder) {
  Arena arena;
  const double v[] = {1, 2};
  const Shaped r = broadcast_binop(parser::BinOp::Sub, scalar(Value::number(10)), column_head(arena, v, 2), arena);
  const TailArray& t = as_tail(r);
  EXPECT_EQ(tail_array_at(t, 1, 0).as_number(), 8);
  EXPECT_EQ(tail_array_at(t, 2, 0).as_number(), 10);
}

TEST(BroadcastTail, TailPlusTailDifferentHeads) {
  Arena arena;
  const double a[] = {1, 2, 3};
  const double b[] = {10, 20};
  const Shaped r = broadcast_binop(parser::BinOp::Add, column_head(arena, a, 3), column_head(arena, b, 2), arena);
  const TailArray& t = as_tail(r);
  EXPECT_EQ(t.head, 3U);
  EXPECT_EQ(tail_array_at(t, 0, 0).as_number(), 11);
  EXPECT_EQ(tail_array_at(t, 1, 0).as_number(), 22);
  EXPECT_EQ(tail_array_at(t, 2, 0).as_number(), 3);
  EXPECT_EQ(tail_array_at(t, 3, 0).as_number(), 0);
}

TEST(BroadcastTail, TailPlusDenseColumnPadsNa) {
  Arena arena;
  const double a[] = {10, 20, 30, 40, 50};
  const double d[] = {1, 2, 3};
  const Shaped r = broadcast_binop(parser::BinOp::Add, column_head(arena, a, 5), dense(arena, 3, 1, d), arena);
  const TailArray& t = as_tail(r);
  EXPECT_EQ(t.rows, kGridRows);
  EXPECT_EQ(t.head, 5U);
  EXPECT_EQ(tail_array_at(t, 0, 0).as_number(), 11);
  EXPECT_EQ(tail_array_at(t, 2, 0).as_number(), 33);
  EXPECT_TRUE(tail_array_at(t, 3, 0).is_error());
  EXPECT_EQ(tail_array_at(t, 3, 0).as_error(), ErrorCode::NA);
  EXPECT_EQ(tail_array_at(t, 4, 0).as_error(), ErrorCode::NA);
  EXPECT_EQ(tail_array_at(t, 5, 0).as_error(), ErrorCode::NA);
  EXPECT_EQ(tail_array_at(t, kGridRows - 1, 0).as_error(), ErrorCode::NA);
}

TEST(BroadcastTail, TailPlusDenseRowStretches) {
  Arena arena;
  const double a[] = {10, 20, 30, 40, 50};
  const double d[] = {1, 2};
  const Shaped r = broadcast_binop(parser::BinOp::Add, column_head(arena, a, 5), dense(arena, 1, 2, d), arena);
  const TailArray& t = as_tail(r);
  EXPECT_EQ(t.rows, kGridRows);
  EXPECT_EQ(t.cols, 2U);
  EXPECT_EQ(t.head, 5U);
  EXPECT_EQ(tail_array_at(t, 0, 0).as_number(), 11);
  EXPECT_EQ(tail_array_at(t, 4, 1).as_number(), 52);
  EXPECT_EQ(tail_array_at(t, 5, 0).as_number(), 1);
  EXPECT_EQ(tail_array_at(t, 5, 1).as_number(), 2);
  EXPECT_EQ(tail_array_at(t, kGridRows - 1, 1).as_number(), 2);
}

TEST(BroadcastTail, ColsAxisPlusScalar) {
  Arena arena;
  Value* cells = arena.create_array<Value>(2);
  cells[0] = Value::number(7);
  cells[1] = Value::number(8);
  Value* tail = arena.create_array<Value>(1);
  tail[0] = Value::blank(BlankGridProjection::kReferenceGridZero);
  Shaped row;
  row.tail_array = make_tail_array(arena, 1U, kGridCols, 2U, TailAxis::kCols, cells, tail, true);
  const Shaped r = broadcast_binop(parser::BinOp::Add, row, scalar(Value::number(0)), arena);
  const TailArray& t = as_tail(r);
  EXPECT_EQ(t.rows, 1U);
  EXPECT_EQ(t.cols, kGridCols);
  EXPECT_EQ(t.axis, TailAxis::kCols);
  EXPECT_EQ(t.head, 2U);
  EXPECT_EQ(tail_array_at(t, 0, 1).as_number(), 8);
  EXPECT_EQ(tail_array_at(t, 0, 2).as_number(), 0);
  EXPECT_EQ(tail_array_at(t, 0, kGridCols - 1).as_number(), 0);
}

TEST(BroadcastTail, MixedAxesDensifiesWithinCap) {
  Arena arena;
  // 4x1 column (head 1) against 1x3 row (head 1): dense 4x3 outer sum.
  Value* ccells = arena.create_array<Value>(1);
  ccells[0] = Value::number(1);
  Value* ctail = arena.create_array<Value>(1);
  ctail[0] = Value::number(0);
  Shaped col;
  col.tail_array = make_tail_array(arena, 4U, 1U, 1U, TailAxis::kRows, ccells, ctail, false);
  Value* rcells = arena.create_array<Value>(1);
  rcells[0] = Value::number(10);
  Value* rtail = arena.create_array<Value>(1);
  rtail[0] = Value::number(0);
  Shaped row;
  row.tail_array = make_tail_array(arena, 1U, 3U, 1U, TailAxis::kCols, rcells, rtail, false);
  const Shaped r = broadcast_binop(parser::BinOp::Add, col, row, arena);
  ASSERT_EQ(r.tail_array, nullptr);
  ASSERT_TRUE(r.value.is_array());
  const ArrayValue* a = r.value.as_array();
  EXPECT_EQ(a->rows, 4U);
  EXPECT_EQ(a->cols, 3U);
  EXPECT_EQ(a->cells[0].as_number(), 11);
  EXPECT_EQ(a->cells[1].as_number(), 1);
  EXPECT_EQ(a->cells[3].as_number(), 10);
  EXPECT_EQ(a->cells[11].as_number(), 0);
}

TEST(BroadcastTail, MixedAxesOverCapIsNum) {
  Arena arena;
  const double v[] = {1};
  Shaped col = column_head(arena, v, 1);
  Value* rcells = arena.create_array<Value>(1);
  rcells[0] = Value::number(1);
  Value* rtail = arena.create_array<Value>(1);
  rtail[0] = Value::number(0);
  Shaped row;
  row.tail_array = make_tail_array(arena, 1U, kGridCols, 1U, TailAxis::kCols, rcells, rtail, true);
  const Shaped r = broadcast_binop(parser::BinOp::Add, col, row, arena);
  ASSERT_EQ(r.tail_array, nullptr);
  ASSERT_TRUE(r.value.is_error());
  EXPECT_EQ(r.value.as_error(), ErrorCode::Num);
}

TEST(BroadcastTail, UnaryMinus) {
  Arena arena;
  const double v[] = {10, 20};
  const Shaped r = broadcast_unary(parser::UnaryOp::Minus, column_head(arena, v, 2), arena);
  const TailArray& t = as_tail(r);
  EXPECT_FALSE(t.from_reference);
  EXPECT_EQ(t.rows, kGridRows);
  EXPECT_EQ(tail_array_at(t, 1, 0).as_number(), -20);
  EXPECT_EQ(tail_array_at(t, 2, 0).as_number(), 0);
}

TEST(BroadcastTail, UnaryOnPlainValueDelegates) {
  Arena arena;
  const Shaped r = broadcast_unary(parser::UnaryOp::Minus, scalar(Value::number(3)), arena);
  ASSERT_EQ(r.tail_array, nullptr);
  EXPECT_EQ(r.value.as_number(), -3);
}

TEST(BroadcastTail, ConcatTailIsText) {
  Arena arena;
  const double v[] = {1};
  const Shaped r = broadcast_binop(parser::BinOp::Concat, column_head(arena, v, 1), scalar(Value::text("x")), arena);
  const TailArray& t = as_tail(r);
  EXPECT_EQ(tail_array_at(t, 0, 0).as_text(), "1x");
  ASSERT_TRUE(tail_array_at(t, 1, 0).is_text());
  EXPECT_EQ(tail_array_at(t, 1, 0).as_text(), "x");
  EXPECT_EQ(tail_array_at(t, kGridRows - 1, 0).as_text(), "x");
}

TEST(BroadcastTail, ComparisonWithEmptyTextTailIsTrue) {
  Arena arena;
  const double v[] = {1};
  const Shaped r = broadcast_binop(parser::BinOp::Eq, column_head(arena, v, 1), scalar(Value::text("")), arena);
  const TailArray& t = as_tail(r);
  EXPECT_FALSE(tail_array_at(t, 0, 0).as_boolean());
  ASSERT_TRUE(tail_array_at(t, 1, 0).is_boolean());
  EXPECT_TRUE(tail_array_at(t, 1, 0).as_boolean());
}

TEST(BroadcastTail, HeadErrorStaysPerCell) {
  Arena arena;
  Value* cells = arena.create_array<Value>(2);
  cells[0] = Value::error(ErrorCode::Div0);
  cells[1] = Value::number(5);
  Value* tail = arena.create_array<Value>(1);
  tail[0] = Value::blank(BlankGridProjection::kReferenceGridZero);
  Shaped col;
  col.tail_array = make_tail_array(arena, kGridRows, 1U, 2U, TailAxis::kRows, cells, tail, true);
  const Shaped r = broadcast_binop(parser::BinOp::Add, col, scalar(Value::number(1)), arena);
  const TailArray& t = as_tail(r);
  EXPECT_EQ(tail_array_at(t, 0, 0).as_error(), ErrorCode::Div0);
  EXPECT_EQ(tail_array_at(t, 1, 0).as_number(), 6);
  EXPECT_EQ(tail_array_at(t, 2, 0).as_number(), 1);
}

TEST(BroadcastTail, TailErrorPropagatesToTail) {
  Arena arena;
  const double v[] = {4};
  const Shaped r = broadcast_binop(parser::BinOp::Div, scalar(Value::number(1)), column_head(arena, v, 1), arena);
  const TailArray& t = as_tail(r);
  EXPECT_EQ(tail_array_at(t, 0, 0).as_number(), 0.25);
  EXPECT_EQ(tail_array_at(t, 1, 0).as_error(), ErrorCode::Div0);
}

TEST(BroadcastTail, FullHeadYieldsDenseResult) {
  Arena arena;
  Value* cells = arena.create_array<Value>(3);
  for (int i = 0; i < 3; ++i) {
    cells[i] = Value::number(i + 1);
  }
  Value* tail = arena.create_array<Value>(1);
  tail[0] = Value::number(0);
  Shaped col;
  col.tail_array = make_tail_array(arena, 3U, 1U, 3U, TailAxis::kRows, cells, tail, true);
  const Shaped r = broadcast_binop(parser::BinOp::Mul, col, scalar(Value::number(2)), arena);
  ASSERT_EQ(r.tail_array, nullptr);
  ASSERT_TRUE(r.value.is_array());
  EXPECT_EQ(r.value.as_array()->cells[2].as_number(), 6);
}

TEST(BroadcastTail, ScalarPlusScalarStaysScalar) {
  Arena arena;
  const Shaped r = broadcast_binop(parser::BinOp::Add, scalar(Value::number(1)), scalar(Value::number(2)), arena);
  ASSERT_EQ(r.tail_array, nullptr);
  ASSERT_FALSE(r.value.is_array());
  EXPECT_EQ(r.value.as_number(), 3);
}

TEST(BroadcastTail, ScalarErrorPropagatesLeftFirst) {
  Arena arena;
  const Shaped l = broadcast_binop(parser::BinOp::Add, scalar(Value::error(ErrorCode::Div0)),
                                   scalar(Value::error(ErrorCode::Ref)), arena);
  ASSERT_FALSE(l.value.is_array());
  EXPECT_EQ(l.value.as_error(), ErrorCode::Div0);
  const double d[] = {1, 2};
  const Shaped r =
      broadcast_binop(parser::BinOp::Add, dense(arena, 2, 1, d), scalar(Value::error(ErrorCode::Ref)), arena);
  ASSERT_FALSE(r.value.is_array());
  EXPECT_EQ(r.value.as_error(), ErrorCode::Ref);
}

TEST(BroadcastTail, DensePlusScalarUnchanged) {
  Arena arena;
  const double d[] = {1, 2};
  const Shaped r = broadcast_binop(parser::BinOp::Add, dense(arena, 2, 1, d), scalar(Value::number(10)), arena);
  ASSERT_EQ(r.tail_array, nullptr);
  ASSERT_TRUE(r.value.is_array());
  EXPECT_EQ(r.value.as_array()->rows, 2U);
  EXPECT_EQ(r.value.as_array()->cells[1].as_number(), 12);
}

TEST(BroadcastTail, UnaryScalarErrorAndDense) {
  Arena arena;
  const Shaped e = broadcast_unary(parser::UnaryOp::Minus, scalar(Value::error(ErrorCode::NA)), arena);
  ASSERT_FALSE(e.value.is_array());
  EXPECT_EQ(e.value.as_error(), ErrorCode::NA);
  const double d[] = {1, 2};
  const Shaped r = broadcast_unary(parser::UnaryOp::Minus, dense(arena, 2, 1, d), arena);
  ASSERT_TRUE(r.value.is_array());
  EXPECT_EQ(r.value.as_array()->cells[1].as_number(), -2);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
