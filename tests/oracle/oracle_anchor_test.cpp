// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `anchor_or_self`, the spill-anchor projection helper used
// by the oracle comparator. Excel 365 reports only the top-left cell of a
// dynamic-array spill region; the helper unwraps Formulon's full Array
// `Value` to that scalar so the comparator's kind dispatch matches.

#include "tests/oracle/oracle_anchor.h"

#include <cstddef>
#include <cstdint>

#include "gtest/gtest.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace tests {
namespace oracle {
namespace {

// Builds an arena-backed `ArrayValue` of shape (rows x cols) populated by
// copying `cells_in` (length must equal rows*cols). The pointer is non-owning
// and stays valid for the lifetime of `arena`.
const ArrayValue* MakeArray(Arena* arena, std::uint32_t rows, std::uint32_t cols, const Value* cells_in) {
  const std::size_t n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  Value* cells = arena->create_array<Value>(n);
  for (std::size_t i = 0; i < n; ++i) {
    cells[i] = cells_in[i];
  }
  ArrayValue* arr = arena->create<ArrayValue>();
  arr->rows = rows;
  arr->cols = cols;
  arr->cells = cells;
  return arr;
}

TEST(OracleAnchor, NonArrayPassthrough) {
  const Value v = Value::number(42.0);
  const Value& got = anchor_or_self(v);
  EXPECT_TRUE(got.is_number());
  EXPECT_EQ(42.0, got.as_number());
  // Identity: pass-through must not copy.
  EXPECT_EQ(&v, &got);
}

TEST(OracleAnchor, BlankPassthrough) {
  const Value v = Value::blank();
  const Value& got = anchor_or_self(v);
  EXPECT_TRUE(got.is_blank());
  EXPECT_EQ(&v, &got);
}

TEST(OracleAnchor, BoolPassthrough) {
  const Value v = Value::boolean(true);
  const Value& got = anchor_or_self(v);
  EXPECT_TRUE(got.is_boolean());
  EXPECT_TRUE(got.as_boolean());
  EXPECT_EQ(&v, &got);
}

TEST(OracleAnchor, ErrorPassthrough) {
  const Value v = Value::error(ErrorCode::Div0);
  const Value& got = anchor_or_self(v);
  EXPECT_TRUE(got.is_error());
  EXPECT_EQ(ErrorCode::Div0, got.as_error());
  EXPECT_EQ(&v, &got);
}

TEST(OracleAnchor, ArrayUnwrapsToAnchor) {
  Arena arena;
  const Value cells[] = {Value::number(10.0), Value::number(20.0), Value::number(30.0),
                         Value::number(40.0), Value::number(50.0), Value::number(60.0)};
  const ArrayValue* arr = MakeArray(&arena, 2, 3, cells);
  const Value v = Value::array(arr);

  const Value& got = anchor_or_self(v);
  EXPECT_TRUE(got.is_number());
  EXPECT_EQ(10.0, got.as_number());
  // The returned reference must alias the arena-owned cell, not a copy.
  EXPECT_EQ(&arr->cells[0], &got);
}

TEST(OracleAnchor, Array1x1UnwrapsToScalar) {
  Arena arena;
  const Value cells[] = {Value::text("hello")};
  const ArrayValue* arr = MakeArray(&arena, 1, 1, cells);
  const Value v = Value::array(arr);

  const Value& got = anchor_or_self(v);
  EXPECT_TRUE(got.is_text());
  EXPECT_EQ("hello", got.as_text());
  EXPECT_EQ(&arr->cells[0], &got);
}

TEST(OracleAnchor, EmptyArrayPassthrough) {
  Arena arena;
  ArrayValue* arr = arena.create<ArrayValue>();
  arr->rows = 0;
  arr->cols = 0;
  arr->cells = nullptr;
  const Value v = Value::array(arr);

  const Value& got = anchor_or_self(v);
  // Empty array must round-trip unchanged so the comparator's kind dispatch
  // can surface a normal mismatch error.
  EXPECT_TRUE(got.is_array());
  EXPECT_EQ(&v, &got);
}

}  // namespace
}  // namespace oracle
}  // namespace tests
}  // namespace formulon
