// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the `Array` variant of `Value`. Scope is limited to the
// data-structure scaffold: factory, accessors, debug formatting, equality,
// and the existing `Array -> #VALUE!` coercion behaviour now that the
// variant is reachable. Producers (operator broadcasting, SUMPRODUCT) are
// out of scope.

#include <cstddef>
#include <cstdint>
#include <string>

#include "eval/coerce.h"
#include "gtest/gtest.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace {

// Helper: builds an `ArrayValue` of shape (rows x cols) populated by copying
// `cells_in` (length must equal rows*cols) into the arena. Returns the
// arena-backed ArrayValue pointer.
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

TEST(ValueArray, FactoryAndAccessors) {
  Arena arena;
  const Value cells[] = {Value::number(1.0), Value::number(2.0), Value::number(3.0),
                         Value::number(4.0), Value::number(5.0), Value::number(6.0)};
  const ArrayValue* arr = MakeArray(&arena, 2, 3, cells);

  Value v = Value::array(arr);
  EXPECT_EQ(ValueKind::Array, v.kind());
  EXPECT_TRUE(v.is_array());
  EXPECT_FALSE(v.is_blank());
  EXPECT_FALSE(v.is_number());
  EXPECT_FALSE(v.is_boolean());
  EXPECT_FALSE(v.is_error());
  EXPECT_FALSE(v.is_text());
  EXPECT_FALSE(v.is_ref());
  EXPECT_FALSE(v.is_lambda());

  EXPECT_EQ(arr, v.as_array());
  EXPECT_EQ(2u, v.as_array_rows());
  EXPECT_EQ(3u, v.as_array_cols());
  EXPECT_EQ(arr->cells, v.as_array_cells());

  // Round-trip individual cells.
  for (std::size_t i = 0; i < 6; ++i) {
    EXPECT_EQ(cells[i], v.as_array_cells()[i]) << "i=" << i;
  }
}

TEST(ValueArray, DebugStringShowsShape) {
  Arena arena;
  const Value cells[] = {Value::number(1.0), Value::number(2.0), Value::number(3.0),
                         Value::number(4.0), Value::number(5.0), Value::number(6.0)};
  const ArrayValue* arr = MakeArray(&arena, 2, 3, cells);
  EXPECT_EQ("Array(2x3)", Value::array(arr).debug_to_string());

  const Value one_cell[] = {Value::number(7.0)};
  const ArrayValue* arr_1x1 = MakeArray(&arena, 1, 1, one_cell);
  EXPECT_EQ("Array(1x1)", Value::array(arr_1x1).debug_to_string());
}

TEST(ValueArray, EqualityRequiresShapeAndCellMatch) {
  Arena arena;
  const Value abcd[] = {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0)};

  const ArrayValue* lhs_2x2 = MakeArray(&arena, 2, 2, abcd);
  const ArrayValue* rhs_2x2 = MakeArray(&arena, 2, 2, abcd);
  EXPECT_EQ(Value::array(lhs_2x2), Value::array(rhs_2x2));

  const ArrayValue* same_cells_1x4 = MakeArray(&arena, 1, 4, abcd);
  EXPECT_NE(Value::array(lhs_2x2), Value::array(same_cells_1x4));

  const Value abcd_diff[] = {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(99.0)};
  const ArrayValue* diff_2x2 = MakeArray(&arena, 2, 2, abcd_diff);
  EXPECT_NE(Value::array(lhs_2x2), Value::array(diff_2x2));
}

TEST(ValueArray, EqualityVsScalarIsFalse) {
  Arena arena;
  const Value cells[] = {Value::number(1.0)};
  const ArrayValue* arr = MakeArray(&arena, 1, 1, cells);
  Value v = Value::array(arr);
  EXPECT_NE(v, Value::number(1.0));
  EXPECT_NE(Value::number(1.0), v);
  EXPECT_NE(v, Value::blank());
  EXPECT_NE(v, Value::text(""));
  EXPECT_NE(v, Value::error(ErrorCode::Value));
  EXPECT_NE(v, Value::boolean(true));
}

TEST(ValueArray, IsArrayQuery) {
  Arena arena;
  const Value cells[] = {Value::number(1.0)};
  const ArrayValue* arr = MakeArray(&arena, 1, 1, cells);
  Value v = Value::array(arr);
  EXPECT_TRUE(v.is_array());

  EXPECT_FALSE(Value::number(1.0).is_array());
  EXPECT_FALSE(Value::text("x").is_array());
  EXPECT_FALSE(Value::blank().is_array());
  EXPECT_FALSE(Value::boolean(true).is_array());
  EXPECT_FALSE(Value::error(ErrorCode::Value).is_array());
}

TEST(ValueArray, CoerceToNumberReturnsValue) {
  Arena arena;
  const Value cells[] = {Value::number(1.0)};
  const ArrayValue* arr = MakeArray(&arena, 1, 1, cells);
  auto r = eval::coerce_to_number(Value::array(arr));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(ErrorCode::Value, r.error());
}

TEST(ValueArray, CoerceToTextReturnsValue) {
  Arena arena;
  const Value cells[] = {Value::number(1.0)};
  const ArrayValue* arr = MakeArray(&arena, 1, 1, cells);
  auto r = eval::coerce_to_text(Value::array(arr));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(ErrorCode::Value, r.error());
}

TEST(ValueArray, CoerceToBoolReturnsValue) {
  Arena arena;
  const Value cells[] = {Value::number(1.0)};
  const ArrayValue* arr = MakeArray(&arena, 1, 1, cells);
  auto r = eval::coerce_to_bool(Value::array(arr));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(ErrorCode::Value, r.error());
}

}  // namespace
}  // namespace formulon
