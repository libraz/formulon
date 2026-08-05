//
// Unit tests for `eval::allocate_array_value`, the single allocation seam
// every `ArrayValue` the evaluator produces goes through.
//
// The seam exists because a bare `rows * cols` wraps on a 32-bit `size_t`
// (the wasm32 main build): the buffer comes back short while the caller's
// fill loop still walks `rows * cols` slots. These tests pin the rejection
// contract on both a 64-bit and a 32-bit `size_t`, since the wrap itself is
// only reachable on the latter.

#include "eval/array_alloc.h"

#include <cstddef>
#include <cstdint>

#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon::eval {
namespace {

TEST(ArrayAlloc, AllocatesRequestedShapeAndBuffer) {
  Arena arena;
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(3U, 4U, arena, buffer);
  ASSERT_NE(arr, nullptr);
  ASSERT_NE(buffer, nullptr);
  EXPECT_EQ(arr->rows, 3U);
  EXPECT_EQ(arr->cols, 4U);
  EXPECT_EQ(arr->cells, buffer);

  // The whole buffer is writable: a short allocation would trip the
  // sanitizer builds here rather than silently corrupting the arena.
  for (std::size_t i = 0; i < 12U; ++i) {
    buffer[i] = Value::number(static_cast<double>(i));
  }
  EXPECT_DOUBLE_EQ(arr->cells[11].as_number(), 11.0);
}

TEST(ArrayAlloc, RejectsDegenerateAxesAndClearsOutBuffer) {
  Arena arena;
  Value* buffer = reinterpret_cast<Value*>(0x1);
  EXPECT_EQ(allocate_array_value(0U, 4U, arena, buffer), nullptr);
  EXPECT_EQ(buffer, nullptr) << "a rejected request must not leave a stale buffer";

  buffer = reinterpret_cast<Value*>(0x1);
  EXPECT_EQ(allocate_array_value(4U, 0U, arena, buffer), nullptr);
  EXPECT_EQ(buffer, nullptr);
}

TEST(ArrayAlloc, RejectsAxisBeyondTheExcelGrid) {
  Arena arena;
  Value* buffer = nullptr;
  EXPECT_EQ(allocate_array_value(Sheet::kMaxRows + 1U, 1U, arena, buffer), nullptr);
  EXPECT_EQ(allocate_array_value(1U, Sheet::kMaxCols + 1U, arena, buffer), nullptr);
}

TEST(ArrayAlloc, RejectsCellCountPastTheCeiling) {
  Arena arena;
  Value* buffer = nullptr;
  // One cell over the caller-supplied ceiling is rejected; exactly at the
  // ceiling is allowed.
  EXPECT_EQ(allocate_array_value(3U, 4U, arena, buffer, 11U), nullptr);
  EXPECT_EQ(buffer, nullptr);
  EXPECT_NE(allocate_array_value(3U, 4U, arena, buffer, 12U), nullptr);
}

TEST(ArrayAlloc, RejectsFullGridRequest) {
  // The full Excel grid is 1,048,576 x 16,384 — about 1.7e10 cells, and
  // past 2^32, which is where a wasm32 `rows * cols` wraps to a short
  // buffer. Both axes are individually in bounds, so only the product
  // check can reject it. The cell ceiling catches it first on a 64-bit
  // `size_t`; the wrap arithmetic itself is covered by
  // `checked_mul_size_t`'s own tests.
  Arena arena;
  Value* buffer = nullptr;
  EXPECT_EQ(allocate_array_value(Sheet::kMaxRows, Sheet::kMaxCols, arena, buffer), nullptr);
  EXPECT_EQ(buffer, nullptr);
  EXPECT_EQ(allocate_array_value(Sheet::kMaxRows, Sheet::kMaxCols, arena, buffer, kMaxDerivedArrayCells), nullptr);
  EXPECT_EQ(buffer, nullptr);
}

TEST(ArrayAlloc, ReportsArenaExhaustion) {
  // A budgeted arena that cannot satisfy the buffer returns nullptr rather
  // than handing back a partially-sized allocation.
  Arena arena(/*initial_chunk_bytes=*/64, /*max_bytes=*/64);
  Value* buffer = nullptr;
  EXPECT_EQ(allocate_array_value(1024U, 1024U, arena, buffer), nullptr);
  EXPECT_EQ(buffer, nullptr);
  EXPECT_TRUE(arena.exhausted());
}

}  // namespace
}  // namespace formulon::eval
