// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for `eval::SpillCommitter`. Companion to
// `eval_spill_dispatch_test.cpp` — that file pins the
// `EvalContext::dispatch_array_result` wrapper and the recursive-resolver
// integration; this file pins the standalone committer that the wrapper
// now delegates to.
//
// What we cover here:
//   * Scalar passthrough (number / bool / blank / error) — every kind
//     flows through unchanged regardless of the committer's binding.
//   * Inactive committer (default-constructed or built with a null
//     sheet) — Array is returned verbatim, no Sheet mutation.
//   * Active committer commits a row-major spill at the bound anchor
//     and returns the anchor scalar (cells[0]).
//   * Collision against a pre-existing literal yields `#SPILL!` and
//     does not leave a registered region.
//   * Degenerate `0 x N` / `N x 0` arrays surface `#VALUE!`.

#include "eval/spill_committer.h"

#include <cstdint>

#include "cell.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

Value MakeArray(Arena& arena, std::uint32_t rows, std::uint32_t cols, std::initializer_list<Value> entries) {
  Value* cells = arena.create_array<Value>(static_cast<std::size_t>(rows) * cols);
  std::size_t i = 0;
  for (const Value& v : entries) {
    cells[i++] = v;
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  arr->rows = rows;
  arr->cols = cols;
  arr->cells = cells;
  return Value::array(arr);
}

TEST(SpillCommitter, ScalarsPassThroughUnchanged) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  const SpillCommitter committer(&sheet, 0U, 0U);
  EXPECT_EQ(committer.commit(Value::number(42.0)), Value::number(42.0));
  EXPECT_EQ(committer.commit(Value::boolean(true)), Value::boolean(true));
  EXPECT_EQ(committer.commit(Value::blank()), Value::blank());
  EXPECT_EQ(committer.commit(Value::error(ErrorCode::Div0)), Value::error(ErrorCode::Div0));
}

TEST(SpillCommitter, InactiveCommitterReturnsArrayVerbatim) {
  Arena arena;
  const Value array_v = MakeArray(arena, 3U, 1U, {Value::number(1.0), Value::number(2.0), Value::number(3.0)});

  // Default-constructed committer is inactive.
  EXPECT_FALSE(SpillCommitter().active());
  const Value out_default = SpillCommitter().commit(array_v);
  EXPECT_TRUE(out_default.is_array());
  EXPECT_EQ(out_default.as_array_rows(), 3U);
  EXPECT_EQ(out_default.as_array_cols(), 1U);

  // Explicit nullptr sheet is also inactive.
  EXPECT_FALSE(SpillCommitter(nullptr, 0U, 0U).active());
  const Value out_null = SpillCommitter(nullptr, 0U, 0U).commit(array_v);
  EXPECT_TRUE(out_null.is_array());
}

TEST(SpillCommitter, ActiveCommitWritesSpillAndReturnsAnchorScalar) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Arena arena;
  const Value array_v = MakeArray(arena, 3U, 1U, {Value::number(10.0), Value::number(20.0), Value::number(30.0)});

  const SpillCommitter committer(&sheet, 0U, 0U);
  EXPECT_TRUE(committer.active());
  const Value out = committer.commit(array_v);
  ASSERT_TRUE(out.is_number());
  EXPECT_DOUBLE_EQ(out.as_number(), 10.0);

  // Phantom cells of the spill region are visible via resolve_cell_value.
  EXPECT_EQ(sheet.resolve_cell_value(0U, 0U), Value::number(10.0));
  EXPECT_EQ(sheet.resolve_cell_value(1U, 0U), Value::number(20.0));
  EXPECT_EQ(sheet.resolve_cell_value(2U, 0U), Value::number(30.0));
}

TEST(SpillCommitter, CollisionWithExistingLiteralYieldsSpillError) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // Pre-populate B1 to obstruct a 1x2 spill at A1.
  sheet.set_cell_value(0U, 1U, Value::number(99.0));

  Arena arena;
  const Value array_v = MakeArray(arena, 1U, 2U, {Value::number(1.0), Value::number(2.0)});

  const SpillCommitter committer(&sheet, 0U, 0U);
  const Value out = committer.commit(array_v);
  ASSERT_TRUE(out.is_error());
  EXPECT_EQ(out.as_error(), ErrorCode::Spill);
}

TEST(SpillCommitter, DegenerateArrayShapesSurfaceValueError) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Arena arena;

  ArrayValue* zero_rows = arena.create<ArrayValue>();
  zero_rows->rows = 0U;
  zero_rows->cols = 4U;
  zero_rows->cells = nullptr;

  ArrayValue* zero_cols = arena.create<ArrayValue>();
  zero_cols->rows = 4U;
  zero_cols->cols = 0U;
  zero_cols->cells = nullptr;

  const SpillCommitter committer(&sheet, 0U, 0U);
  EXPECT_EQ(committer.commit(Value::array(zero_rows)), Value::error(ErrorCode::Value));
  EXPECT_EQ(committer.commit(Value::array(zero_cols)), Value::error(ErrorCode::Value));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
