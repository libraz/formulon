// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `Sheet::set_cell_cached_value`'s Text deep-copy contract.
//
// `Sheet::set_cell_cached_value` is the recalc engine's post-evaluation write
// path: it stores a freshly computed `Value` into a formula cell's
// `cached_value` slot. Because the recalc engine resets its per-evaluation
// `Arena` between cells, any Text payload it produces lives in arena-owned
// bytes that vanish on the next reset. The Sheet implementation therefore
// deep-copies Text payloads into `Cell::cached_text_owned` so the stored
// `cached_value.as_text()` view survives across recalc passes.
//
// These tests pin down that contract by:
//   * Mutating the caller's source buffer after the write and confirming the
//     stored Text is unaffected (the deep copy decouples lifetimes).
//   * Confirming the stored buffer's address differs from the source buffer.
//   * Exercising overwrite paths that toggle between Text and non-Text.
//   * Exercising the empty-string boundary.

#include <cstdint>
#include <string>
#include <string_view>

#include "cell.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"

namespace formulon {
namespace {

TEST(SheetCachedTextTest, TextValueDeepCopiedFromCallerStorage) {
  Sheet s("Sheet1");
  std::string scratch = "hello world";
  const char* scratch_data_before = scratch.data();
  s.set_cell_cached_value(0U, 0U, Value::text(scratch));

  // Mutate the caller's storage. A view-only assignment would now reflect
  // the corrupted bytes; a deep copy is independent.
  scratch = "DIFFERENT BYTES THAT WOULD CORRUPT";

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_text());
  EXPECT_EQ(cell->cached_value.as_text(), "hello world");

  // Storage identity check: the stored view must not point into the
  // caller's `scratch` buffer. (We compare against the original data()
  // because the post-mutation `scratch` may have rebound to a different
  // SSO/heap slot independent of the cell.)
  EXPECT_NE(cell->cached_value.as_text().data(), scratch_data_before);
  EXPECT_NE(cell->cached_value.as_text().data(), scratch.data());
}

TEST(SheetCachedTextTest, OverwriteTextWithLongerString) {
  Sheet s("Sheet1");
  s.set_cell_cached_value(0U, 0U, Value::text("a"));

  const std::string long_text(200U, 'x');
  ASSERT_GT(long_text.size(), 23U) << "long_text must exceed typical SSO capacity to exercise heap path";
  s.set_cell_cached_value(0U, 0U, Value::text(long_text));

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_text());
  EXPECT_EQ(cell->cached_value.as_text(), std::string_view(long_text));
  EXPECT_EQ(cell->cached_value.as_text().size(), 200U);
}

TEST(SheetCachedTextTest, OverwriteTextWithNonText) {
  Sheet s("Sheet1");
  s.set_cell_cached_value(0U, 0U, Value::text("placeholder"));
  s.set_cell_cached_value(0U, 0U, Value::number(42.0));

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), 42.0);
  // `cached_text_owned` is documented as harmlessly retained on the
  // non-Text path; the new `cached_value` does not reference it. We don't
  // assert the buffer is empty (that's an internal detail), only that the
  // stored value reads back correctly.
}

TEST(SheetCachedTextTest, OverwriteNonTextWithText) {
  Sheet s("Sheet1");
  s.set_cell_cached_value(0U, 0U, Value::number(7.0));
  s.set_cell_cached_value(0U, 0U, Value::text("after-number"));

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_text());
  EXPECT_EQ(cell->cached_value.as_text(), "after-number");
}

TEST(SheetCachedTextTest, EmptyTextRoundTrip) {
  Sheet s("Sheet1");
  s.set_cell_cached_value(0U, 0U, Value::text(""));

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_text());
  EXPECT_TRUE(cell->cached_value.as_text().empty());
}

TEST(SheetCachedTextTest, TextSurvivesRowVectorGrowth) {
  // The Sheet's row store is a hash map of dense per-row vectors. Writing
  // to a higher column in the same row can resize that vector and move
  // every existing Cell. The `cached_value`'s `string_view` must remain
  // valid across such moves — a bare `std::string` (SSO) inside Cell
  // would relocate its inline bytes and dangle the view, but the
  // implementation parks the bytes on the heap via
  // `unique_ptr<std::string>` to keep the address stable.
  Sheet s("Sheet1");
  s.set_cell_cached_value(0U, 0U, Value::text("anchor"));

  // Force the row vector to grow several times so the previous Cell at
  // (0,0) gets relocated. Different columns trigger increasingly large
  // resizes; col 100 forces the vector to span 101 slots.
  s.set_cell_cached_value(0U, 5U, Value::number(1.0));
  s.set_cell_cached_value(0U, 30U, Value::number(2.0));
  s.set_cell_cached_value(0U, 100U, Value::number(3.0));

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_text());
  EXPECT_EQ(cell->cached_value.as_text(), "anchor");
}

TEST(SheetCachedTextTest, TextSurvivesRowMapRehash) {
  // Adding many distinct rows forces the underlying unordered_map to
  // rehash, which move-constructs every row's vector. Each Cell inside
  // those vectors is move-constructed in turn. The Text payload's
  // `string_view` must still resolve to the original bytes.
  Sheet s("Sheet1");
  s.set_cell_cached_value(0U, 0U, Value::text("first-row"));
  s.set_cell_cached_value(1U, 0U, Value::text("second-row"));

  // Touch a large number of distinct rows to provoke at least one rehash
  // of the row map.
  for (std::uint32_t r = 2U; r < 200U; ++r) {
    s.set_cell_cached_value(r, 0U, Value::number(static_cast<double>(r)));
  }

  const Cell* a1 = s.cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "first-row");

  const Cell* a2 = s.cell_at(1U, 0U);
  ASSERT_NE(a2, nullptr);
  ASSERT_TRUE(a2->cached_value.is_text());
  EXPECT_EQ(a2->cached_value.as_text(), "second-row");
}

TEST(SheetCachedTextTest, RewriteSameCellWithItsOwnCachedTextDoesNotCorrupt) {
  // Defensive: callers occasionally cycle a cached Text value back through
  // the same cell (e.g. iterative-mode commit reading the previous
  // iteration's value back). The implementation must not corrupt the
  // source view by overwriting `cached_text_owned` before consuming it.
  Sheet s("Sheet1");
  s.set_cell_cached_value(0U, 0U, Value::text("seed-text"));

  const Cell* read_back = s.cell_at(0U, 0U);
  ASSERT_NE(read_back, nullptr);
  Value v_from_cell = read_back->cached_value;
  s.set_cell_cached_value(0U, 0U, v_from_cell);

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_text());
  EXPECT_EQ(cell->cached_value.as_text(), "seed-text");
}

}  // namespace
}  // namespace formulon
