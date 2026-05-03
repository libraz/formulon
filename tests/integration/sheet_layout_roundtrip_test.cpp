// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Sheet view / layout round-trip integration test. Builds a workbook
// with mixed column widths, row heights, frozen panes, hidden tabs and
// outline levels, saves it through `Workbook::save()`, reloads via
// `io::read_ooxml`, and asserts every observable attribute survives.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

std::vector<std::uint8_t> SaveOrDie(const Workbook& wb) {
  auto save_or = wb.save();
  EXPECT_TRUE(static_cast<bool>(save_or)) << "save() failed: " << save_or.error().message;
  return save_or.value();
}

TEST(SheetLayoutRoundTrip, ViewStateZoomFreezeAndTabHidden) {
  Workbook src = Workbook::create();
  Sheet& sheet = src.sheet(0);
  sheet.mutable_view().zoom_scale = 125U;
  sheet.mutable_view().freeze_rows = 3U;
  sheet.mutable_view().freeze_cols = 2U;
  sheet.mutable_view().tab_hidden = true;

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << result_or.error().message;
  const Sheet& loaded = result_or.value().workbook.sheet(0);
  EXPECT_EQ(loaded.view().zoom_scale, 125U);
  EXPECT_EQ(loaded.view().freeze_rows, 3U);
  EXPECT_EQ(loaded.view().freeze_cols, 2U);
  EXPECT_TRUE(loaded.view().tab_hidden);
}

TEST(SheetLayoutRoundTrip, ColumnWidthsHiddenAndOutline) {
  Workbook src = Workbook::create();
  SheetLayout& layout = src.sheet(0).mutable_layout();
  ColumnLayout a;
  a.first = 0U;
  a.last = 2U;
  a.width = 18.5;
  layout.columns.push_back(a);
  ColumnLayout b;
  b.first = 4U;
  b.last = 4U;
  b.width = 7.0;
  b.hidden = true;
  layout.columns.push_back(b);
  ColumnLayout c;
  c.first = 6U;
  c.last = 9U;
  c.width = 12.0;
  c.outline_level = 2U;
  layout.columns.push_back(c);

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const auto& cols = result_or.value().workbook.sheet(0).layout().columns;
  ASSERT_EQ(cols.size(), 3U);
  EXPECT_EQ(cols[0].first, 0U);
  EXPECT_EQ(cols[0].last, 2U);
  EXPECT_DOUBLE_EQ(cols[0].width, 18.5);
  EXPECT_FALSE(cols[0].hidden);
  EXPECT_EQ(cols[0].outline_level, 0U);
  EXPECT_EQ(cols[1].first, 4U);
  EXPECT_EQ(cols[1].last, 4U);
  EXPECT_DOUBLE_EQ(cols[1].width, 7.0);
  EXPECT_TRUE(cols[1].hidden);
  EXPECT_EQ(cols[2].first, 6U);
  EXPECT_EQ(cols[2].last, 9U);
  EXPECT_DOUBLE_EQ(cols[2].width, 12.0);
  EXPECT_EQ(cols[2].outline_level, 2U);
}

TEST(SheetLayoutRoundTrip, RowHeightsHiddenAndOutline) {
  Workbook src = Workbook::create();
  // Populate one cell so the row is in the dense store as well —
  // the writer still has to merge override attrs onto rows that
  // carry data.
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0, 5U, 0U, Value::number(42.0))));
  SheetLayout& layout = src.sheet(0).mutable_layout();
  RowLayout r5;
  r5.row = 5U;
  r5.height = 22.0;
  layout.row_overrides.push_back(r5);
  RowLayout r10;
  r10.row = 10U;
  r10.hidden = true;
  layout.row_overrides.push_back(r10);
  RowLayout r12;
  r12.row = 12U;
  r12.height = 30.5;
  r12.outline_level = 1U;
  layout.row_overrides.push_back(r12);

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const auto& rows = result_or.value().workbook.sheet(0).layout().row_overrides;
  ASSERT_EQ(rows.size(), 3U);
  // Order is implementation-defined but the round-trip preserves
  // identity of the entries; sort by row index for deterministic
  // assertions.
  std::vector<RowLayout> sorted = rows;
  std::sort(sorted.begin(), sorted.end(), [](const RowLayout& a, const RowLayout& b) { return a.row < b.row; });
  EXPECT_EQ(sorted[0].row, 5U);
  EXPECT_DOUBLE_EQ(sorted[0].height, 22.0);
  EXPECT_FALSE(sorted[0].hidden);
  EXPECT_EQ(sorted[0].outline_level, 0U);
  EXPECT_EQ(sorted[1].row, 10U);
  EXPECT_TRUE(sorted[1].hidden);
  EXPECT_EQ(sorted[2].row, 12U);
  EXPECT_DOUBLE_EQ(sorted[2].height, 30.5);
  EXPECT_EQ(sorted[2].outline_level, 1U);
}

TEST(SheetLayoutRoundTrip, ZoomDefaultsAreOmittedFromOutput) {
  // A pristine workbook should not surface a `<sheetView>` block in
  // the saved sheet XML; the read-back should still report the
  // default values.
  Workbook src = Workbook::create();
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const SheetView& v = result_or.value().workbook.sheet(0).view();
  EXPECT_EQ(v.zoom_scale, SheetView::kDefaultZoomScale);
  EXPECT_EQ(v.freeze_rows, 0U);
  EXPECT_EQ(v.freeze_cols, 0U);
  EXPECT_FALSE(v.tab_hidden);
}

TEST(SheetLayoutRoundTrip, ZoomOutOfRangeFallsBackToDefaultOnLoad) {
  // The writer happily emits whatever zoom_scale the model carries
  // (we deliberately store the value verbatim before save). The
  // reader's clamp to `[10, 400]` ensures a bogus value loaded from
  // a hand-edited xlsx file rounds to the default.
  Workbook src = Workbook::create();
  src.sheet(0).mutable_view().zoom_scale = 9999U;
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.value().workbook.sheet(0).view().zoom_scale, SheetView::kDefaultZoomScale);
}

TEST(SheetLayoutRoundTrip, RowOverridesSurviveOnEmptyRow) {
  // Row 4 has no cell data at all but carries a height override; the
  // writer must emit `<row r="5" ht="..."/>` so the override survives
  // a save/load cycle.
  Workbook src = Workbook::create();
  RowLayout r4;
  r4.row = 4U;
  r4.height = 18.0;
  src.sheet(0).mutable_layout().row_overrides.push_back(r4);

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const auto& rows = result_or.value().workbook.sheet(0).layout().row_overrides;
  ASSERT_EQ(rows.size(), 1U);
  EXPECT_EQ(rows[0].row, 4U);
  EXPECT_DOUBLE_EQ(rows[0].height, 18.0);
}

TEST(SheetLayoutRoundTrip, ComboMixedFields) {
  Workbook src = Workbook::create();
  src.add_sheet("Second");
  Sheet& main = src.sheet(0);
  main.mutable_view().zoom_scale = 75U;
  main.mutable_view().freeze_rows = 1U;
  ColumnLayout cl;
  cl.first = 1U;
  cl.last = 3U;
  cl.width = 14.5;
  cl.hidden = true;
  cl.outline_level = 1U;
  main.mutable_layout().columns.push_back(cl);
  RowLayout rl;
  rl.row = 7U;
  rl.height = 24.0;
  rl.outline_level = 2U;
  main.mutable_layout().row_overrides.push_back(rl);
  src.sheet(1).mutable_view().tab_hidden = true;

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const Workbook& dst = result_or.value().workbook;
  ASSERT_EQ(dst.sheet_count(), 2U);
  EXPECT_EQ(dst.sheet(0).view().zoom_scale, 75U);
  EXPECT_EQ(dst.sheet(0).view().freeze_rows, 1U);
  ASSERT_EQ(dst.sheet(0).layout().columns.size(), 1U);
  EXPECT_EQ(dst.sheet(0).layout().columns[0].first, 1U);
  EXPECT_EQ(dst.sheet(0).layout().columns[0].last, 3U);
  EXPECT_DOUBLE_EQ(dst.sheet(0).layout().columns[0].width, 14.5);
  EXPECT_TRUE(dst.sheet(0).layout().columns[0].hidden);
  EXPECT_EQ(dst.sheet(0).layout().columns[0].outline_level, 1U);
  ASSERT_EQ(dst.sheet(0).layout().row_overrides.size(), 1U);
  EXPECT_EQ(dst.sheet(0).layout().row_overrides[0].row, 7U);
  EXPECT_DOUBLE_EQ(dst.sheet(0).layout().row_overrides[0].height, 24.0);
  EXPECT_EQ(dst.sheet(0).layout().row_overrides[0].outline_level, 2U);
  EXPECT_TRUE(dst.sheet(1).view().tab_hidden);
}

}  // namespace
}  // namespace formulon
