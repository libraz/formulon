//
// Sheet view / layout round-trip integration test. Builds a workbook
// with mixed column widths, row heights, frozen panes, hidden tabs and
// outline levels, saves it through `Workbook::save()`, reloads via
// `io::read_ooxml`, and asserts every observable attribute survives.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
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

/// Extracts `xl/worksheets/sheet1.xml` from a saved package as a string.
std::string ReadSheet1Xml(const std::vector<std::uint8_t>& bytes) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  auto entry_or = zip.read_entry("xl/worksheets/sheet1.xml");
  EXPECT_TRUE(static_cast<bool>(entry_or)) << "read_entry: " << entry_or.error().message;
  const std::vector<std::uint8_t>& body = entry_or.value();
  return std::string(reinterpret_cast<const char*>(body.data()), body.size());
}

TEST(SheetLayoutRoundTrip, DimensionReflectsPopulatedBoundingBox) {
  // Cells span B2:D5; the writer must emit <dimension ref="B2:D5"/> so a
  // reader / the Name Box seeds the correct used range.
  Workbook src = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0, 1U, 1U, Value::number(1.0))));  // B2
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0, 4U, 3U, Value::number(2.0))));  // D5

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  const std::string xml = ReadSheet1Xml(bytes);
  EXPECT_NE(xml.find("<dimension ref=\"B2:D5\"/>"), std::string::npos) << xml;
}

TEST(SheetLayoutRoundTrip, DimensionOnEmptySheetIsA1) {
  Workbook src = Workbook::create();
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  const std::string xml = ReadSheet1Xml(bytes);
  EXPECT_NE(xml.find("<dimension ref=\"A1\"/>"), std::string::npos) << xml;
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

TEST(SheetLayoutRoundTrip, ReShowingSheetRemovesLegacySheetPrTabHidden) {
  Workbook src = Workbook::create();
  Sheet& sheet = src.sheet(0);
  // Some older producers record visibility in sheetPr instead of in the
  // workbook's <sheet state>.  The raw capture must not override a later
  // explicit re-show through a binding setter.
  sheet.mutable_print_settings().sheet_pr_xml = "<sheetPr tabHidden=\"1\"><tabHidden/></sheetPr>";
  sheet.mutable_view().tab_hidden = true;

  const std::vector<std::uint8_t> hidden_bytes = SaveOrDie(src);
  auto hidden_or = io::read_ooxml(SpanOf(hidden_bytes));
  ASSERT_TRUE(static_cast<bool>(hidden_or)) << hidden_or.error().message;
  ASSERT_TRUE(hidden_or.value().workbook.sheet(0).view().tab_hidden);

  hidden_or.value().workbook.sheet(0).mutable_view().tab_hidden = false;
  const std::vector<std::uint8_t> visible_bytes = SaveOrDie(hidden_or.value().workbook);
  const std::string visible_sheet_xml = ReadSheet1Xml(visible_bytes);
  EXPECT_EQ(visible_sheet_xml.find("tabHidden"), std::string::npos) << visible_sheet_xml;

  auto visible_or = io::read_ooxml(SpanOf(visible_bytes));
  ASSERT_TRUE(static_cast<bool>(visible_or)) << visible_or.error().message;
  EXPECT_FALSE(visible_or.value().workbook.sheet(0).view().tab_hidden);
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

TEST(SheetLayoutRoundTrip, LegacyNonZeroWidthIsLogicallyExplicitAcrossFormats) {
  Workbook src = Workbook::create();
  ColumnLayout legacy;
  legacy.first = 0U;
  legacy.last = 0U;
  legacy.width = 12.5;
  // Five-field aggregate/model construction historically left this raw bit
  // clear; the logical contract still treats a non-zero width as explicit.
  EXPECT_FALSE(legacy.has_width);
  src.sheet(0).mutable_layout().columns.push_back(legacy);

  const std::vector<std::uint8_t> xlsx = SaveOrDie(src);
  auto xlsx_or = io::read_ooxml(SpanOf(xlsx));
  ASSERT_TRUE(static_cast<bool>(xlsx_or)) << xlsx_or.error().message;
  const auto& xlsx_col = xlsx_or.value().workbook.sheet(0).layout().columns[0];
  EXPECT_TRUE(xlsx_col.has_width);
  EXPECT_DOUBLE_EQ(xlsx_col.width, 12.5);

  auto xlsb_or = io::xlsb::write_xlsb(src);
  ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message;
  auto xlsb_read_or = io::xlsb::read_xlsb(SpanOf(xlsb_or.value()));
  ASSERT_TRUE(static_cast<bool>(xlsb_read_or)) << xlsb_read_or.error().message;
  const auto& xlsb_col = xlsb_read_or.value().workbook.sheet(0).layout().columns[0];
  EXPECT_TRUE(xlsb_col.has_width);
  EXPECT_DOUBLE_EQ(xlsb_col.width, 12.5);
  // BrtColInfo's ixfe is mandatory; style 0 is the XLSB canonical default.
  EXPECT_TRUE(xlsb_col.has_style);
  EXPECT_EQ(xlsb_col.style_xf, 0U);
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

TEST(SheetLayoutRoundTrip, OutlineOnlyEmptyRowIsEmittedAndReloaded) {
  Workbook src = Workbook::create();
  RowLayout outline_only;
  outline_only.row = 12U;
  outline_only.outline_level = 3U;
  src.sheet(0).mutable_layout().row_overrides.push_back(outline_only);

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const auto& rows = result_or.value().workbook.sheet(0).layout().row_overrides;
  ASSERT_EQ(rows.size(), 1U);
  EXPECT_EQ(rows[0].row, 12U);
  EXPECT_DOUBLE_EQ(rows[0].height, 0.0);
  EXPECT_FALSE(rows[0].hidden);
  EXPECT_EQ(rows[0].outline_level, 3U);
}

TEST(SheetLayoutRoundTrip, AttributePresenceSurvivesWithoutCellMaterialization) {
  Workbook src = Workbook::create();
  Sheet& sheet = src.sheet(0);
  SheetLayout& layout = sheet.mutable_layout();

  ColumnLayout style_zero;
  style_zero.first = 0U;
  style_zero.last = 0U;
  style_zero.has_style = true;
  style_zero.style_xf = 0U;
  layout.columns.push_back(style_zero);

  ColumnLayout width_zero;
  width_zero.first = 1U;
  width_zero.last = 1U;
  width_zero.has_width = true;
  width_zero.width = 0.0;
  layout.columns.push_back(width_zero);

  ColumnLayout hidden_only;
  hidden_only.first = 2U;
  hidden_only.last = 2U;
  hidden_only.hidden = true;
  layout.columns.push_back(hidden_only);

  ColumnLayout outline_only;
  outline_only.first = 3U;
  outline_only.last = 3U;
  outline_only.outline_level = 2U;
  layout.columns.push_back(outline_only);

  ColumnLayout style_nonzero;
  style_nonzero.first = 4U;
  style_nonzero.last = 4U;
  style_nonzero.has_style = true;
  style_nonzero.style_xf = 3U;
  layout.columns.push_back(style_nonzero);

  RowLayout row_style_zero;
  row_style_zero.row = 12U;
  row_style_zero.has_style = true;
  row_style_zero.style_xf = 0U;
  layout.row_overrides.push_back(row_style_zero);

  RowLayout row_style_nonzero;
  row_style_nonzero.row = 13U;
  row_style_nonzero.has_style = true;
  row_style_nonzero.style_xf = 3U;
  layout.row_overrides.push_back(row_style_nonzero);

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  const std::string xml = ReadSheet1Xml(bytes);
  EXPECT_NE(xml.find("<col min=\"1\" max=\"1\" style=\"0\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<col min=\"2\" max=\"2\" width=\"0\" customWidth=\"1\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<col min=\"3\" max=\"3\" hidden=\"1\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<col min=\"4\" max=\"4\" outlineLevel=\"2\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<col min=\"5\" max=\"5\" style=\"3\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<row r=\"13\" s=\"0\" customFormat=\"1\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<row r=\"14\" s=\"3\" customFormat=\"1\"/>"), std::string::npos) << xml;

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const Sheet& loaded = result_or.value().workbook.sheet(0);
  ASSERT_EQ(loaded.layout().columns.size(), 5U);
  EXPECT_TRUE(loaded.layout().columns[0].has_style);
  EXPECT_EQ(loaded.layout().columns[0].style_xf, 0U);
  EXPECT_TRUE(loaded.layout().columns[1].has_width);
  EXPECT_DOUBLE_EQ(loaded.layout().columns[1].width, 0.0);
  EXPECT_TRUE(loaded.layout().columns[2].hidden);
  EXPECT_FALSE(loaded.layout().columns[2].has_width);
  EXPECT_EQ(loaded.layout().columns[3].outline_level, 2U);
  EXPECT_FALSE(loaded.layout().columns[3].has_width);
  EXPECT_TRUE(loaded.layout().columns[4].has_style);
  EXPECT_EQ(loaded.layout().columns[4].style_xf, 3U);

  ASSERT_EQ(loaded.layout().row_overrides.size(), 2U);
  std::vector<RowLayout> rows = loaded.layout().row_overrides;
  std::sort(rows.begin(), rows.end(), [](const RowLayout& a, const RowLayout& b) { return a.row < b.row; });
  EXPECT_EQ(rows[0].row, 12U);
  EXPECT_TRUE(rows[0].has_style);
  EXPECT_EQ(rows[0].style_xf, 0U);
  EXPECT_EQ(rows[1].row, 13U);
  EXPECT_TRUE(rows[1].has_style);
  EXPECT_EQ(rows[1].style_xf, 3U);
  EXPECT_EQ(loaded.cell_at(12U, 0U), nullptr);
  EXPECT_EQ(loaded.cell_at(13U, 0U), nullptr);
}

TEST(SheetLayoutRoundTrip, ColumnWidthAndRowHeightPreserveFullPrecision) {
  // A recalc-save must not drift the dimension metrics. These values need
  // more than the six significant digits the previous %.6g writer kept;
  // 8.7109375 is a width Excel actually stores. The writer now uses a
  // round-trip-safe format so they survive byte-stable.
  constexpr double kPreciseWidth = 8.7109375;
  constexpr double kPreciseHeight = 12.733329999999999;

  Workbook src = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0, 3U, 0U, Value::number(1.0))));
  SheetLayout& layout = src.sheet(0).mutable_layout();
  ColumnLayout col;
  col.first = 0U;
  col.last = 0U;
  col.width = kPreciseWidth;
  layout.columns.push_back(col);
  RowLayout row;
  row.row = 3U;
  row.height = kPreciseHeight;
  layout.row_overrides.push_back(row);

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Sheet& loaded = result_or.value().workbook.sheet(0);
  ASSERT_EQ(loaded.layout().columns.size(), 1U);
  EXPECT_DOUBLE_EQ(loaded.layout().columns[0].width, kPreciseWidth);
  ASSERT_EQ(loaded.layout().row_overrides.size(), 1U);
  EXPECT_DOUBLE_EQ(loaded.layout().row_overrides[0].height, kPreciseHeight);
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
