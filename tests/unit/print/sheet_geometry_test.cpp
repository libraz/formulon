#include "print/sheet_geometry.h"

#include <cstdint>

#include "gtest/gtest.h"
#include "sheet.h"
#include "styles.h"
#include "workbook.h"

namespace formulon {
namespace print {
namespace {

// Calibri 11 on Windows: display `Range.Width` at 96 DPI and the
// print-layout figure the page-break captures agree with.
constexpr double kDisplayMdw = 5.25;
constexpr double kDisplayPad = 3.75;
constexpr double kPrintMdw = 39.0 / 7.0;
constexpr double kPrintPad = 27.0 / 7.0;
constexpr double kStandardRowPt = 102.0 / 7.0;
constexpr double kStandardColChars = 8.43;

void SetNormalFont(Workbook* wb, const char* name, double size) {
  FontRecord& font = wb->mutable_styles().fonts[0];
  font.name = name;
  font.size = size;
}

void AddColumn(Sheet* sheet, std::uint32_t first, std::uint32_t last, double width, bool hidden) {
  ColumnLayout span;
  span.first = first;
  span.last = last;
  span.width = width;
  span.has_width = !hidden;
  span.hidden = hidden;
  sheet->mutable_layout().columns.push_back(span);
}

void AddRow(Sheet* sheet, std::uint32_t row, double height, bool hidden) {
  RowLayout layout;
  layout.row = row;
  layout.height = height;
  layout.has_height = !hidden;
  layout.hidden = hidden;
  sheet->mutable_layout().row_overrides.push_back(layout);
}

TEST(SheetGeometry, CalibriModelDiffersByMode) {
  Workbook wb = Workbook::create();
  const ColumnWidthModel display = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_DOUBLE_EQ(display.points_per_char, kDisplayMdw);
  EXPECT_DOUBLE_EQ(display.padding_pt, kDisplayPad);
  EXPECT_TRUE(display.calibrated);
  EXPECT_EQ(display.normal_font_name, "Calibri");
  EXPECT_DOUBLE_EQ(display.normal_font_size, 11.0);
  EXPECT_STREQ(display.platform, "win");

  const ColumnWidthModel print = resolve_column_width_model(wb.styles(), GeometryMode::kPrint);
  EXPECT_DOUBLE_EQ(print.points_per_char, kPrintMdw);
  EXPECT_DOUBLE_EQ(print.padding_pt, kPrintPad);
  EXPECT_TRUE(print.calibrated);
}

TEST(SheetGeometry, ThirtyCharColumnMatchesMeasuredWidths) {
  Workbook wb = Workbook::create();
  const ColumnWidthModel display = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  const ColumnWidthModel print = resolve_column_width_model(wb.styles(), GeometryMode::kPrint);
  // Measured Range.Width for a 30-character Calibri 11 column at 96 DPI.
  EXPECT_DOUBLE_EQ(column_chars_to_points(30.0, display), 161.25);
  // Print-layout figure, 1197/7.
  EXPECT_DOUBLE_EQ(column_chars_to_points(30.0, print), 171.0);
}

TEST(SheetGeometry, ZeroCharsIsZeroPointsAndInverseRoundTrips) {
  Workbook wb = Workbook::create();
  const ColumnWidthModel model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_EQ(column_chars_to_points(0.0, model), 0.0);
  EXPECT_EQ(column_points_to_chars(0.0, model), 0.0);
  EXPECT_NEAR(column_points_to_chars(161.25, model), 30.0, 1e-12);
}

TEST(SheetGeometry, SampledFontSizesUseTheirOwnCalibration) {
  Workbook wb = Workbook::create();
  SetNormalFont(&wb, "游ゴシック", 11.0);
  ColumnWidthModel model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_TRUE(model.calibrated);
  EXPECT_DOUBLE_EQ(model.points_per_char, 6.0);
  EXPECT_DOUBLE_EQ(model.padding_pt, 3.75);

  SetNormalFont(&wb, "Calibri", 14.0);
  model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_TRUE(model.calibrated);
  EXPECT_DOUBLE_EQ(model.points_per_char, 7.5);
  EXPECT_DOUBLE_EQ(model.padding_pt, 5.25);
}

TEST(SheetGeometry, UncalibratedFontReportsExtrapolated) {
  Workbook wb = Workbook::create();
  SetNormalFont(&wb, "Comic Sans MS", 11.0);
  ColumnWidthModel model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_FALSE(model.calibrated);
  EXPECT_EQ(model.normal_font_name, "Comic Sans MS");
  // The Calibri 11 figure stands in for the unmeasured font.
  EXPECT_DOUBLE_EQ(model.points_per_char, kDisplayMdw);
  EXPECT_DOUBLE_EQ(model.padding_pt, kDisplayPad);

  // A calibrated family at an unsampled size is extrapolated too.
  SetNormalFont(&wb, "Calibri", 13.0);
  model = resolve_column_width_model(wb.styles(), GeometryMode::kPrint);
  EXPECT_FALSE(model.calibrated);
  EXPECT_DOUBLE_EQ(model.points_per_char, kPrintMdw);
}

TEST(SheetGeometry, DefaultsFallBackToExcelStandards) {
  Workbook wb = Workbook::create();
  const Sheet& sheet = wb.sheet(0);
  EXPECT_DOUBLE_EQ(default_column_width_chars(sheet), kStandardColChars);
  EXPECT_DOUBLE_EQ(default_row_height_pt(sheet), kStandardRowPt);
}

TEST(SheetGeometry, SheetDefaultsOverrideStandards) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.mutable_format_defaults().has_default_col_width = true;
  sheet.mutable_format_defaults().default_col_width = 12.0;
  sheet.mutable_format_defaults().has_default_row_height = true;
  sheet.mutable_format_defaults().default_row_height = 20.0;
  EXPECT_DOUBLE_EQ(effective_column_width_chars(sheet, 5), 12.0);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 5), 20.0);
}

TEST(SheetGeometry, OverridesAndHiddenTracksShapeEffectiveSizes) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  AddColumn(&sheet, 1, 1, 10.0, false);
  AddColumn(&sheet, 2, 2, 0.0, true);
  AddRow(&sheet, 1, 30.0, false);
  AddRow(&sheet, 2, 0.0, true);

  EXPECT_DOUBLE_EQ(effective_column_width_chars(sheet, 0), kStandardColChars);
  EXPECT_DOUBLE_EQ(effective_column_width_chars(sheet, 1), 10.0);
  EXPECT_DOUBLE_EQ(effective_column_width_chars(sheet, 2), 0.0);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 0), kStandardRowPt);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 1), 30.0);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 2), 0.0);

  const std::vector<double> widths = column_widths_chars(sheet, 0, 3);
  ASSERT_EQ(widths.size(), 4U);
  EXPECT_DOUBLE_EQ(widths[1], 10.0);
  EXPECT_DOUBLE_EQ(widths[2], 0.0);
  EXPECT_DOUBLE_EQ(widths[3], kStandardColChars);
  const std::vector<double> heights = row_heights_pt(sheet, 1, 3);
  ASSERT_EQ(heights.size(), 3U);
  EXPECT_DOUBLE_EQ(heights[0], 30.0);
  EXPECT_DOUBLE_EQ(heights[1], 0.0);
  EXPECT_DOUBLE_EQ(heights[2], kStandardRowPt);
}

TEST(SheetGeometry, CellRectHasTopLeftOriginAndSkipsHiddenTracks) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  AddColumn(&sheet, 1, 1, 10.0, false);
  AddColumn(&sheet, 2, 2, 0.0, true);
  AddRow(&sheet, 1, 30.0, false);
  AddRow(&sheet, 2, 0.0, true);
  const ColumnWidthModel model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);

  const RectPt a1 = cell_rect_pt(sheet, 0, 0, model);
  EXPECT_DOUBLE_EQ(a1.x, 0.0);
  EXPECT_DOUBLE_EQ(a1.y, 0.0);
  EXPECT_DOUBLE_EQ(a1.width, kStandardColChars * kDisplayMdw + kDisplayPad);
  EXPECT_DOUBLE_EQ(a1.height, kStandardRowPt);

  // D4: column C is hidden (zero width) and row 3 is hidden (zero height).
  const RectPt d4 = cell_rect_pt(sheet, 3, 3, model);
  const double col_a = kStandardColChars * kDisplayMdw + kDisplayPad;
  const double col_b = 10.0 * kDisplayMdw + kDisplayPad;
  EXPECT_DOUBLE_EQ(d4.x, col_a + col_b);
  EXPECT_DOUBLE_EQ(d4.y, kStandardRowPt + 30.0);
  EXPECT_DOUBLE_EQ(d4.width, col_a);
  EXPECT_DOUBLE_EQ(d4.height, kStandardRowPt);

  const RectPt hidden = cell_rect_pt(sheet, 2, 2, model);
  EXPECT_DOUBLE_EQ(hidden.width, 0.0);
  EXPECT_DOUBLE_EQ(hidden.height, 0.0);
}

}  // namespace
}  // namespace print
}  // namespace formulon
