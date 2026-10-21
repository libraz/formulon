#include "print/sheet_geometry.h"

#include <cmath>
#include <cstdint>
#include <limits>

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
// Stored widths (padding included) Excel writes for Calibri 11 columns of
// 8.43 (the base-8 default, 64 px) and 30 displayed characters (215 px).
constexpr double kStandardColChars = 2340.0 / 256.0;
constexpr double kStandardColPt = 48.0;
constexpr double kThirtyCharsStored = 7862.0 / 256.0;

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

TEST(SheetGeometry, ThirtyCharColumnMatchesExcelWidths) {
  Workbook wb = Workbook::create();
  const ColumnWidthModel display = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  const ColumnWidthModel print = resolve_column_width_model(wb.styles(), GeometryMode::kPrint);
  // Range.Width for a 30-character Calibri 11 column at 96 DPI.
  EXPECT_DOUBLE_EQ(column_chars_to_points(kThirtyCharsStored, display), 161.25);
  // Print-layout figure, 1197/7.
  EXPECT_DOUBLE_EQ(column_chars_to_points(kThirtyCharsStored, print), 171.0);
}

TEST(SheetGeometry, ZeroCharsIsZeroPointsAndInverseRoundTrips) {
  Workbook wb = Workbook::create();
  const ColumnWidthModel model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_EQ(column_chars_to_points(0.0, model), 0.0);
  EXPECT_EQ(column_points_to_chars(0.0, model), 0.0);
  EXPECT_DOUBLE_EQ(column_points_to_chars(161.25, model), kThirtyCharsStored);
  EXPECT_DOUBLE_EQ(column_points_to_chars(kStandardColPt, model), kStandardColChars);
}

TEST(SheetGeometry, StoredWidthSnapsToWholeScreenPixels) {
  Workbook wb = Workbook::create();
  const ColumnWidthModel model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_EQ(model.max_digit_width_px, 7);
  EXPECT_EQ(model.padding_px, 5);
  // The stored width already carries the padding: 9.140625 is 64 px, not
  // 9.140625 characters plus padding.
  EXPECT_DOUBLE_EQ(column_chars_to_points(kStandardColChars, model), kStandardColPt);
  // 10 stored characters: Truncate((2560 + 18) / 256 * 7) = 70 px.
  EXPECT_DOUBLE_EQ(column_chars_to_points(10.0, model), 52.5);
  for (int px = 1; px <= 300; ++px) {
    const double points = px * 0.75;
    EXPECT_DOUBLE_EQ(column_chars_to_points(column_points_to_chars(points, model), model), points) << px;
  }
}

TEST(SheetGeometry, DefaultColumnFollowsBaseColWidthAndNormalFont) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  const ColumnWidthModel calibri = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  // Base 10: 10 * 7 + 5 = 75 px, rounded up to 80 px (10.71 characters).
  sheet.mutable_format_defaults().base_col_width = 10.0;
  EXPECT_DOUBLE_EQ(default_column_width_chars(sheet, calibri), 2925.0 / 256.0);
  EXPECT_DOUBLE_EQ(effective_column_width_pt(sheet, 0, calibri), 60.0);

  // 游ゴシック 11, base 8: 8 * 8 + 5 = 69 px, rounded up to 72 px (8.38).
  SetNormalFont(&wb, "游ゴシック", 11.0);
  sheet.mutable_format_defaults().base_col_width = 8.0;
  const ColumnWidthModel yu = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_DOUBLE_EQ(default_column_width_chars(sheet, yu), 9.0);
  EXPECT_DOUBLE_EQ(effective_column_width_pt(sheet, 0, yu), 54.0);
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

TEST(SheetGeometry, UnsupportedFontSizesFallBackWithoutChangingReportedSize) {
  Workbook wb = Workbook::create();
  const GeometryMode modes[] = {GeometryMode::kDisplay, GeometryMode::kPrint};
  const double unsupported_sizes[] = {std::numeric_limits<double>::quiet_NaN(),
                                      std::numeric_limits<double>::infinity(),
                                      -std::numeric_limits<double>::infinity(),
                                      std::numeric_limits<double>::max(),
                                      -std::numeric_limits<double>::max(),
                                      11.5,
                                      -1.0,
                                      0.0,
                                      13.0};

  for (const GeometryMode mode : modes) {
    const double fallback_mdw = mode == GeometryMode::kDisplay ? kDisplayMdw : kPrintMdw;
    const double fallback_padding = mode == GeometryMode::kDisplay ? kDisplayPad : kPrintPad;
    for (const double size : unsupported_sizes) {
      SCOPED_TRACE(::testing::Message() << "mode=" << static_cast<int>(mode) << ", size=" << size);
      SetNormalFont(&wb, "Calibri", size);

      const ColumnWidthModel model = resolve_column_width_model(wb.styles(), mode);
      EXPECT_FALSE(model.calibrated);
      EXPECT_EQ(model.normal_font_name, "Calibri");
      EXPECT_DOUBLE_EQ(model.points_per_char, fallback_mdw);
      EXPECT_DOUBLE_EQ(model.padding_pt, fallback_padding);
      if (std::isnan(size)) {
        EXPECT_TRUE(std::isnan(model.normal_font_size));
      } else {
        EXPECT_DOUBLE_EQ(model.normal_font_size, size);
      }
    }

    SetNormalFont(&wb, "Calibri", 11.0);
    const ColumnWidthModel sampled = resolve_column_width_model(wb.styles(), mode);
    EXPECT_TRUE(sampled.calibrated);
    EXPECT_DOUBLE_EQ(sampled.normal_font_size, 11.0);
    EXPECT_DOUBLE_EQ(sampled.points_per_char, mode == GeometryMode::kDisplay ? kDisplayMdw : kPrintMdw);
    EXPECT_DOUBLE_EQ(sampled.padding_pt, mode == GeometryMode::kDisplay ? kDisplayPad : kPrintPad);
  }
}

TEST(SheetGeometry, UncalibratedFontReportsExtrapolated) {
  Workbook wb = Workbook::create();
  SetNormalFont(&wb, "Comic Sans MS", 11.0);
  ColumnWidthModel model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_FALSE(model.calibrated);
  EXPECT_EQ(model.normal_font_name, "Comic Sans MS");
  // The Calibri 11 figure stands in for a font without a table.
  EXPECT_DOUBLE_EQ(model.points_per_char, kDisplayMdw);
  EXPECT_DOUBLE_EQ(model.padding_pt, kDisplayPad);

  // A calibrated family at an unsampled size is extrapolated too.
  SetNormalFont(&wb, "Calibri", 13.0);
  model = resolve_column_width_model(wb.styles(), GeometryMode::kPrint);
  EXPECT_FALSE(model.calibrated);
  EXPECT_DOUBLE_EQ(model.points_per_char, kPrintMdw);
}

TEST(SheetGeometry, LocalizedNormalStyleIsFoundByBuiltinId) {
  Workbook wb = Workbook::create();
  StylesTable& styles = wb.mutable_styles();
  FontRecord yu;
  yu.name = "游ゴシック";
  yu.size = 11.0;
  styles.fonts.push_back(yu);
  CellXf style_xf;
  style_xf.font_index = 1U;
  styles.cell_style_xfs.push_back(style_xf);
  styles.cell_styles[0].name = "標準";
  styles.cell_styles[0].xf_id = 1U;

  const FontRecord* font = resolve_normal_font(styles);
  ASSERT_NE(font, nullptr);
  EXPECT_EQ(font->name, "游ゴシック");
}

TEST(SheetGeometry, DefaultsFallBackToExcelStandards) {
  Workbook wb = Workbook::create();
  const Sheet& sheet = wb.sheet(0);
  const ColumnWidthModel model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_DOUBLE_EQ(default_column_width_chars(sheet, model), kStandardColChars);
  EXPECT_DOUBLE_EQ(default_row_height_pt(sheet), kStandardRowPt);
}

TEST(SheetGeometry, SheetDefaultsOverrideStandards) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.mutable_format_defaults().has_default_col_width = true;
  sheet.mutable_format_defaults().default_col_width = 12.0;
  sheet.mutable_format_defaults().has_default_row_height = true;
  sheet.mutable_format_defaults().default_row_height = 20.0;
  const ColumnWidthModel model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);
  EXPECT_DOUBLE_EQ(effective_column_width_chars(sheet, 5, model), 12.0);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 5), 20.0);
}

TEST(SheetGeometry, DefaultHiddenRowsHaveZeroImplicitExtent) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.mutable_format_defaults().zero_height = true;

  EXPECT_DOUBLE_EQ(default_row_height_pt(sheet), kStandardRowPt);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 0), 0.0);
  EXPECT_DOUBLE_EQ(row_offset_pt(sheet, 4), 0.0);

  RowLayout visible;
  visible.row = 1U;
  visible.hidden = false;
  visible.has_height = false;  // Explicitly visible, using the nominal default.
  sheet.mutable_layout().row_overrides.push_back(visible);
  AddRow(&sheet, 2U, 25.0, false);
  AddRow(&sheet, 3U, 0.0, true);

  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 0), 0.0);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 1), kStandardRowPt);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 2), 25.0);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 3), 0.0);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 4), 0.0);

  const std::vector<double> heights = row_heights_pt(sheet, 0, 4);
  ASSERT_EQ(heights.size(), 5U);
  EXPECT_DOUBLE_EQ(heights[0], 0.0);
  EXPECT_DOUBLE_EQ(heights[1], kStandardRowPt);
  EXPECT_DOUBLE_EQ(heights[2], 25.0);
  EXPECT_DOUBLE_EQ(heights[3], 0.0);
  EXPECT_DOUBLE_EQ(heights[4], 0.0);
  EXPECT_DOUBLE_EQ(row_offset_pt(sheet, 4), kStandardRowPt + 25.0);
}

TEST(SheetGeometry, OverridesAndHiddenTracksShapeEffectiveSizes) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  AddColumn(&sheet, 1, 1, 10.0, false);
  AddColumn(&sheet, 2, 2, 0.0, true);
  AddRow(&sheet, 1, 30.0, false);
  AddRow(&sheet, 2, 0.0, true);
  const ColumnWidthModel model = resolve_column_width_model(wb.styles(), GeometryMode::kDisplay);

  EXPECT_DOUBLE_EQ(effective_column_width_chars(sheet, 0, model), kStandardColChars);
  EXPECT_DOUBLE_EQ(effective_column_width_chars(sheet, 1, model), 10.0);
  EXPECT_DOUBLE_EQ(effective_column_width_chars(sheet, 2, model), 0.0);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 0), kStandardRowPt);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 1), 30.0);
  EXPECT_DOUBLE_EQ(effective_row_height_pt(sheet, 2), 0.0);

  const std::vector<double> widths = column_widths_chars(sheet, 0, 3, model);
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

TEST(SheetGeometry, RowOffsetUsesFirstUnsortedDuplicateOverrideOnce) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  AddRow(&sheet, 5U, 20.0, false);
  AddRow(&sheet, 2U, 30.0, false);
  AddRow(&sheet, 5U, 40.0, false);

  const double expected = kStandardRowPt * 6.0 + (20.0 - kStandardRowPt) + (30.0 - kStandardRowPt);
  EXPECT_DOUBLE_EQ(row_offset_pt(sheet, 6U), expected);
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
  EXPECT_DOUBLE_EQ(a1.width, kStandardColPt);
  EXPECT_DOUBLE_EQ(a1.height, kStandardRowPt);

  // D4: column C is hidden (zero width) and row 3 is hidden (zero height).
  const RectPt d4 = cell_rect_pt(sheet, 3, 3, model);
  const double col_a = kStandardColPt;
  const double col_b = 52.5;
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
