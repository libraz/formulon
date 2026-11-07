// Sheet tab colour through fm_sheet_get_tab_color / fm_sheet_set_tab_color.

#include <cmath>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

std::string SheetPrXml(fm_workbook_t* wb) {
  const char* xml = nullptr;
  EXPECT_EQ(fm_sheet_get_sheet_pr_xml(wb, 0, &xml), 0);
  return xml == nullptr ? std::string() : std::string(xml);
}

fm_color_spec_t GetColor(fm_workbook_t* wb) {
  fm_color_spec_t color{};
  EXPECT_EQ(fm_sheet_get_tab_color(wb, 0, &color), 0);
  return color;
}

TEST(FormulonCApiTabColor, FreshSheetHasNoColor) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(GetColor(wb.handle).kind, kFmColorNone);
}

TEST(FormulonCApiTabColor, RgbRoundTripsAndIsWrittenAsArgbHex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_color_spec_t color{};
  color.kind = kFmColorRgb;
  color.rgb = 0xFF1A2B3CU;
  ASSERT_EQ(fm_sheet_set_tab_color(wb.handle, 0, &color), 0);
  const fm_color_spec_t got = GetColor(wb.handle);
  EXPECT_EQ(got.kind, kFmColorRgb);
  EXPECT_EQ(got.rgb, 0xFF1A2B3CU);
  EXPECT_EQ(SheetPrXml(wb.handle), "<sheetPr><tabColor rgb=\"FF1A2B3C\"/></sheetPr>");
}

TEST(FormulonCApiTabColor, ThemeIndexedAndAutoRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_color_spec_t theme{};
  theme.kind = kFmColorTheme;
  theme.theme = 4;
  theme.tint = -0.25;
  ASSERT_EQ(fm_sheet_set_tab_color(wb.handle, 0, &theme), 0);
  fm_color_spec_t got = GetColor(wb.handle);
  EXPECT_EQ(got.kind, kFmColorTheme);
  EXPECT_EQ(got.theme, 4U);
  EXPECT_DOUBLE_EQ(got.tint, -0.25);

  fm_color_spec_t indexed{};
  indexed.kind = kFmColorIndexed;
  indexed.indexed = 12;
  ASSERT_EQ(fm_sheet_set_tab_color(wb.handle, 0, &indexed), 0);
  got = GetColor(wb.handle);
  EXPECT_EQ(got.kind, kFmColorIndexed);
  EXPECT_EQ(got.indexed, 12U);

  fm_color_spec_t automatic{};
  automatic.kind = kFmColorAuto;
  ASSERT_EQ(fm_sheet_set_tab_color(wb.handle, 0, &automatic), 0);
  EXPECT_EQ(GetColor(wb.handle).kind, kFmColorAuto);
}

TEST(FormulonCApiTabColor, NullAndNoneClearTheColor) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_color_spec_t color{};
  color.kind = kFmColorRgb;
  color.rgb = 0xFF00FF00U;

  ASSERT_EQ(fm_sheet_set_tab_color(wb.handle, 0, &color), 0);
  ASSERT_EQ(fm_sheet_set_tab_color(wb.handle, 0, nullptr), 0);
  EXPECT_EQ(GetColor(wb.handle).kind, kFmColorNone);
  EXPECT_EQ(SheetPrXml(wb.handle), "");

  ASSERT_EQ(fm_sheet_set_tab_color(wb.handle, 0, &color), 0);
  fm_color_spec_t none{};
  ASSERT_EQ(fm_sheet_set_tab_color(wb.handle, 0, &none), 0);
  EXPECT_EQ(GetColor(wb.handle).kind, kFmColorNone);
}

TEST(FormulonCApiTabColor, KeepsTheRestOfSheetPr) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(
      fm_sheet_set_sheet_pr_xml(wb.handle, 0, "<sheetPr codeName=\"S1\"><pageSetUpPr fitToPage=\"1\"/></sheetPr>"), 0);
  fm_color_spec_t color{};
  color.kind = kFmColorRgb;
  color.rgb = 0xFFFF0000U;
  ASSERT_EQ(fm_sheet_set_tab_color(wb.handle, 0, &color), 0);
  const std::string xml = SheetPrXml(wb.handle);
  EXPECT_NE(xml.find("codeName=\"S1\""), std::string::npos);
  EXPECT_NE(xml.find("fitToPage=\"1\""), std::string::npos);
  EXPECT_LT(xml.find("<tabColor"), xml.find("<pageSetUpPr"));

  ASSERT_EQ(fm_sheet_set_tab_color(wb.handle, 0, nullptr), 0);
  const std::string cleared = SheetPrXml(wb.handle);
  EXPECT_EQ(cleared.find("tabColor"), std::string::npos);
  EXPECT_NE(cleared.find("codeName=\"S1\""), std::string::npos);
  EXPECT_NE(cleared.find("fitToPage=\"1\""), std::string::npos);
}

TEST(FormulonCApiTabColor, ReadsAnExistingSixDigitRgb) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_sheet_pr_xml(wb.handle, 0, "<sheetPr><tabColor rgb=\"336699\"/></sheetPr>"), 0);
  const fm_color_spec_t got = GetColor(wb.handle);
  EXPECT_EQ(got.kind, kFmColorRgb);
  EXPECT_EQ(got.rgb, 0xFF336699U);
}

TEST(FormulonCApiTabColor, RejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_color_spec_t color{};
  EXPECT_EQ(fm_sheet_get_tab_color(wb.handle, 0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_tab_color(nullptr, 0, &color), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_tab_color(wb.handle, 99, &color), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_tab_color(wb.handle, 99, &color), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);

  fm_color_spec_t unknown{};
  unknown.kind = 9;
  EXPECT_EQ(fm_sheet_set_tab_color(wb.handle, 0, &unknown), FM_STATUS_INVALID_ARGUMENT);

  fm_color_spec_t nan_tint{};
  nan_tint.kind = kFmColorTheme;
  nan_tint.tint = std::nan("");
  EXPECT_EQ(fm_sheet_set_tab_color(wb.handle, 0, &nan_tint), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(GetColor(wb.handle).kind, kFmColorNone);
}

}  // namespace
