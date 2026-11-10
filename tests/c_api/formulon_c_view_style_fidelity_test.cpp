//
// C ABI tests for the outline / view / style / calc / defined-name surface
// that a load -> save cycle has to keep: `collapsed` flags, `showFormulas` and
// `showOutlineSymbols`, font toggles, gradient fills, custom icon sets, the
// typed `<calcPr>` accessors, the 1904 date system and defined-name
// attributes. Every case crosses the ABI in both directions.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_setters.h"
#include "formulon_c_sheet_view_patch.h"
#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_destroy(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

/// Saves `wb` and loads the bytes into `out`.
void RoundTrip(const WorkbookGuard& wb, WorkbookGuard* out, std::string* part_xml = nullptr,
               const char* part = "xl/worksheets/sheet1.xml") {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
  if (part_xml != nullptr) {
    formulon::io::ZipReader zip;
    ASSERT_TRUE(static_cast<bool>(zip.open(formulon::io::ByteSpan{saved.data, saved.len})));
    auto entry = zip.read_entry(part);
    ASSERT_TRUE(static_cast<bool>(entry));
    part_xml->assign(reinterpret_cast<const char*>(entry.value().data()), entry.value().size());
  }
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &out->handle), 0) << fm_get_last_error_message();
}

bool Has(const std::string& text, const char* needle) {
  return text.find(needle) != std::string::npos;
}

fm_row_layout_info_t RowOverride(const WorkbookGuard& wb, std::uint32_t row) {
  size_t count = 0;
  EXPECT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  for (size_t i = 0; i < count; ++i) {
    fm_row_layout_info_t entry = FM_STRUCT_INIT(fm_row_layout_info_t);
    EXPECT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, i, &entry), 0);
    if (entry.row == row) {
      return entry;
    }
  }
  return {};
}

}  // namespace

// ---------------------------------------------------------------------------
// Outline
// ---------------------------------------------------------------------------

TEST(FormulonCApiOutline, CollapsedIsReadBackAndRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 1, 1, 2), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 2, 2, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_collapsed(wb.handle, 0, 3, 3, 1), 0);
  ASSERT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 2, 3, 3), 0);
  ASSERT_EQ(fm_sheet_set_column_collapsed(wb.handle, 0, 4, 4, 1), 0);

  EXPECT_EQ(RowOverride(wb, 3).collapsed, 1);
  EXPECT_EQ(RowOverride(wb, 1).collapsed, 0);
  EXPECT_EQ(RowOverride(wb, 1).outline_level, 2);

  WorkbookGuard reloaded;
  std::string xml;
  RoundTrip(wb, &reloaded, &xml);
  EXPECT_TRUE(Has(xml, "outlineLevelRow=\"2\"")) << xml;
  EXPECT_TRUE(Has(xml, "outlineLevelCol=\"3\"")) << xml;
  EXPECT_EQ(RowOverride(reloaded, 3).collapsed, 1);
  EXPECT_EQ(RowOverride(reloaded, 2).collapsed, 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_column_layout_count(reloaded.handle, 0, &count), 0);
  int collapsed_spans = 0;
  for (size_t i = 0; i < count; ++i) {
    fm_column_layout_info_t col = FM_STRUCT_INIT(fm_column_layout_info_t);
    ASSERT_EQ(fm_sheet_get_column_layout_at(reloaded.handle, 0, i, &col), 0);
    collapsed_spans += col.collapsed != 0 ? 1 : 0;
  }
  EXPECT_EQ(collapsed_spans, 1);
}

TEST(FormulonCApiOutline, ClearingCollapsedRemovesAnOtherwiseEmptyOverride) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_collapsed(wb.handle, 0, 5, 5, 1), 0);
  ASSERT_EQ(fm_sheet_set_column_collapsed(wb.handle, 0, 1, 2, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_collapsed(wb.handle, 0, 5, 5, 0), 0);
  ASSERT_EQ(fm_sheet_set_column_collapsed(wb.handle, 0, 1, 2, 0), 0);
  ASSERT_EQ(fm_sheet_set_row_collapsed(wb.handle, 0, 9, 9, 0), 0);
  size_t rows = 0;
  size_t cols = 0;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &rows), 0);
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &cols), 0);
  EXPECT_EQ(rows, 0U);
  EXPECT_EQ(cols, 0U);
}

TEST(FormulonCApiOutline, CollapsedSettersRejectBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_row_collapsed(wb.handle, 9, 0, 0, 1), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_row_collapsed(wb.handle, 0, 1048576U, 1048576U, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_collapsed(wb.handle, 0, 3, 2, 1), kInvalidArgument);
}

// ---------------------------------------------------------------------------
// Sheet view
// ---------------------------------------------------------------------------

TEST(FormulonCApiSheetViewFlags, ShowFormulasAndOutlineSymbolsRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_view_info_t view = FM_STRUCT_INIT(fm_sheet_view_info_t);
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &view), 0);
  EXPECT_EQ(view.show_formulas, 0);
  EXPECT_EQ(view.show_outline_symbols, 1);

  ASSERT_EQ(
      SetViewFlag(wb.handle, 0, &fm_sheet_view_patch_t::show_formulas_op, &fm_sheet_view_patch_t::show_formulas, 1), 0);
  ASSERT_EQ(SetViewFlag(wb.handle, 0, &fm_sheet_view_patch_t::show_outline_symbols_op,
                        &fm_sheet_view_patch_t::show_outline_symbols, 0),
            0);
  EXPECT_EQ(
      SetViewFlag(wb.handle, 7, &fm_sheet_view_patch_t::show_formulas_op, &fm_sheet_view_patch_t::show_formulas, 1),
      FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);

  WorkbookGuard reloaded;
  std::string xml;
  RoundTrip(wb, &reloaded, &xml);
  EXPECT_TRUE(Has(xml, "showFormulas=\"1\"")) << xml;
  EXPECT_TRUE(Has(xml, "showOutlineSymbols=\"0\"")) << xml;
  ASSERT_EQ(fm_sheet_get_view(reloaded.handle, 0, &view), 0);
  EXPECT_EQ(view.show_formulas, 1);
  EXPECT_EQ(view.show_outline_symbols, 0);
}

// ---------------------------------------------------------------------------
// Fonts
// ---------------------------------------------------------------------------

TEST(FormulonCApiFontToggles, OutlineShadowCondenseExtendRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_font_t font{};
  font.name = "Arial";
  font.size_pt = 12.0;
  font.has_size_pt = 1;
  font.color.kind = FM_COLOR_KIND_RGB;
  font.color.argb = 0xFF000000U;
  font.has_outline = 1;
  font.outline = 1;
  font.has_shadow = 1;
  font.shadow = 0;
  font.has_condense = 1;
  font.condense = 1;
  font.has_extend = 1;
  font.extend = 1;
  std::uint32_t index = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &font, &index), 0);
  std::uint32_t same = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &font, &same), 0);
  EXPECT_EQ(same, index);
  font.extend = 0;
  std::uint32_t other = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &font, &other), 0);
  EXPECT_NE(other, index);

  fm_cell_xf_t xf{};
  xf.font_index = static_cast<uint32_t>(index);
  xf.apply_font = 1;
  std::uint32_t xf_index = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &xf_index), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, xf_index), 0);

  WorkbookGuard reloaded;
  RoundTrip(wb, &reloaded);
  fm_font_t back{};
  ASSERT_EQ(fm_workbook_get_font_at(reloaded.handle, index, &back), 0);
  EXPECT_EQ(back.has_outline, 1);
  EXPECT_EQ(back.outline, 1);
  EXPECT_EQ(back.has_shadow, 1);
  EXPECT_EQ(back.shadow, 0);
  EXPECT_EQ(back.condense, 1);
  EXPECT_EQ(back.extend, 1);
  EXPECT_STREQ(back.name, "Arial");
}

// ---------------------------------------------------------------------------
// Gradient fills
// ---------------------------------------------------------------------------

namespace {

fm_fill_t GradientFill(fm_gradient_stop_t* stops, size_t count) {
  fm_fill_t fill{};
  fill.has_gradient = 1;
  fill.gradient_type = FM_GRADIENT_TYPE_PATH;
  fill.gradient_angle_deg = 0.0;
  fill.gradient_left = 0.5;
  fill.gradient_right = 0.5;
  fill.gradient_top = 0.25;
  fill.gradient_bottom = 0.75;
  fill.gradient_stops = stops;
  fill.gradient_stop_count = static_cast<uint32_t>(count);
  return fill;
}

}  // namespace

TEST(FormulonCApiGradientFill, AddGetAndDedup) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_gradient_stop_t stops[2]{};
  stops[0].position = 0.0;
  stops[0].color.kind = FM_COLOR_KIND_RGB;
  stops[0].color.argb = 0xFFFF0000U;
  stops[1].position = 1.0;
  stops[1].color.kind = FM_COLOR_KIND_THEME;
  stops[1].color.theme = 4;
  stops[1].color.tint = 0.5;
  fm_fill_t fill = GradientFill(stops, 2);

  std::uint32_t index = 0;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &fill, &index), 0);
  std::uint32_t same = 0;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &fill, &same), 0);
  EXPECT_EQ(same, index);
  stops[1].position = 0.9;
  std::uint32_t other = 0;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &fill, &other), 0);
  EXPECT_NE(other, index);

  fm_fill_t back{};
  ASSERT_EQ(fm_workbook_get_fill_at(wb.handle, index, &back), 0);
  EXPECT_EQ(back.has_gradient, 1);
  EXPECT_EQ(back.gradient_type, FM_GRADIENT_TYPE_PATH);
  EXPECT_EQ(back.gradient_left, 0.5);
  EXPECT_EQ(back.gradient_top, 0.25);
  EXPECT_EQ(back.gradient_bottom, 0.75);
  ASSERT_EQ(back.gradient_stop_count, 2U);
  EXPECT_EQ(back.gradient_stops[0].color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(back.gradient_stops[0].color.argb, 0xFFFF0000U);
  EXPECT_EQ(back.gradient_stops[1].color.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(back.gradient_stops[1].color.theme, 4U);
  EXPECT_EQ(back.gradient_stops[1].position, 1.0);
}

TEST(FormulonCApiGradientFill, PatternFillsReportNoGradient) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_fill_t back{};
  ASSERT_EQ(fm_workbook_get_fill_at(wb.handle, 0, &back), 0);
  EXPECT_EQ(back.has_gradient, 0);
  EXPECT_EQ(back.gradient_stop_count, 0U);
  EXPECT_EQ(back.gradient_stops, nullptr);
}

TEST(FormulonCApiGradientFill, NonFiniteNumbersAndNullStopsAreSanitised) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_fill_t fill = GradientFill(nullptr, 5);
  fill.gradient_angle_deg = std::numeric_limits<double>::infinity();
  std::uint32_t index = 0;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &fill, &index), 0);
  fm_fill_t back{};
  ASSERT_EQ(fm_workbook_get_fill_at(wb.handle, index, &back), 0);
  EXPECT_EQ(back.gradient_angle_deg, 0.0);
  EXPECT_EQ(back.gradient_stop_count, 0U);
}

TEST(FormulonCApiGradientFill, SurvivesSaveLoadOnCellsAndDxfs) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_gradient_stop_t stops[3]{};
  for (int i = 0; i < 3; ++i) {
    stops[i].position = i * 0.5;
    stops[i].color.kind = FM_COLOR_KIND_RGB;
    stops[i].color.argb = 0xFF000000U | (0x40U * static_cast<uint32_t>(i + 1));
  }
  fm_fill_t fill = GradientFill(stops, 3);
  std::uint32_t fill_index = 0;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &fill, &fill_index), 0);
  fm_cell_xf_t xf{};
  xf.fill_index = static_cast<uint32_t>(fill_index);
  xf.apply_fill = 1;
  std::uint32_t xf_index = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &xf_index), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, xf_index), 0);

  fm_dxf_t dxf{};
  dxf.has_fill = 1;
  dxf.fill = GradientFill(stops, 3);
  dxf.fill.gradient_type = FM_GRADIENT_TYPE_LINEAR;
  dxf.fill.gradient_angle_deg = 45.0;
  dxf.fill.gradient_left = dxf.fill.gradient_right = dxf.fill.gradient_top = dxf.fill.gradient_bottom = 0.0;
  std::uint32_t dxf_index = 0;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &dxf_index), 0);

  WorkbookGuard reloaded;
  std::string xml;
  RoundTrip(wb, &reloaded, &xml, "xl/styles.xml");
  EXPECT_TRUE(Has(xml, "<gradientFill type=\"path\" left=\"0.5\" right=\"0.5\" top=\"0.25\" bottom=\"0.75\">")) << xml;
  EXPECT_TRUE(Has(xml, "<gradientFill degree=\"45\">")) << xml;

  fm_fill_t back{};
  ASSERT_EQ(fm_workbook_get_fill_at(reloaded.handle, fill_index, &back), 0);
  ASSERT_EQ(back.has_gradient, 1);
  ASSERT_EQ(back.gradient_stop_count, 3U);
  EXPECT_EQ(back.gradient_stops[2].position, 1.0);
  EXPECT_EQ(back.gradient_stops[2].color.argb, 0xFF0000C0U);

  fm_dxf_t dxf_back{};
  ASSERT_EQ(fm_workbook_get_dxf_at(reloaded.handle, dxf_index, &dxf_back), 0);
  ASSERT_EQ(dxf_back.has_fill, 1);
  EXPECT_EQ(dxf_back.fill.has_gradient, 1);
  EXPECT_EQ(dxf_back.fill.gradient_angle_deg, 45.0);
  EXPECT_EQ(dxf_back.fill.gradient_stop_count, 3U);
}

// ---------------------------------------------------------------------------
// Custom icon sets
// ---------------------------------------------------------------------------

namespace {

fm_cf_rule_t IconRule(const fm_cell_rect_t* sqref, fm_cfvo_t* thresholds, const fm_cf_icon_t* icons,
                      uint32_t icon_count) {
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 4;  // IconSet
  rule.ranges = sqref;
  rule.range_count = 1;
  rule.has_icon_set = 1;
  rule.icon_set = 0;       // Three_Arrows
  thresholds[0].type = 1;  // Percent
  thresholds[0].value = "33";
  thresholds[0].gte = FM_TRISTATE_TRUE;
  thresholds[1].type = 1;
  thresholds[1].value = "67";
  thresholds[1].gte = FM_TRISTATE_TRUE;
  rule.icon_set_thresholds = thresholds;
  rule.icon_set_threshold_count = 2;
  rule.icon_set_show_value = 1;
  rule.icon_set_percent = 1;
  rule.icon_set_icons = icons;
  rule.icon_set_icon_count = icon_count;
  return rule;
}

}  // namespace

TEST(FormulonCApiCustomIcons, RoundTripThroughSaveAndLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cfvo_t thresholds[2]{};
  const fm_cf_icon_t icons[3] = {
      {FM_ICON_SET_NAME_THREE_TRIANGLES, 0}, {FM_ICON_SET_NAME_NO_ICONS, 0}, {FM_ICON_SET_NAME_FIVE_BOXES, 4}};
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(IconRule(&sqref, thresholds, icons, 3)), ignored_id()), 0)
      << fm_get_last_error_message();

  fm_cf_rule_t added = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &added), 0);
  ASSERT_EQ(added.icon_set_icon_count, 3U);
  EXPECT_EQ(added.icon_set_icons[1].icon_set, FM_ICON_SET_NAME_NO_ICONS);

  WorkbookGuard reloaded;
  std::string xml;
  RoundTrip(wb, &reloaded, &xml);
  EXPECT_TRUE(Has(xml, "<x14:cfIcon iconSet=\"3Triangles\" iconId=\"0\"/>")) << xml;
  EXPECT_TRUE(Has(xml, "<x14:cfIcon iconSet=\"NoIcons\" iconId=\"0\"/>")) << xml;
  EXPECT_TRUE(Has(xml, "<x14:cfIcon iconSet=\"5Boxes\" iconId=\"4\"/>")) << xml;
  fm_cf_rule_t back = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(reloaded.handle, 0, 0, &back), 0);
  ASSERT_EQ(back.icon_set_icon_count, 3U);
  EXPECT_EQ(back.icon_set_icons[0].icon_set, FM_ICON_SET_NAME_THREE_TRIANGLES);
  EXPECT_EQ(back.icon_set_icons[1].icon_set, FM_ICON_SET_NAME_NO_ICONS);
  EXPECT_EQ(back.icon_set_icons[2].icon_set, FM_ICON_SET_NAME_FIVE_BOXES);
  EXPECT_EQ(back.icon_set_icons[2].icon_index, 4);
}

TEST(FormulonCApiCustomIcons, RulesWithoutIconsReportNone) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cfvo_t thresholds[2]{};
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(IconRule(&sqref, thresholds, nullptr, 0)), ignored_id()), 0);
  fm_cf_rule_t back = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &back), 0);
  EXPECT_EQ(back.icon_set_icon_count, 0U);
  EXPECT_EQ(back.icon_set_icons, nullptr);
}

TEST(FormulonCApiCustomIcons, MalformedIconListsAreRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cfvo_t thresholds[2]{};
  const fm_cf_icon_t too_few[2] = {{FM_ICON_SET_NAME_THREE_ARROWS, 0}, {FM_ICON_SET_NAME_THREE_ARROWS, 1}};
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(IconRule(&sqref, thresholds, too_few, 2)), ignored_id()),
            kInvalidArgument);
  const fm_cf_icon_t past_last[3] = {
      {FM_ICON_SET_NAME_THREE_ARROWS, 0}, {FM_ICON_SET_NAME_NO_ICONS + 1, 1}, {FM_ICON_SET_NAME_THREE_ARROWS, 2}};
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(IconRule(&sqref, thresholds, past_last, 3)), ignored_id()),
            kInvalidArgument);
  const fm_cf_icon_t negative[3] = {{FM_ICON_SET_NAME_THREE_ARROWS, 0}, {-1, 1}, {FM_ICON_SET_NAME_THREE_ARROWS, 2}};
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(IconRule(&sqref, thresholds, negative, 3)), ignored_id()),
            kInvalidArgument);
  const fm_cf_icon_t id_out_of_set[3] = {
      {FM_ICON_SET_NAME_THREE_ARROWS, 0}, {FM_ICON_SET_NAME_THREE_ARROWS, 3}, {FM_ICON_SET_NAME_THREE_ARROWS, 2}};
  EXPECT_EQ(
      fm_sheet_add_cf_rule(wb.handle, 0, record_arg(IconRule(&sqref, thresholds, id_out_of_set, 3)), ignored_id()),
      kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(IconRule(&sqref, thresholds, nullptr, 3)), ignored_id()),
            kInvalidArgument);
  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

// ---------------------------------------------------------------------------
// calcPr and date system
// ---------------------------------------------------------------------------

TEST(FormulonCApiCalcOptions, DefaultsAndRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  int32_t flag = -1;
  ASSERT_EQ(fm_workbook_get_full_precision(wb.handle, &flag), 0);
  EXPECT_EQ(flag, 1);
  ASSERT_EQ(fm_workbook_get_calc_on_save(wb.handle, &flag), 0);
  EXPECT_EQ(flag, 1);
  ASSERT_EQ(fm_workbook_get_date1904(wb.handle, &flag), 0);
  EXPECT_EQ(flag, 0);
  int32_t mode = FM_REF_MODE_R1C1;
  ASSERT_EQ(fm_workbook_get_ref_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_REF_MODE_A1);

  ASSERT_EQ(fm_workbook_set_full_precision(wb.handle, 0), 0);
  ASSERT_EQ(fm_workbook_set_calc_on_save(wb.handle, 0), 0);
  ASSERT_EQ(fm_workbook_set_ref_mode(wb.handle, FM_REF_MODE_R1C1), 0);
  ASSERT_EQ(fm_workbook_set_date1904(wb.handle, 1), 0);
  EXPECT_EQ(fm_workbook_set_ref_mode(wb.handle, 9), kInvalidArgument);

  WorkbookGuard reloaded;
  std::string xml;
  RoundTrip(wb, &reloaded, &xml, "xl/workbook.xml");
  EXPECT_TRUE(Has(xml, "fullPrecision=\"0\"")) << xml;
  EXPECT_TRUE(Has(xml, "calcOnSave=\"0\"")) << xml;
  EXPECT_TRUE(Has(xml, "refMode=\"R1C1\"")) << xml;
  EXPECT_TRUE(Has(xml, "date1904=\"1\"")) << xml;
  ASSERT_EQ(fm_workbook_get_full_precision(reloaded.handle, &flag), 0);
  EXPECT_EQ(flag, 0);
  ASSERT_EQ(fm_workbook_get_calc_on_save(reloaded.handle, &flag), 0);
  EXPECT_EQ(flag, 0);
  ASSERT_EQ(fm_workbook_get_date1904(reloaded.handle, &flag), 0);
  EXPECT_EQ(flag, 1);
  ASSERT_EQ(fm_workbook_get_ref_mode(reloaded.handle, &mode), 0);
  EXPECT_EQ(mode, FM_REF_MODE_R1C1);
}

TEST(FormulonCApiCalcOptions, NullArgumentsAreRejected) {
  int32_t flag = 0;
  EXPECT_NE(fm_workbook_get_full_precision(nullptr, &flag), 0);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_NE(fm_workbook_get_date1904(wb.handle, nullptr), 0);
  EXPECT_NE(fm_workbook_set_calc_on_save(nullptr, 1), 0);
}

// ---------------------------------------------------------------------------
// Defined-name attributes
// ---------------------------------------------------------------------------

TEST(FormulonCApiDefinedNameAttrs, SetReadBackAndRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_defined_name_t rate = FM_STRUCT_INIT(fm_defined_name_t);
  rate.name = "Rate";
  rate.formula = "Sheet1!$A$1";
  rate.local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK;
  rate.comment = "Tax rate";
  rate.hidden = 1;
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &rate), 0);

  fm_defined_name_t dn = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &dn), 0);
  EXPECT_STREQ(dn.comment, "Tax rate");
  EXPECT_EQ(dn.hidden, 1);

  EXPECT_EQ(fm_workbook_get_defined_name_at(wb.handle, 4, &dn), FM_STATUS_INDEX_OUT_OF_RANGE);

  WorkbookGuard reloaded;
  RoundTrip(wb, &reloaded);
  ASSERT_EQ(fm_workbook_get_defined_name_at(reloaded.handle, 0, &dn), 0);
  EXPECT_STREQ(dn.comment, "Tax rate");
  EXPECT_EQ(dn.hidden, 1);

  // An empty comment clears it.
  fm_defined_name_t cleared = FM_STRUCT_INIT(fm_defined_name_t);
  cleared.name = "Rate";
  cleared.formula = "Sheet1!$A$1";
  cleared.local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK;
  cleared.comment = "";
  cleared.hidden = 1;
  ASSERT_EQ(fm_workbook_set_defined_name(reloaded.handle, &cleared), 0);
  ASSERT_EQ(fm_workbook_get_defined_name_at(reloaded.handle, 0, &dn), 0);
  EXPECT_EQ(dn.comment, nullptr);
  EXPECT_EQ(dn.hidden, 1);
}

// ---------------------------------------------------------------------------
// Sheet format defaults: customHeight / zeroHeight
// ---------------------------------------------------------------------------

TEST(FormulonCApiFormatDefaultFlags, CustomHeightAndZeroHeightRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_format_defaults_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &d), 0);
  EXPECT_EQ(d.custom_height, 0);
  EXPECT_EQ(d.zero_height, 0);

  fm_sheet_format_defaults_patch_t custom = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  custom.default_row_height_pt_op = FM_PATCH_OP_SET;
  custom.default_row_height_pt = 18.0;
  custom.custom_height_op = FM_PATCH_OP_SET;
  custom.custom_height = 1;
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &custom), 0);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &d), 0);
  EXPECT_EQ(d.custom_height, 1);
  EXPECT_EQ(d.zero_height, 0);

  WorkbookGuard reloaded;
  std::string xml;
  RoundTrip(wb, &reloaded, &xml);
  EXPECT_TRUE(Has(xml, "customHeight=\"1\"")) << xml;
  fm_sheet_format_defaults_t back = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(reloaded.handle, 0, &back), 0);
  EXPECT_EQ(back.custom_height, 1);
  EXPECT_EQ(back.default_row_height_pt, 18.0);

  // zeroHeight is patched on its own, and SET 0 clears each flag.
  fm_sheet_format_defaults_patch_t zero = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  zero.zero_height_op = FM_PATCH_OP_SET;
  zero.zero_height = 1;
  ASSERT_EQ(fm_sheet_update_format_defaults(reloaded.handle, 0, &zero), 0);
  WorkbookGuard again;
  RoundTrip(reloaded, &again, &xml);
  EXPECT_TRUE(Has(xml, "zeroHeight=\"1\"")) << xml;
  ASSERT_EQ(fm_sheet_get_format_defaults(again.handle, 0, &back), 0);
  EXPECT_EQ(back.zero_height, 1);
  EXPECT_EQ(back.effective_default_row_height_pt, 0.0);
  EXPECT_EQ(back.custom_height, 1);
  fm_sheet_format_defaults_patch_t cleared = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  cleared.custom_height_op = FM_PATCH_OP_SET;
  cleared.custom_height = 0;
  cleared.zero_height_op = FM_PATCH_OP_SET;
  cleared.zero_height = 0;
  ASSERT_EQ(fm_sheet_update_format_defaults(again.handle, 0, &cleared), 0);
  ASSERT_EQ(fm_sheet_get_format_defaults(again.handle, 0, &back), 0);
  EXPECT_EQ(back.custom_height, 0);
  EXPECT_EQ(back.zero_height, 0);
}
