//
// Stable C ABI tests for drawing shapes and sparkline groups: layout pins,
// insert / list / update / remove, argument validation, and a save and reload.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_shape_t) == (kWasm32 ? 192U : 200U), "fm_shape_t ABI layout changed");
static_assert(sizeof(fm_sparkline_t) == 2U * sizeof(void*), "fm_sparkline_t ABI layout changed");
static_assert(sizeof(fm_sparkline_group_t) == (kWasm32 ? 296U : 304U), "fm_sparkline_group_t ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

void SaveAndReload(fm_workbook_t* wb, WorkbookGuard* out) {
  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb, &buf.data, &buf.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &out->handle), 0) << fm_get_last_error_message();
}

fm_shape_t Callout() {
  fm_shape_t s = FM_STRUCT_INIT(fm_shape_t);
  s.fields = FM_SHAPE_FIELD_NAME | FM_SHAPE_FIELD_PRESET_GEOMETRY | FM_SHAPE_FIELD_ANCHOR | FM_SHAPE_FIELD_TEXT |
             FM_SHAPE_FIELD_FILL | FM_SHAPE_FIELD_LINE | FM_SHAPE_FIELD_LINE_WIDTH | FM_SHAPE_FIELD_ROTATION;
  s.name = "Callout";
  s.preset_geometry = "wedgeRectCallout";
  s.text = "Hello\nWorld";
  s.anchor_kind = FM_ANCHOR_KIND_TWO_CELL;
  s.edit_as = FM_ANCHOR_EDIT_AS_ONE_CELL;
  s.from_row = 2;
  s.from_col = 1;
  s.to_row = 6;
  s.to_col = 4;
  s.fill_kind = FM_SHAPE_PAINT_SOLID;
  s.fill_color.kind = kFmColorRgb;
  s.fill_color.rgb = 0xFFFFCC00U;
  s.line_kind = FM_SHAPE_PAINT_NONE;
  s.line_width_emu = 12700;
  s.rotation = 300000;
  return s;
}

TEST(FormulonCApiShapes, EnumOrdinalsAndFieldBits) {
  EXPECT_EQ(FM_SHAPE_PAINT_UNSET, -1);
  EXPECT_EQ(FM_SHAPE_PAINT_SOLID, 0);
  EXPECT_EQ(FM_SHAPE_PAINT_NONE, 1);
  EXPECT_EQ(FM_SHAPE_PAINT_OTHER, 2);
  EXPECT_EQ(FM_SHAPE_FIELD_HIDDEN, 0x200u);
  EXPECT_EQ(FM_SPARKLINE_STACKED, 2);
  EXPECT_EQ(FM_SPARKLINE_AXIS_CUSTOM, 2);
  EXPECT_EQ(FM_SPARKLINE_EMPTY_SPAN, 2);
}

TEST(FormulonCApiShapes, InsertListUpdateRemoveAndReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  size_t count = 99;
  ASSERT_EQ(fm_sheet_get_shape_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  const fm_shape_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(id, 2U);

  fm_shape_t connector = FM_STRUCT_INIT(fm_shape_t);
  connector.fields = FM_SHAPE_FIELD_PRESET_GEOMETRY | FM_SHAPE_FIELD_ANCHOR | FM_SHAPE_FIELD_LINE;
  connector.preset_geometry = "straightConnector1";
  connector.anchor_kind = FM_ANCHOR_KIND_ONE_CELL;
  connector.from_row = 10;
  connector.from_col_offset_emu = 500;
  connector.width_emu = 914400;
  connector.height_emu = 0;
  connector.line_kind = FM_SHAPE_PAINT_SOLID;
  connector.line_color.kind = kFmColorRgb;
  connector.line_color.rgb = 0xFF00FF00U;
  uint32_t line_id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &connector, &line_id), 0) << fm_get_last_error_message();

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  ASSERT_EQ(fm_sheet_get_shape_count(again.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  fm_shape_t out = FM_STRUCT_INIT(fm_shape_t);
  ASSERT_EQ(fm_sheet_get_shape_at(again.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.object_id, id);
  EXPECT_EQ(out.kind, FM_DRAWING_OBJECT_SHAPE);
  EXPECT_STREQ(out.name, "Callout");
  EXPECT_STREQ(out.preset_geometry, "wedgeRectCallout");
  EXPECT_STREQ(out.text, "Hello\nWorld");
  EXPECT_EQ(out.anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(out.edit_as, FM_ANCHOR_EDIT_AS_ONE_CELL);
  EXPECT_EQ(out.from_row, 2U);
  EXPECT_EQ(out.to_col, 4U);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_SOLID);
  EXPECT_EQ(out.fill_color.kind, kFmColorRgb);
  EXPECT_EQ(out.fill_color.rgb, 0xFFFFCC00U);
  EXPECT_EQ(out.line_kind, FM_SHAPE_PAINT_NONE);
  EXPECT_EQ(out.line_width_emu, 12700);
  EXPECT_EQ(out.rotation, 300000);
  EXPECT_EQ(out.fields, 0U);
  ASSERT_EQ(fm_sheet_get_shape_at(again.handle, 0, 1, &out), 0);
  EXPECT_EQ(out.kind, FM_DRAWING_OBJECT_CONNECTOR);
  EXPECT_EQ(out.from_col_offset_emu, 500);
  EXPECT_EQ(out.width_emu, 914400);
  EXPECT_EQ(out.line_color.kind, kFmColorRgb);
  EXPECT_EQ(out.line_color.rgb, 0xFF00FF00U);

  fm_shape_t patch = FM_STRUCT_INIT(fm_shape_t);
  patch.fields = FM_SHAPE_FIELD_NAME | FM_SHAPE_FIELD_HIDDEN | FM_SHAPE_FIELD_FILL;
  patch.name = "Renamed";
  patch.hidden = 1;
  patch.fill_kind = FM_SHAPE_PAINT_NONE;
  ASSERT_EQ(fm_sheet_update_shape(again.handle, 0, id, &patch), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_shape_at(again.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Renamed");
  EXPECT_EQ(out.hidden, 1);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_NONE);
  EXPECT_STREQ(out.preset_geometry, "wedgeRectCallout");  // unnamed fields stay
  EXPECT_EQ(out.rotation, 300000);

  ASSERT_EQ(fm_sheet_remove_shape(again.handle, 0, id), 0);
  ASSERT_EQ(fm_sheet_get_shape_count(again.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  EXPECT_EQ(fm_sheet_remove_shape(again.handle, 0, id), FM_STATUS_NOT_FOUND);
}

TEST(FormulonCApiShapes, ThemeColorsRoundTripAndAreValidated) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_t spec = Callout();
  spec.fill_color = fm_color_spec_t{};
  spec.fill_color.kind = kFmColorTheme;
  spec.fill_color.theme = 4;
  spec.fill_color.tint = 0.4;
  spec.line_kind = FM_SHAPE_PAINT_SOLID;
  spec.line_color.kind = kFmColorTheme;
  spec.line_color.theme = 1;
  spec.line_color.tint = -0.25;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  fm_shape_t out = FM_STRUCT_INIT(fm_shape_t);
  ASSERT_EQ(fm_sheet_get_shape_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_SOLID);
  EXPECT_EQ(out.fill_color.kind, kFmColorTheme);
  EXPECT_EQ(out.fill_color.theme, 4U);
  EXPECT_DOUBLE_EQ(out.fill_color.tint, 0.4);
  EXPECT_EQ(out.line_color.kind, kFmColorTheme);
  EXPECT_EQ(out.line_color.theme, 1U);
  EXPECT_DOUBLE_EQ(out.line_color.tint, -0.25);

  struct Bad {
    uint8_t kind;
    uint32_t theme;
    double tint;
  };
  const Bad bad[] = {
      {kFmColorIndexed, 0, 0.0}, {kFmColorAuto, 0, 0.0},   {kFmColorTheme, 12, 0.0},
      {kFmColorTheme, 4, 1.5},   {kFmColorTheme, 4, -1.5}, {kFmColorTheme, 4, std::numeric_limits<double>::quiet_NaN()},
  };
  for (const Bad& b : bad) {
    fm_shape_t s = Callout();
    s.fill_color = fm_color_spec_t{};
    s.fill_color.kind = b.kind;
    s.fill_color.theme = b.theme;
    s.fill_color.tint = b.tint;
    EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &s, &id), kInvalidArgument)
        << "kind=" << int(b.kind) << " theme=" << b.theme << " tint=" << b.tint;
  }

  // A colour on a paint that is not solid is refused rather than dropped.
  fm_shape_t none = Callout();
  none.fill_kind = FM_SHAPE_PAINT_NONE;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &none, &id), kInvalidArgument);
  none = Callout();
  none.line_color.kind = kFmColorRgb;
  none.line_color.rgb = 0xFF000000U;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &none, &id), kInvalidArgument);
  // Without the field bit the colour is not read at all.
  none.fields &= ~FM_SHAPE_FIELD_LINE;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &none, &id), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiShapes, Validation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  uint32_t id = 0;
  fm_shape_t spec = Callout();
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 7, &spec, &id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  spec.anchor_kind = 9;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.fields = FM_SHAPE_FIELD_PRESET_GEOMETRY;  // no anchor
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.fields |= 0x80000000u;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.fill_kind = FM_SHAPE_PAINT_OTHER;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.to_row = 0;  // two-cell box ends before it starts
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  size_t count = 5;
  ASSERT_EQ(fm_sheet_get_shape_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  fm_shape_t out = FM_STRUCT_INIT(fm_shape_t);
  EXPECT_EQ(fm_sheet_get_shape_at(wb.handle, 0, 0, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_shape_at(wb.handle, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_shape_count(wb.handle, 0, nullptr), kBindingNullPointer);
}

fm_sparkline_group_t ColumnGroup() {
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  g.type = FM_SPARKLINE_COLUMN;
  g.negative = 1;
  g.markers = 1;
  g.display_x_axis = 1;
  g.min_axis_type = FM_SPARKLINE_AXIS_CUSTOM;
  g.has_manual_min = 1;
  g.manual_min = -1.5;
  g.line_weight = 1.25;
  g.empty_cells_as = FM_SPARKLINE_EMPTY_ZERO;
  g.color_series.kind = kFmColorRgb;
  g.color_series.rgb = 0xFF336699U;
  g.color_negative.kind = kFmColorTheme;
  g.color_negative.theme = 5;
  g.color_negative.tint = -0.25;
  g.color_low.kind = kFmColorIndexed;
  g.color_low.indexed = 12;
  g.color_first.kind = kFmColorAuto;
  return g;
}

TEST(FormulonCApiSparklines, AddReadUpdateReplaceRemoveAndReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  size_t count = 9;
  ASSERT_EQ(fm_sheet_get_sparkline_group_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}, {"Sheet1!A2:D2", "F2"}};
  fm_sparkline_group_t g = ColumnGroup();
  std::size_t group_index = 99;
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 2, &group_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(group_index, 0U);
  g.type = FM_SPARKLINE_LINE;
  g.markers = 0;
  g.has_manual_min = 0;
  g.color_series = fm_color_spec_t{};
  const fm_sparkline_t more[] = {{"Sheet1!A9:D9", "G9"}};
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, more, 1, &group_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(group_index, 1U);

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  ASSERT_EQ(fm_sheet_get_sparkline_group_count(again.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  fm_sparkline_group_t out = FM_STRUCT_INIT(fm_sparkline_group_t);
  ASSERT_EQ(fm_sheet_get_sparkline_group(again.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.type, FM_SPARKLINE_COLUMN);
  EXPECT_EQ(out.negative, 1);
  EXPECT_EQ(out.markers, 1);
  EXPECT_EQ(out.high, 0);
  EXPECT_EQ(out.display_x_axis, 1);
  EXPECT_EQ(out.min_axis_type, FM_SPARKLINE_AXIS_CUSTOM);
  EXPECT_EQ(out.max_axis_type, FM_SPARKLINE_AXIS_INDIVIDUAL);
  EXPECT_EQ(out.has_manual_min, 1);
  EXPECT_DOUBLE_EQ(out.manual_min, -1.5);
  EXPECT_EQ(out.has_manual_max, 0);
  EXPECT_DOUBLE_EQ(out.line_weight, 1.25);
  EXPECT_EQ(out.empty_cells_as, FM_SPARKLINE_EMPTY_ZERO);
  EXPECT_EQ(out.color_series.kind, kFmColorRgb);
  EXPECT_EQ(out.color_series.rgb, 0xFF336699U);
  EXPECT_EQ(out.color_negative.kind, kFmColorTheme);
  EXPECT_EQ(out.color_negative.theme, 5U);
  EXPECT_DOUBLE_EQ(out.color_negative.tint, -0.25);
  EXPECT_EQ(out.color_low.indexed, 12U);
  EXPECT_EQ(out.color_first.kind, kFmColorAuto);
  EXPECT_EQ(out.color_high.kind, kFmColorNone);
  EXPECT_STREQ(out.date_axis_ref, "");
  ASSERT_EQ(out.sparkline_count, 2U);
  fm_sparkline_t sp{};
  ASSERT_EQ(fm_sheet_get_sparkline(again.handle, 0, 0, 1, &sp), 0);
  EXPECT_STREQ(sp.formula, "Sheet1!A2:D2");
  EXPECT_STREQ(sp.location, "F2");
  EXPECT_EQ(fm_sheet_get_sparkline(again.handle, 0, 0, 2, &sp), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_sparkline(again.handle, 0, 5, 0, &sp), FM_STATUS_INDEX_OUT_OF_RANGE);

  // Update without items rewrites the group's modelled fields and keeps its sparklines.
  out.type = FM_SPARKLINE_STACKED;
  out.date_axis = 1;
  const char* date_ref = "Sheet1!A1:D1";
  out.date_axis_ref = date_ref;
  out.high = 1;
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, 0, &out, nullptr, 0), 0) << fm_get_last_error_message();
  fm_sparkline_group_t changed = FM_STRUCT_INIT(fm_sparkline_group_t);
  ASSERT_EQ(fm_sheet_get_sparkline_group(again.handle, 0, 0, &changed), 0);
  EXPECT_EQ(changed.type, FM_SPARKLINE_STACKED);
  EXPECT_EQ(changed.high, 1);
  EXPECT_EQ(changed.date_axis, 1);
  EXPECT_STREQ(changed.date_axis_ref, "Sheet1!A1:D1");
  EXPECT_EQ(changed.sparkline_count, 2U);

  // Update with items replaces the sparklines; the group read back and written
  // unchanged keeps every attribute.
  const fm_sparkline_t replacement[] = {{"Sheet1!B1:B5", "H1"}};
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, 0, &changed, replacement, 1), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_sparkline_group(again.handle, 0, 0, &changed), 0);
  EXPECT_EQ(changed.sparkline_count, 1U);
  EXPECT_EQ(changed.type, FM_SPARKLINE_STACKED);
  EXPECT_EQ(changed.high, 1);
  EXPECT_STREQ(changed.date_axis_ref, "Sheet1!A1:D1");
  ASSERT_EQ(fm_sheet_get_sparkline(again.handle, 0, 0, 0, &sp), 0);
  EXPECT_STREQ(sp.formula, "Sheet1!B1:B5");
  EXPECT_STREQ(sp.location, "H1");

  ASSERT_EQ(fm_sheet_remove_sparkline_group(again.handle, 0, 0), 0);
  ASSERT_EQ(fm_sheet_get_sparkline_group_count(again.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  ASSERT_EQ(fm_sheet_remove_sparkline_group(again.handle, 0, 0), 0);
  EXPECT_EQ(fm_sheet_remove_sparkline_group(again.handle, 0, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiSparklines, Validation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t good[] = {{"Sheet1!A1:D1", "F1"}};
  fm_sparkline_group_t g = ColumnGroup();
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, nullptr, good, 1, ignored_index()), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, nullptr, 1, ignored_index()), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, good, 0, ignored_index()), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 3, &g, good, 1, ignored_index()),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  const fm_sparkline_t range_location[] = {{"Sheet1!A1:D1", "F1:F2"}};
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, range_location, 1, ignored_index()), kInvalidArgument);
  const fm_sparkline_t no_formula[] = {{"", "F1"}};
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, no_formula, 1, ignored_index()), kInvalidArgument);
  fm_sparkline_group_t bad = g;
  bad.type = 7;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &bad, good, 1, ignored_index()), kInvalidArgument);
  bad = g;
  bad.line_weight = -1.0;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &bad, good, 1, ignored_index()), kInvalidArgument);
  size_t count = 5;
  ASSERT_EQ(fm_sheet_get_sparkline_group_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &g, nullptr, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &g, good, 1), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, nullptr, good, 1), kBindingNullPointer);
  fm_sparkline_group_t out = FM_STRUCT_INIT(fm_sparkline_group_t);
  EXPECT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, 0, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, 0, nullptr), kBindingNullPointer);
}

TEST(FormulonCApiSparklines, UpdateWithItemsIsAllOrNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}, {"Sheet1!A2:D2", "F2"}};
  fm_sparkline_group_t g = ColumnGroup();
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 2, ignored_index()), 0)
      << fm_get_last_error_message();

  fm_sparkline_group_t changed = g;
  changed.type = FM_SPARKLINE_LINE;
  // A group cannot hold zero sparklines (CT_Sparklines requires one), so a
  // zero count is rejected like an invalid item, and neither half lands.
  const fm_sparkline_t none[] = {{"Sheet1!A1:D1", "F1"}};
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &changed, none, 0), kInvalidArgument);
  const fm_sparkline_t bad[] = {{"Sheet1!A1:D1", "F1:F2"}};
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &changed, bad, 1), kInvalidArgument);
  // An invalid group with valid items does not replace the items either.
  fm_sparkline_group_t invalid = g;
  invalid.line_weight = -1.0;
  const fm_sparkline_t one[] = {{"Sheet1!B1:B5", "H1"}};
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &invalid, one, 1), kInvalidArgument);

  fm_sparkline_group_t read = FM_STRUCT_INIT(fm_sparkline_group_t);
  ASSERT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, 0, &read), 0);
  EXPECT_EQ(read.type, FM_SPARKLINE_COLUMN);
  EXPECT_EQ(read.sparkline_count, 2U);
}

}  // namespace
