//
// Stable C ABI tests for drawing shapes and sparkline groups: layout pins,
// insert / list / update / remove, argument validation, and a save and reload.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_shape) == (kWasm32 ? 136U : 144U), "fm_shape ABI layout changed");
static_assert(sizeof(fm_sparkline) == 2U * sizeof(void*), "fm_sparkline ABI layout changed");
static_assert(sizeof(fm_sparkline_group) == (kWasm32 ? 288U : 296U), "fm_sparkline_group ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

void SaveAndReload(fm_workbook_t* wb, WorkbookGuard* out) {
  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb, &buf.data, &buf.len), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &out->handle), 0) << fm_last_error_message();
}

fm_shape Callout() {
  fm_shape s{};
  s.fields = FM_SHAPE_FIELD_NAME | FM_SHAPE_FIELD_PRST | FM_SHAPE_FIELD_ANCHOR | FM_SHAPE_FIELD_TEXT |
             FM_SHAPE_FIELD_FILL | FM_SHAPE_FIELD_LINE | FM_SHAPE_FIELD_LINE_WIDTH | FM_SHAPE_FIELD_ROTATION;
  s.name = "Callout";
  s.prst = "wedgeRectCallout";
  s.text = "Hello\nWorld";
  s.anchor_kind = FM_ANCHOR_KIND_TWO_CELL;
  s.edit_as = FM_ANCHOR_EDIT_AS_ONE_CELL;
  s.from_row = 2;
  s.from_col = 1;
  s.to_row = 6;
  s.to_col = 4;
  s.fill_kind = FM_SHAPE_PAINT_SOLID;
  s.fill_rgb = 0xFFCC00;
  s.line_kind = FM_SHAPE_PAINT_NONE;
  s.line_width_emu = 12700;
  s.rotation = 300000;
  return s;
}

TEST(FormulonCApiShapes, EnumOrdinalsAndFieldBits) {
  EXPECT_EQ(FM_SHAPE_PAINT_UNSET, 0);
  EXPECT_EQ(FM_SHAPE_PAINT_SOLID, 1);
  EXPECT_EQ(FM_SHAPE_PAINT_NONE, 2);
  EXPECT_EQ(FM_SHAPE_PAINT_OTHER, 3);
  EXPECT_EQ(FM_SHAPE_FIELD_HIDDEN, 0x200u);
  EXPECT_EQ(FM_SPARKLINE_STACKED, 2);
  EXPECT_EQ(FM_SPARKLINE_AXIS_CUSTOM, 2);
  EXPECT_EQ(FM_SPARKLINE_EMPTY_SPAN, 2);
}

TEST(FormulonCApiShapes, InsertListUpdateRemoveAndReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  size_t count = 99;
  ASSERT_EQ(fm_sheet_shape_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  const fm_shape spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_insert_shape(wb.handle, 0, &spec, &id), 0) << fm_last_error_message();
  EXPECT_EQ(id, 2U);

  fm_shape connector{};
  connector.fields = FM_SHAPE_FIELD_PRST | FM_SHAPE_FIELD_ANCHOR | FM_SHAPE_FIELD_LINE;
  connector.prst = "straightConnector1";
  connector.anchor_kind = FM_ANCHOR_KIND_ONE_CELL;
  connector.from_row = 10;
  connector.from_col_off = 500;
  connector.cx = 914400;
  connector.cy = 0;
  connector.line_kind = FM_SHAPE_PAINT_SOLID;
  connector.line_rgb = 0x00FF00;
  uint32_t line_id = 0;
  ASSERT_EQ(fm_sheet_insert_shape(wb.handle, 0, &connector, &line_id), 0) << fm_last_error_message();

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  ASSERT_EQ(fm_sheet_shape_count(again.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  fm_shape out{};
  ASSERT_EQ(fm_sheet_shape_at(again.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.object_id, id);
  EXPECT_EQ(out.kind, FM_DRAWING_OBJECT_SHAPE);
  EXPECT_STREQ(out.name, "Callout");
  EXPECT_STREQ(out.prst, "wedgeRectCallout");
  EXPECT_STREQ(out.text, "Hello\nWorld");
  EXPECT_EQ(out.anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(out.edit_as, FM_ANCHOR_EDIT_AS_ONE_CELL);
  EXPECT_EQ(out.from_row, 2U);
  EXPECT_EQ(out.to_col, 4U);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_SOLID);
  EXPECT_EQ(out.fill_rgb, 0xFFCC00U);
  EXPECT_EQ(out.line_kind, FM_SHAPE_PAINT_NONE);
  EXPECT_EQ(out.line_width_emu, 12700);
  EXPECT_EQ(out.rotation, 300000);
  EXPECT_EQ(out.fields, 0U);
  ASSERT_EQ(fm_sheet_shape_at(again.handle, 0, 1, &out), 0);
  EXPECT_EQ(out.kind, FM_DRAWING_OBJECT_CONNECTOR);
  EXPECT_EQ(out.from_col_off, 500);
  EXPECT_EQ(out.cx, 914400);
  EXPECT_EQ(out.line_rgb, 0x00FF00U);

  fm_shape patch{};
  patch.fields = FM_SHAPE_FIELD_NAME | FM_SHAPE_FIELD_HIDDEN | FM_SHAPE_FIELD_FILL;
  patch.name = "Renamed";
  patch.hidden = 1;
  patch.fill_kind = FM_SHAPE_PAINT_NONE;
  ASSERT_EQ(fm_sheet_update_shape(again.handle, 0, id, &patch), 0) << fm_last_error_message();
  ASSERT_EQ(fm_sheet_shape_at(again.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Renamed");
  EXPECT_EQ(out.hidden, 1);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_NONE);
  EXPECT_STREQ(out.prst, "wedgeRectCallout");  // unnamed fields stay
  EXPECT_EQ(out.rotation, 300000);

  ASSERT_EQ(fm_sheet_remove_shape(again.handle, 0, id), 0);
  ASSERT_EQ(fm_sheet_shape_count(again.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  EXPECT_EQ(fm_sheet_remove_shape(again.handle, 0, id), kInvalidArgument);
}

TEST(FormulonCApiShapes, Validation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t id = 0;
  fm_shape spec = Callout();
  EXPECT_EQ(fm_sheet_insert_shape(wb.handle, 0, nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_insert_shape(wb.handle, 0, &spec, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_insert_shape(wb.handle, 7, &spec, &id), kInvalidArgument);
  spec.anchor_kind = 9;
  EXPECT_EQ(fm_sheet_insert_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.fields = FM_SHAPE_FIELD_PRST;  // no anchor
  EXPECT_EQ(fm_sheet_insert_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.fields |= 0x80000000u;
  EXPECT_EQ(fm_sheet_insert_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.fill_kind = FM_SHAPE_PAINT_OTHER;
  EXPECT_EQ(fm_sheet_insert_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.to_row = 0;  // two-cell box ends before it starts
  EXPECT_EQ(fm_sheet_insert_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  size_t count = 5;
  ASSERT_EQ(fm_sheet_shape_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  fm_shape out{};
  EXPECT_EQ(fm_sheet_shape_at(wb.handle, 0, 0, &out), kInvalidArgument);
  EXPECT_EQ(fm_sheet_shape_at(wb.handle, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_shape_count(wb.handle, 0, nullptr), kBindingNullPointer);
}

fm_sparkline_group ColumnGroup() {
  fm_sparkline_group g{};
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  size_t count = 9;
  ASSERT_EQ(fm_sheet_sparkline_group_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  const fm_sparkline items[] = {{"Sheet1!A1:D1", "F1"}, {"Sheet1!A2:D2", "F2"}};
  fm_sparkline_group g = ColumnGroup();
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 2), 0) << fm_last_error_message();
  g.type = FM_SPARKLINE_LINE;
  g.markers = 0;
  g.has_manual_min = 0;
  g.color_series = fm_color_spec{};
  const fm_sparkline more[] = {{"Sheet1!A9:D9", "G9"}};
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, more, 1), 0) << fm_last_error_message();

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  ASSERT_EQ(fm_sheet_sparkline_group_count(again.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  fm_sparkline_group out{};
  ASSERT_EQ(fm_sheet_sparkline_group_at(again.handle, 0, 0, &out), 0);
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
  fm_sparkline sp{};
  ASSERT_EQ(fm_sheet_sparkline_at(again.handle, 0, 0, 1, &sp), 0);
  EXPECT_STREQ(sp.formula, "Sheet1!A2:D2");
  EXPECT_STREQ(sp.location, "F2");
  EXPECT_EQ(fm_sheet_sparkline_at(again.handle, 0, 0, 2, &sp), kInvalidArgument);
  EXPECT_EQ(fm_sheet_sparkline_at(again.handle, 0, 5, 0, &sp), kInvalidArgument);

  // Update rewrites the group's modelled fields and keeps its sparklines.
  out.type = FM_SPARKLINE_STACKED;
  out.date_axis = 1;
  const char* date_ref = "Sheet1!A1:D1";
  out.date_axis_ref = date_ref;
  out.high = 1;
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, 0, &out), 0) << fm_last_error_message();
  fm_sparkline_group changed{};
  ASSERT_EQ(fm_sheet_sparkline_group_at(again.handle, 0, 0, &changed), 0);
  EXPECT_EQ(changed.type, FM_SPARKLINE_STACKED);
  EXPECT_EQ(changed.high, 1);
  EXPECT_EQ(changed.date_axis, 1);
  EXPECT_STREQ(changed.date_axis_ref, "Sheet1!A1:D1");
  EXPECT_EQ(changed.sparkline_count, 2U);

  const fm_sparkline replacement[] = {{"Sheet1!B1:B5", "H1"}};
  ASSERT_EQ(fm_sheet_set_sparkline_group_sparklines(again.handle, 0, 0, replacement, 1), 0);
  ASSERT_EQ(fm_sheet_sparkline_group_at(again.handle, 0, 0, &changed), 0);
  EXPECT_EQ(changed.sparkline_count, 1U);

  ASSERT_EQ(fm_sheet_remove_sparkline_group(again.handle, 0, 0), 0);
  ASSERT_EQ(fm_sheet_sparkline_group_count(again.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  ASSERT_EQ(fm_sheet_remove_sparkline_group(again.handle, 0, 0), 0);
  EXPECT_EQ(fm_sheet_remove_sparkline_group(again.handle, 0, 0), kInvalidArgument);
}

TEST(FormulonCApiSparklines, Validation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const fm_sparkline good[] = {{"Sheet1!A1:D1", "F1"}};
  fm_sparkline_group g = ColumnGroup();
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, nullptr, good, 1), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, nullptr, 1), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, good, 0), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 3, &g, good, 1), kInvalidArgument);
  const fm_sparkline range_location[] = {{"Sheet1!A1:D1", "F1:F2"}};
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, range_location, 1), kInvalidArgument);
  const fm_sparkline no_formula[] = {{"", "F1"}};
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, no_formula, 1), kInvalidArgument);
  fm_sparkline_group bad = g;
  bad.type = 7;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &bad, good, 1), kInvalidArgument);
  bad = g;
  bad.line_weight = -1.0;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &bad, good, 1), kInvalidArgument);
  size_t count = 5;
  ASSERT_EQ(fm_sheet_sparkline_group_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &g), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_sparkline_group_sparklines(wb.handle, 0, 0, good, 1), kInvalidArgument);
  fm_sparkline_group out{};
  EXPECT_EQ(fm_sheet_sparkline_group_at(wb.handle, 0, 0, &out), kInvalidArgument);
  EXPECT_EQ(fm_sheet_sparkline_group_at(wb.handle, 0, 0, nullptr), kBindingNullPointer);
}

}  // namespace
