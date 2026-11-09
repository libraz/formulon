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
#include "miniz.h"
#include "support/ooxml_package_fixture.h"
#include "utils/error.h"

namespace {

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_shape_info_t) == (kWasm32 ? 184U : 192U), "fm_shape_info_t ABI layout changed");
static_assert(sizeof(fm_shape_patch_t) == (kWasm32 ? 296U : 320U), "fm_shape_patch_t ABI layout changed");
static_assert(sizeof(fm_sparkline_group_patch_t) == (kWasm32 ? 432U : 448U),
              "fm_sparkline_group_patch_t ABI layout changed");
static_assert(sizeof(fm_sparkline_t) == 2U * sizeof(void*), "fm_sparkline_t ABI layout changed");
static_assert(sizeof(fm_sparkline_group_t) == (kWasm32 ? 296U : 304U), "fm_sparkline_group_t ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

void SaveAndReload(fm_workbook_t* wb, WorkbookGuard* out) {
  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb, &buf.data, &buf.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &out->handle), 0) << fm_get_last_error_message();
}

template <typename T, typename V>
void Set(fm_patch_op_t& op, T& field, V value) {
  op = FM_PATCH_OP_SET;
  field = value;
}

fm_color_t Rgb(uint32_t argb) {
  fm_color_t c{};
  c.kind = FM_COLOR_KIND_RGB;
  c.argb = argb;
  return c;
}

fm_color_t Theme(uint32_t theme, double tint) {
  fm_color_t c{};
  c.kind = FM_COLOR_KIND_THEME;
  c.theme = theme;
  c.tint = tint;
  return c;
}

// A two-cell callout with text, a solid fill, no outline, a line width and a rotation.
fm_shape_patch_t Callout() {
  fm_shape_patch_t s = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(s.name_op, s.name, "Callout");
  Set(s.preset_geometry_op, s.preset_geometry, "wedgeRectCallout");
  Set(s.text_op, s.text, "Hello\nWorld");
  Set(s.anchor_kind_op, s.anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
  Set(s.edit_as_op, s.edit_as, FM_ANCHOR_EDIT_AS_ONE_CELL);
  Set(s.from_row_op, s.from_row, 2U);
  Set(s.from_col_op, s.from_col, 1U);
  Set(s.to_row_op, s.to_row, 6U);
  Set(s.to_col_op, s.to_col, 4U);
  Set(s.fill_kind_op, s.fill_kind, FM_SHAPE_PAINT_SOLID);
  Set(s.fill_color_op, s.fill_color, Rgb(0xFFFFCC00U));
  Set(s.line_kind_op, s.line_kind, FM_SHAPE_PAINT_NONE);
  Set(s.line_width_pt_op, s.line_width_pt, 1.0);
  Set(s.rotation_deg_op, s.rotation_deg, 5.0);
  return s;
}

fm_shape_info_t ShapeAt(const fm_workbook_t* wb, size_t index) {
  fm_shape_info_t out = FM_STRUCT_INIT(fm_shape_info_t);
  EXPECT_EQ(fm_sheet_get_shape_at(wb, 0, index, &out), 0) << fm_get_last_error_message();
  return out;
}

TEST(FormulonCApiShapes, EnumOrdinals) {
  EXPECT_EQ(FM_SHAPE_PAINT_ABSENT, -1);
  EXPECT_EQ(FM_SHAPE_PAINT_SOLID, 0);
  EXPECT_EQ(FM_SHAPE_PAINT_NONE, 1);
  EXPECT_EQ(FM_SHAPE_PAINT_OTHER, 2);
  EXPECT_EQ(FM_SPARKLINE_TYPE_STACKED, 2);
  EXPECT_EQ(FM_SPARKLINE_AXIS_TYPE_CUSTOM, 2);
}

TEST(FormulonCApiShapes, InsertListUpdateRemoveAndReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  size_t count = 99;
  ASSERT_EQ(fm_sheet_get_shape_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  const fm_shape_patch_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(id, 2U);

  fm_shape_patch_t connector = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(connector.preset_geometry_op, connector.preset_geometry, "straightConnector1");
  Set(connector.anchor_kind_op, connector.anchor_kind, FM_ANCHOR_KIND_ONE_CELL);
  Set(connector.from_row_op, connector.from_row, 10U);
  Set(connector.from_col_offset_emu_op, connector.from_col_offset_emu, 500);
  Set(connector.width_emu_op, connector.width_emu, 914400);
  Set(connector.line_kind_op, connector.line_kind, FM_SHAPE_PAINT_SOLID);
  Set(connector.line_color_op, connector.line_color, Rgb(0xFF00FF00U));
  uint32_t line_id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &connector, &line_id), 0) << fm_get_last_error_message();

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  ASSERT_EQ(fm_sheet_get_shape_count(again.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  fm_shape_info_t out = ShapeAt(again.handle, 0);
  EXPECT_EQ(out.object_id, id);
  EXPECT_EQ(out.kind, FM_DRAWING_OBJECT_KIND_SHAPE);
  EXPECT_STREQ(out.name, "Callout");
  EXPECT_STREQ(out.preset_geometry, "wedgeRectCallout");
  EXPECT_STREQ(out.text, "Hello\nWorld");
  EXPECT_EQ(out.anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(out.edit_as, FM_ANCHOR_EDIT_AS_ONE_CELL);
  EXPECT_EQ(out.from_row, 2U);
  EXPECT_EQ(out.to_col, 4U);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_SOLID);
  EXPECT_EQ(out.fill_color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(out.fill_color.argb, 0xFFFFCC00U);
  EXPECT_EQ(out.line_kind, FM_SHAPE_PAINT_NONE);
  EXPECT_DOUBLE_EQ(out.line_width_pt, 1.0);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 5.0);
  out = ShapeAt(again.handle, 1);
  EXPECT_EQ(out.kind, FM_DRAWING_OBJECT_KIND_CONNECTOR);
  EXPECT_EQ(out.from_col_offset_emu, 500);
  EXPECT_EQ(out.width_emu, 914400);
  EXPECT_EQ(out.line_color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(out.line_color.argb, 0xFF00FF00U);

  fm_shape_patch_t patch = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(patch.name_op, patch.name, "Renamed");
  Set(patch.hidden_op, patch.hidden, 1);
  Set(patch.fill_kind_op, patch.fill_kind, FM_SHAPE_PAINT_NONE);
  ASSERT_EQ(fm_sheet_update_shape(again.handle, 0, id, &patch), 0) << fm_get_last_error_message();
  out = ShapeAt(again.handle, 0);
  EXPECT_STREQ(out.name, "Renamed");
  EXPECT_EQ(out.hidden, 1);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_NONE);
  EXPECT_STREQ(out.preset_geometry, "wedgeRectCallout");  // KEEP fields stay
  EXPECT_DOUBLE_EQ(out.rotation_deg, 5.0);

  EXPECT_EQ(fm_sheet_update_shape(again.handle, 0, 99, &patch), FM_STATUS_NOT_FOUND);
  ASSERT_EQ(fm_sheet_remove_shape(again.handle, 0, id), 0);
  ASSERT_EQ(fm_sheet_get_shape_count(again.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  EXPECT_EQ(fm_sheet_remove_shape(again.handle, 0, id), FM_STATUS_NOT_FOUND);
}

TEST(FormulonCApiShapes, RotationDegreesAndLineWidthPointsSurviveReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_patch_t spec = Callout();
  spec.rotation_deg = 12.5;
  Set(spec.line_kind_op, spec.line_kind, FM_SHAPE_PAINT_SOLID);
  Set(spec.line_color_op, spec.line_color, Rgb(0xFF102030U));
  spec.line_width_pt = 2.25;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  fm_shape_info_t out = ShapeAt(wb.handle, 0);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 12.5);
  EXPECT_DOUBLE_EQ(out.line_width_pt, 2.25);

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  out = ShapeAt(again.handle, 0);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 12.5);
  EXPECT_DOUBLE_EQ(out.line_width_pt, 2.25);

  // A patch that sets only the rotation turns the shape and leaves the width.
  fm_shape_patch_t patch = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(patch.rotation_deg_op, patch.rotation_deg, 90.0);
  ASSERT_EQ(fm_sheet_update_shape(again.handle, 0, id, &patch), 0) << fm_get_last_error_message();
  out = ShapeAt(again.handle, 0);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 90.0);
  EXPECT_DOUBLE_EQ(out.line_width_pt, 2.25);
}

TEST(FormulonCApiShapes, RotationIsNormalisedAndRangeChecked) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_patch_t spec = Callout();
  spec.rotation_deg = -90.0;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_DOUBLE_EQ(ShapeAt(wb.handle, 0).rotation_deg, 270.0);

  spec.rotation_deg = 450.0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_DOUBLE_EQ(ShapeAt(wb.handle, 1).rotation_deg, 90.0);

  for (const double bad :
       {3601.0, -3601.0, std::numeric_limits<double>::infinity(), std::numeric_limits<double>::quiet_NaN()}) {
    spec.rotation_deg = bad;
    EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument) << bad;
  }
}

TEST(FormulonCApiShapes, ThemeColorsRoundTripAndAreValidated) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_patch_t spec = Callout();
  spec.fill_color = Theme(4, 0.4);
  Set(spec.line_kind_op, spec.line_kind, FM_SHAPE_PAINT_SOLID);
  Set(spec.line_color_op, spec.line_color, Theme(1, -0.25));
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  const fm_shape_info_t out = ShapeAt(wb.handle, 0);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_SOLID);
  EXPECT_EQ(out.fill_color.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(out.fill_color.theme, 4U);
  EXPECT_DOUBLE_EQ(out.fill_color.tint, 0.4);
  EXPECT_EQ(out.line_color.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(out.line_color.theme, 1U);
  EXPECT_DOUBLE_EQ(out.line_color.tint, -0.25);

  struct Bad {
    int32_t kind;
    uint32_t theme;
    double tint;
  };
  const Bad bad[] = {
      {FM_COLOR_KIND_INDEXED, 0, 0.0}, {FM_COLOR_KIND_AUTO, 0, 0.0},
      {FM_COLOR_KIND_THEME, 12, 0.0},  {FM_COLOR_KIND_THEME, 4, 1.5},
      {FM_COLOR_KIND_THEME, 4, -1.5},  {FM_COLOR_KIND_THEME, 4, std::numeric_limits<double>::quiet_NaN()},
  };
  for (const Bad& b : bad) {
    fm_shape_patch_t s = Callout();
    s.fill_color = fm_color_t{};
    s.fill_color.kind = b.kind;
    s.fill_color.theme = b.theme;
    s.fill_color.tint = b.tint;
    EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &s, &id), kInvalidArgument)
        << "kind=" << b.kind << " theme=" << b.theme << " tint=" << b.tint;
  }

  // A colour on a paint that is not solid is refused rather than dropped.
  fm_shape_patch_t none = Callout();
  none.fill_kind = FM_SHAPE_PAINT_NONE;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &none, &id), kInvalidArgument);
  none = Callout();
  Set(none.line_color_op, none.line_color, Rgb(0xFF000000U));
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &none, &id), kInvalidArgument);
  // A colour left KEEP is not read at all.
  none.line_color_op = FM_PATCH_OP_KEEP;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &none, &id), 0) << fm_get_last_error_message();
}

// CLEAR on a paint field removes the shape's own fill or outline so the style's
// paint applies; it reads ABSENT and survives a save and reload.
TEST(FormulonCApiShapes, PaintClearRemovesTheShapesOwnFillAndOutline) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_shape_patch_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  fm_shape_patch_t clear = FM_STRUCT_INIT(fm_shape_patch_t);
  clear.fill_kind_op = FM_PATCH_OP_CLEAR;
  clear.line_color_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, &clear), 0) << fm_get_last_error_message();
  fm_shape_info_t out = ShapeAt(wb.handle, 0);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_ABSENT);
  EXPECT_EQ(out.fill_color.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(out.line_kind, FM_SHAPE_PAINT_ABSENT);
  EXPECT_DOUBLE_EQ(out.line_width_pt, 1.0);  // the outline keeps its width

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  out = ShapeAt(again.handle, 0);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_ABSENT);
  EXPECT_EQ(out.line_kind, FM_SHAPE_PAINT_ABSENT);

  // CLEAR on one of a pair with SET on the other is refused, and nothing changes.
  fm_shape_patch_t mixed = FM_STRUCT_INIT(fm_shape_patch_t);
  mixed.fill_kind_op = FM_PATCH_OP_CLEAR;
  Set(mixed.fill_color_op, mixed.fill_color, Rgb(0xFF123456U));
  Set(mixed.hidden_op, mixed.hidden, 1);
  EXPECT_EQ(fm_sheet_update_shape(again.handle, 0, id, &mixed), kInvalidArgument);
  out = ShapeAt(again.handle, 0);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_ABSENT);
  EXPECT_EQ(out.hidden, 0);
}

// A paint pair combines with the stored paint: a colour alone makes the paint
// solid, and a solid kind alone keeps the stored colour.
TEST(FormulonCApiShapes, PaintPairCombinesWithTheStoredPaint) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_shape_patch_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  fm_shape_patch_t colour = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(colour.line_color_op, colour.line_color, Rgb(0xFF0000FFU));
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, &colour), 0) << fm_get_last_error_message();
  fm_shape_info_t out = ShapeAt(wb.handle, 0);
  EXPECT_EQ(out.line_kind, FM_SHAPE_PAINT_SOLID);
  EXPECT_EQ(out.line_color.argb, 0xFF0000FFU);

  fm_shape_patch_t kind = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(kind.fill_kind_op, kind.fill_kind, FM_SHAPE_PAINT_SOLID);
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, &kind), 0) << fm_get_last_error_message();
  out = ShapeAt(wb.handle, 0);
  EXPECT_EQ(out.fill_kind, FM_SHAPE_PAINT_SOLID);
  EXPECT_EQ(out.fill_color.argb, 0xFFFFCC00U);
}

// An update applies its anchor fields to the current anchor and cannot change
// its kind.
TEST(FormulonCApiShapes, UpdateMergesAnchorFieldsIntoTheCurrentAnchor) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_shape_patch_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  fm_shape_patch_t move = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(move.to_col_op, move.to_col, 7U);
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, &move), 0) << fm_get_last_error_message();
  fm_shape_info_t out = ShapeAt(wb.handle, 0);
  EXPECT_EQ(out.from_row, 2U);
  EXPECT_EQ(out.from_col, 1U);
  EXPECT_EQ(out.to_row, 6U);
  EXPECT_EQ(out.to_col, 7U);
  EXPECT_EQ(out.edit_as, FM_ANCHOR_EDIT_AS_ONE_CELL);

  fm_shape_patch_t same = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(same.anchor_kind_op, same.anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, &same), 0) << fm_get_last_error_message();
  fm_shape_patch_t other = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(other.anchor_kind_op, other.anchor_kind, FM_ANCHOR_KIND_ONE_CELL);
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, &other), kInvalidArgument);
  EXPECT_EQ(ShapeAt(wb.handle, 0).anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
}

TEST(FormulonCApiShapes, Validation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  uint32_t id = 0;
  fm_shape_patch_t spec = Callout();
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 7, &spec, &id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  spec.anchor_kind = 9;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.anchor_kind = FM_ANCHOR_KIND_ABSOLUTE;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.anchor_kind_op = FM_PATCH_OP_KEEP;  // no anchor
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.name_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.text = nullptr;  // a SET string cannot be NULL
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.name = "\xC3";
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  spec = Callout();
  spec.fill_kind = FM_SHAPE_PAINT_OTHER;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.to_row = 0;  // two-cell box ends before it starts
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  size_t count = 5;
  ASSERT_EQ(fm_sheet_get_shape_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  fm_shape_info_t out = FM_STRUCT_INIT(fm_shape_info_t);
  EXPECT_EQ(fm_sheet_get_shape_at(wb.handle, 0, 0, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_shape_at(wb.handle, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_shape_count(wb.handle, 0, nullptr), kBindingNullPointer);
}

fm_sparkline_group_t ColumnGroup() {
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  g.type = FM_SPARKLINE_TYPE_COLUMN;
  g.negative = 1;
  g.markers = 1;
  g.display_x_axis = 1;
  g.min_axis_type = FM_SPARKLINE_AXIS_TYPE_CUSTOM;
  g.has_manual_min = 1;
  g.manual_min = -1.5;
  g.line_weight = 1.25;
  g.empty_cells_as = FM_BLANKS_AS_GAP;
  g.color_series.kind = FM_COLOR_KIND_RGB;
  g.color_series.argb = 0xFF336699U;
  g.color_negative.kind = FM_COLOR_KIND_THEME;
  g.color_negative.theme = 5;
  g.color_negative.tint = -0.25;
  g.color_low.kind = FM_COLOR_KIND_INDEXED;
  g.color_low.indexed = 12;
  g.color_first.kind = FM_COLOR_KIND_AUTO;
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
  g.type = FM_SPARKLINE_TYPE_LINE;
  g.markers = 0;
  g.has_manual_min = 0;
  g.color_series = fm_color_t{};
  const fm_sparkline_t more[] = {{"Sheet1!A9:D9", "G9"}};
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, more, 1, &group_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(group_index, 1U);

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  ASSERT_EQ(fm_sheet_get_sparkline_group_count(again.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  fm_sparkline_group_t out = FM_STRUCT_INIT(fm_sparkline_group_t);
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(again.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.type, FM_SPARKLINE_TYPE_COLUMN);
  EXPECT_EQ(out.negative, 1);
  EXPECT_EQ(out.markers, 1);
  EXPECT_EQ(out.high, 0);
  EXPECT_EQ(out.display_x_axis, 1);
  EXPECT_EQ(out.min_axis_type, FM_SPARKLINE_AXIS_TYPE_CUSTOM);
  EXPECT_EQ(out.max_axis_type, FM_SPARKLINE_AXIS_TYPE_INDIVIDUAL);
  EXPECT_EQ(out.has_manual_min, 1);
  EXPECT_DOUBLE_EQ(out.manual_min, -1.5);
  EXPECT_EQ(out.has_manual_max, 0);
  EXPECT_DOUBLE_EQ(out.line_weight, 1.25);
  EXPECT_EQ(out.empty_cells_as, FM_BLANKS_AS_GAP);
  EXPECT_EQ(out.color_series.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(out.color_series.argb, 0xFF336699U);
  EXPECT_EQ(out.color_negative.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(out.color_negative.theme, 5U);
  EXPECT_DOUBLE_EQ(out.color_negative.tint, -0.25);
  EXPECT_EQ(out.color_low.indexed, 12U);
  EXPECT_EQ(out.color_first.kind, FM_COLOR_KIND_AUTO);
  EXPECT_EQ(out.color_high.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(out.date_axis_ref, nullptr);
  ASSERT_EQ(out.sparkline_count, 2U);
  fm_sparkline_t sp{};
  ASSERT_EQ(fm_sheet_get_sparkline_at(again.handle, 0, 0, 1, &sp), 0);
  EXPECT_STREQ(sp.formula, "=Sheet1!A2:D2");
  EXPECT_STREQ(sp.location, "F2");
  EXPECT_EQ(fm_sheet_get_sparkline_at(again.handle, 0, 0, 2, &sp), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_sparkline_at(again.handle, 0, 5, 0, &sp), FM_STATUS_INDEX_OUT_OF_RANGE);

  // A patch without items writes its SET fields and keeps the sparklines.
  fm_sparkline_group_patch_t patch = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  Set(patch.type_op, patch.type, FM_SPARKLINE_TYPE_STACKED);
  Set(patch.date_axis_op, patch.date_axis, 1);
  Set(patch.date_axis_ref_op, patch.date_axis_ref, "Sheet1!A1:D1");
  Set(patch.high_op, patch.high, 1);
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, 0, &patch), 0) << fm_get_last_error_message();
  fm_sparkline_group_t changed = FM_STRUCT_INIT(fm_sparkline_group_t);
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(again.handle, 0, 0, &changed), 0);
  EXPECT_EQ(changed.type, FM_SPARKLINE_TYPE_STACKED);
  EXPECT_EQ(changed.high, 1);
  EXPECT_EQ(changed.date_axis, 1);
  EXPECT_STREQ(changed.date_axis_ref, "=Sheet1!A1:D1");
  EXPECT_EQ(changed.sparkline_count, 2U);
  EXPECT_EQ(changed.negative, 1);  // KEEP fields stay
  EXPECT_DOUBLE_EQ(changed.manual_min, -1.5);

  // SET items replaces the sparklines and keeps every attribute.
  const fm_sparkline_t replacement[] = {{"Sheet1!B1:B5", "H1"}};
  fm_sparkline_group_patch_t replace = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  replace.items_op = FM_PATCH_OP_SET;
  replace.items = replacement;
  replace.items_count = 1;
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, 0, &replace), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(again.handle, 0, 0, &changed), 0);
  EXPECT_EQ(changed.sparkline_count, 1U);
  EXPECT_EQ(changed.type, FM_SPARKLINE_TYPE_STACKED);
  EXPECT_EQ(changed.high, 1);
  EXPECT_STREQ(changed.date_axis_ref, "=Sheet1!A1:D1");
  ASSERT_EQ(fm_sheet_get_sparkline_at(again.handle, 0, 0, 0, &sp), 0);
  EXPECT_STREQ(sp.formula, "=Sheet1!B1:B5");
  EXPECT_STREQ(sp.location, "H1");

  // CLEAR removes the date range.
  fm_sparkline_group_patch_t clear = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  clear.date_axis_ref_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, 0, &clear), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(again.handle, 0, 0, &changed), 0);
  EXPECT_EQ(changed.date_axis_ref, nullptr);

  // CLEAR removes a colour, which then reads as no colour.
  fm_sparkline_group_patch_t uncolour = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  uncolour.color_high_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, 0, &uncolour), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(again.handle, 0, 0, &changed), 0);
  EXPECT_EQ(changed.color_high.kind, FM_COLOR_KIND_NONE);

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
  // A group always states how it draws empty cells.
  bad = g;
  bad.empty_cells_as = FM_BLANKS_AS_ABSENT;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &bad, good, 1, ignored_index()), kInvalidArgument);
  bad.empty_cells_as = FM_BLANKS_AS_SPAN + 1;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &bad, good, 1, ignored_index()), kInvalidArgument);
  size_t count = 5;
  ASSERT_EQ(fm_sheet_get_sparkline_group_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  const fm_sparkline_group_patch_t keep = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &keep), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, nullptr), kBindingNullPointer);
  fm_sparkline_group_t out = FM_STRUCT_INIT(fm_sparkline_group_t);
  EXPECT_EQ(fm_sheet_get_sparkline_group_at(wb.handle, 0, 0, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_sparkline_group_at(wb.handle, 0, 0, nullptr), kBindingNullPointer);
}

TEST(FormulonCApiSparklines, UpdateWithItemsIsAllOrNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}, {"Sheet1!A2:D2", "F2"}};
  fm_sparkline_group_t g = ColumnGroup();
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 2, ignored_index()), 0)
      << fm_get_last_error_message();

  // A group cannot hold zero sparklines (CT_Sparklines requires one), so an
  // empty SET items is rejected like an invalid item, and no field lands.
  fm_sparkline_group_patch_t patch = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  Set(patch.type_op, patch.type, FM_SPARKLINE_TYPE_LINE);
  patch.items_op = FM_PATCH_OP_SET;
  const fm_sparkline_t none[] = {{"Sheet1!A1:D1", "F1"}};
  patch.items = none;
  patch.items_count = 0;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &patch), kInvalidArgument);
  const fm_sparkline_t bad[] = {{"Sheet1!A1:D1", "F1:F2"}};
  patch.items = bad;
  patch.items_count = 1;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &patch), kInvalidArgument);
  patch.items = nullptr;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &patch), kBindingNullPointer);
  // An invalid field with valid items does not replace the items either.
  const fm_sparkline_t one[] = {{"Sheet1!B1:B5", "H1"}};
  patch.items = one;
  Set(patch.line_weight_op, patch.line_weight, -1.0);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &patch), kInvalidArgument);
  // Out-of-domain values and ops are refused before anything is written.
  fm_sparkline_group_patch_t bad_field = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  Set(bad_field.type_op, bad_field.type, 7);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &bad_field), kInvalidArgument);
  bad_field = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  bad_field.markers_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &bad_field), kInvalidArgument);
  bad_field = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  Set(bad_field.date_axis_ref_op, bad_field.date_axis_ref, nullptr);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &bad_field), kInvalidArgument);
  bad_field = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  Set(bad_field.empty_cells_as_op, bad_field.empty_cells_as, FM_BLANKS_AS_ABSENT);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 0, &bad_field), kInvalidArgument);

  fm_sparkline_group_t read = FM_STRUCT_INIT(fm_sparkline_group_t);
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(wb.handle, 0, 0, &read), 0);
  EXPECT_EQ(read.type, FM_SPARKLINE_TYPE_COLUMN);
  EXPECT_EQ(read.sparkline_count, 2U);
}

// An out-of-range position names the parameter it came from in the error context.
TEST(FormulonCApiSparklines, OutOfRangeContextNamesTheParameter) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}};
  fm_sparkline_group_t g = ColumnGroup();
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, ignored_index()), 0)
      << fm_get_last_error_message();
  const auto context_has = [](const char* key) {
    return std::string(fm_get_last_error_context()).find(key) != std::string::npos;
  };
  fm_sparkline_group_t out = FM_STRUCT_INIT(fm_sparkline_group_t);
  EXPECT_EQ(fm_sheet_get_sparkline_group_at(wb.handle, 0, 3, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_TRUE(context_has("group_index=3"));
  fm_sparkline_t sp{};
  EXPECT_EQ(fm_sheet_get_sparkline_at(wb.handle, 0, 2, 0, &sp), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_TRUE(context_has("group_index=2"));
  EXPECT_EQ(fm_sheet_get_sparkline_at(wb.handle, 0, 0, 4, &sp), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_TRUE(context_has("sparkline_index=4"));
  const fm_sparkline_group_patch_t keep = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, 5, &keep), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_TRUE(context_has("group_index=5"));
  EXPECT_EQ(fm_sheet_remove_sparkline_group(wb.handle, 0, 6), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_TRUE(context_has("group_index=6"));
}

std::string SavedSheetXml(fm_workbook_t* wb) {
  BufferGuard buf;
  EXPECT_EQ(fm_workbook_save(wb, &buf.data, &buf.len), 0) << fm_get_last_error_message();
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, buf.data, buf.len, 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return {};
  }
  std::size_t size = 0;
  void* bytes = mz_zip_reader_extract_file_to_heap(&reader, "xl/worksheets/sheet1.xml", &size, 0);
  mz_zip_reader_end(&reader);
  if (bytes == nullptr) {
    ADD_FAILURE() << "xl/worksheets/sheet1.xml not found";
    return {};
  }
  std::string xml(static_cast<const char*>(bytes), size);
  mz_free(bytes);
  return xml;
}

std::string SparklineSheetXml(const char* empty_cells_as) {
  return std::string(
             "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
             "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
             "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><sheetData/><extLst>"
             "<ext uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\"><x14:sparklineGroups>"
             "<x14:sparklineGroup displayEmptyCellsAs=\"") +
         empty_cells_as +
         "\"><x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:D1</xm:f><xm:sqref>F1</xm:sqref>"
         "</x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst></worksheet>";
}

// The sparkline group shares the chart's blanks table: each `displayEmptyCellsAs`
// token reads as its member, and each member writes its token.
TEST(FormulonCApiSparklines, EmptyCellsAsMapsEachTokenToTheSharedBlanksMember) {
  const struct {
    int32_t blanks_as;
    const char* token;
  } kCases[] = {{FM_BLANKS_AS_ZERO, "zero"}, {FM_BLANKS_AS_GAP, "gap"}, {FM_BLANKS_AS_SPAN, "span"}};
  for (const auto& c : kCases) {
    const std::vector<std::uint8_t> package = formulon::test::BuildOoxmlPackage(
        formulon::test::OoxmlContentTypes(formulon::test::kXlsxWorkbookContentType), SparklineSheetXml(c.token));
    WorkbookGuard loaded;
    ASSERT_EQ(fm_workbook_load(package.data(), package.size(), &loaded.handle), 0) << fm_get_last_error_message();
    fm_sparkline_group_t read = FM_STRUCT_INIT(fm_sparkline_group_t);
    ASSERT_EQ(fm_sheet_get_sparkline_group_at(loaded.handle, 0, 0, &read), 0) << c.token;
    EXPECT_EQ(read.empty_cells_as, c.blanks_as) << c.token;

    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}};
    fm_sparkline_group_t g = ColumnGroup();
    g.empty_cells_as = c.blanks_as;
    ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, ignored_index()), 0)
        << fm_get_last_error_message();
    // Zero is the schema default, which the writer leaves implicit.
    const bool written =
        SavedSheetXml(wb.handle).find(std::string("displayEmptyCellsAs=\"") + c.token + "\"") != std::string::npos;
    EXPECT_EQ(written, c.blanks_as != FM_BLANKS_AS_ZERO) << c.token;
  }
}

// A sparkline formula and a date-axis range read with a leading `=`; input
// takes either spelling and stores the same range.
TEST(FormulonCApiSparklines, FormulasReadWithALeadingEqualsAndTakeEitherSpelling) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"=Sheet1!A1:D1", "F1"}, {"Sheet1!A2:D2", "F2"}};
  fm_sparkline_group_t g = ColumnGroup();
  g.date_axis = 1;
  g.date_axis_ref = "=Sheet1!A5:D5";
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 2, ignored_index()), 0)
      << fm_get_last_error_message();
  g.date_axis_ref = "Sheet1!A6:D6";
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 2, ignored_index()), 0)
      << fm_get_last_error_message();

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  for (fm_workbook_t* book : {wb.handle, again.handle}) {
    fm_sparkline_t sp{};
    ASSERT_EQ(fm_sheet_get_sparkline_at(book, 0, 0, 0, &sp), 0);
    EXPECT_STREQ(sp.formula, "=Sheet1!A1:D1");
    ASSERT_EQ(fm_sheet_get_sparkline_at(book, 0, 0, 1, &sp), 0);
    EXPECT_STREQ(sp.formula, "=Sheet1!A2:D2");
    fm_sparkline_group_t read = FM_STRUCT_INIT(fm_sparkline_group_t);
    ASSERT_EQ(fm_sheet_get_sparkline_group_at(book, 0, 0, &read), 0);
    EXPECT_STREQ(read.date_axis_ref, "=Sheet1!A5:D5");
    ASSERT_EQ(fm_sheet_get_sparkline_group_at(book, 0, 1, &read), 0);
    EXPECT_STREQ(read.date_axis_ref, "=Sheet1!A6:D6");
  }
  // The file keeps the `xm:f` spelling without `=`.
  const std::string xml = SavedSheetXml(wb.handle);
  EXPECT_NE(xml.find("<xm:f>Sheet1!A1:D1</xm:f>"), std::string::npos);
  EXPECT_EQ(xml.find("<xm:f>=Sheet1!"), std::string::npos);
}

}  // namespace
