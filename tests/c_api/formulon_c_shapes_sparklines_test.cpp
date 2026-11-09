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
static_assert(sizeof(fm_shape_info_t) == (kWasm32 ? 160U : 168U), "fm_shape_info_t ABI layout changed");
static_assert(sizeof(fm_shape_spec_t) == (kWasm32 ? 240U : 264U), "fm_shape_spec_t ABI layout changed");
static_assert(sizeof(fm_shape_patch_t) == (kWasm32 ? 184U : 200U), "fm_shape_patch_t ABI layout changed");
static_assert(sizeof(fm_sparkline_group_patch_t) == (kWasm32 ? 432U : 448U),
              "fm_sparkline_group_patch_t ABI layout changed");
static_assert(sizeof(fm_sparkline_t) == 2U * sizeof(void*), "fm_sparkline_t ABI layout changed");
static_assert(sizeof(fm_sparkline_group_t) == (kWasm32 ? 296U : 312U), "fm_sparkline_group_t ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

void SaveAndReload(fm_workbook_t* wb, WorkbookGuard* out) {
  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &buf.data, &buf.len), 0) << fm_get_last_error_message();
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

// A two-cell callout with text, a solid fill, no outline and a rotation.
fm_shape_spec_t Callout() {
  fm_shape_spec_t s = FM_STRUCT_INIT(fm_shape_spec_t);
  s.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  s.placement.name = "Callout";
  s.placement.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  s.placement.anchor.edit_as = FM_ANCHOR_EDIT_AS_ONE_CELL;
  s.placement.anchor.from_row = 2;
  s.placement.anchor.from_col = 1;
  s.placement.anchor.to_row = 6;
  s.placement.anchor.to_col = 4;
  s.preset_geometry = "wedgeRectCallout";
  s.text = "Hello\nWorld";
  s.fill.color = Rgb(0xFFFFCC00U);
  s.line.none = 1;
  s.rotation_deg = 5.0;
  return s;
}

// The shape fields of the id'th object of the first sheet's drawing.
fm_shape_info_t ShapeById(const fm_workbook_t* wb, std::uint32_t id) {
  fm_shape_info_t out = FM_STRUCT_INIT(fm_shape_info_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_shape(wb, 0, id, &present, &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return out;
}

// The common drawing fields (name, anchor, visibility) of object `id`.
fm_drawing_object_info_t ObjectById(const fm_workbook_t* wb, std::uint32_t id) {
  fm_drawing_object_info_t out = FM_STRUCT_INIT(fm_drawing_object_info_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_drawing_object(wb, 0, id, &present, &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return out;
}

// The id of the object at `index` of the first sheet's drawing.
std::uint32_t IdAt(const fm_workbook_t* wb, size_t index) {
  fm_drawing_object_info_t out = FM_STRUCT_INIT(fm_drawing_object_info_t);
  EXPECT_EQ(fm_sheet_get_drawing_object_at(wb, 0, index, &out), 0) << fm_get_last_error_message();
  return out.object_id;
}

size_t ObjectCount(const fm_workbook_t* wb) {
  size_t count = 99;
  EXPECT_EQ(fm_sheet_get_drawing_object_count(wb, 0, &count), 0);
  return count;
}

TEST(FormulonCApiShapes, EnumOrdinals) {
  EXPECT_EQ(FM_SPARKLINE_TYPE_STACKED, 2);
  EXPECT_EQ(FM_SPARKLINE_AXIS_TYPE_CUSTOM, 2);
}

TEST(FormulonCApiShapes, InsertReadUpdateRemoveAndReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(ObjectCount(wb.handle), 0U);

  const fm_shape_spec_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(id, 2U);

  fm_shape_spec_t connector = FM_STRUCT_INIT(fm_shape_spec_t);
  connector.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  connector.preset_geometry = "straightConnector1";
  connector.placement.anchor.kind = FM_ANCHOR_KIND_ONE_CELL;
  connector.placement.anchor.from_row = 10;
  connector.placement.anchor.from_col_offset_emu = 500;
  connector.placement.anchor.width_emu = 914400;
  connector.placement.anchor.height_emu = 914400;
  connector.line.color = Rgb(0xFF00FF00U);
  uint32_t line_id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &connector, &line_id), 0) << fm_get_last_error_message();

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  ASSERT_EQ(ObjectCount(again.handle), 2U);
  ASSERT_EQ(IdAt(again.handle, 0), id);
  fm_shape_info_t out = ShapeById(again.handle, id);
  EXPECT_EQ(out.object_id, id);
  EXPECT_EQ(out.kind, FM_DRAWING_OBJECT_KIND_SHAPE);
  EXPECT_STREQ(out.preset_geometry, "wedgeRectCallout");
  EXPECT_STREQ(out.text, "Hello\nWorld");
  EXPECT_EQ(out.fill.color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(out.fill.color.argb, 0xFFFFCC00U);
  EXPECT_EQ(out.fill.none, 0);
  EXPECT_EQ(out.line.none, 1);
  EXPECT_EQ(out.line.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 5.0);
  fm_drawing_object_info_t frame = ObjectById(again.handle, id);
  EXPECT_STREQ(frame.name, "Callout");
  EXPECT_EQ(frame.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(frame.anchor.edit_as, FM_ANCHOR_EDIT_AS_ONE_CELL);
  EXPECT_EQ(frame.anchor.from_row, 2U);
  EXPECT_EQ(frame.anchor.to_col, 4U);
  out = ShapeById(again.handle, line_id);
  EXPECT_EQ(out.kind, FM_DRAWING_OBJECT_KIND_CONNECTOR);
  EXPECT_EQ(out.line.color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(out.line.color.argb, 0xFF00FF00U);
  frame = ObjectById(again.handle, line_id);
  EXPECT_EQ(frame.anchor.from_col_offset_emu, 500);
  EXPECT_EQ(frame.anchor.width_emu, 914400);

  // The name and visibility (the common half) and the paint (the shape half) change in one call.
  fm_drawing_patch_t common = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(common.name_op, common.name, "Renamed");
  Set(common.hidden_op, common.hidden, 1);
  fm_shape_patch_t patch = FM_STRUCT_INIT(fm_shape_patch_t);
  patch.fill.color_op = FM_PATCH_OP_CLEAR;
  Set(patch.fill.none_op, patch.fill.none, 1);
  ASSERT_EQ(fm_sheet_update_shape(again.handle, 0, id, &common, &patch), 0) << fm_get_last_error_message();
  out = ShapeById(again.handle, id);
  // Strings are read-scratch views: compare the shape's before the frame read replaces them.
  EXPECT_STREQ(out.preset_geometry, "wedgeRectCallout");  // KEEP fields stay
  frame = ObjectById(again.handle, id);
  EXPECT_STREQ(frame.name, "Renamed");
  EXPECT_EQ(frame.hidden, 1);
  EXPECT_EQ(out.fill.none, 1);
  EXPECT_EQ(out.fill.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 5.0);

  EXPECT_EQ(fm_sheet_update_shape(again.handle, 0, 99, nullptr, &patch), FM_STATUS_NOT_FOUND);
  ASSERT_EQ(fm_sheet_remove_drawing(again.handle, 0, id), 0) << fm_get_last_error_message();
  EXPECT_EQ(ObjectCount(again.handle), 1U);
  EXPECT_EQ(fm_sheet_remove_drawing(again.handle, 0, id), FM_STATUS_NOT_FOUND);
  int32_t present = 1;
  fm_shape_info_t gone = FM_STRUCT_INIT(fm_shape_info_t);
  ASSERT_EQ(fm_sheet_get_shape(again.handle, 0, id, &present, &gone), 0);
  EXPECT_EQ(present, 0);
}

TEST(FormulonCApiShapes, PlacementDescriptionAndHiddenReachTheShape) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t spec = Callout();
  spec.placement.description = "a callout";
  spec.placement.hidden = FM_TRISTATE_TRUE;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t frame = ObjectById(wb.handle, id);
  EXPECT_STREQ(frame.description, "a callout");
  EXPECT_EQ(frame.hidden, 1);

  // A default placement takes Excel's defaults: a two-cell anchor, the generated name, shown.
  fm_shape_spec_t plain = FM_STRUCT_INIT(fm_shape_spec_t);
  plain.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  plain.placement.anchor.from_row = 1;
  plain.placement.anchor.to_row = 3;
  plain.placement.anchor.to_col = 2;
  uint32_t plain_id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &plain, &plain_id), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t made = ObjectById(wb.handle, plain_id);
  EXPECT_EQ(made.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(made.hidden, 0);
  EXPECT_EQ(std::string(made.name), "Shape " + std::to_string(plain_id));
  const fm_shape_info_t shape = ShapeById(wb.handle, plain_id);
  EXPECT_STREQ(shape.preset_geometry, "rect");
  // All-zero paint records: the shape draws its style's fill and outline.
  EXPECT_EQ(shape.fill.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(shape.fill.none, 0);
  EXPECT_EQ(shape.fill.other, 0);
  EXPECT_EQ(shape.line.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(shape.line.none, 0);
  EXPECT_EQ(shape.line.has_width_pt, 0);
}

TEST(FormulonCApiShapes, AStraightConnectorMayBeFlatButNotSizeless) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t line = FM_STRUCT_INIT(fm_shape_spec_t);
  line.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  line.preset_geometry = "straightConnector1";
  line.placement.anchor.kind = FM_ANCHOR_KIND_ONE_CELL;
  line.placement.anchor.from_row = 2;
  line.placement.anchor.width_emu = 914400;
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &line, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(ObjectById(wb.handle, id).anchor.height_emu, 0);

  line.placement.anchor.width_emu = 0;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &line, &id), FM_STATUS_INVALID_ARGUMENT);
}

TEST(FormulonCApiShapes, RotationDegreesAndLineWidthPointsSurviveReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t spec = Callout();
  spec.rotation_deg = 12.5;
  spec.line = fm_chart_line_t{};
  spec.line.color = Rgb(0xFF102030U);
  spec.line.has_width_pt = 1;
  spec.line.width_pt = 2.25;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  fm_shape_info_t out = ShapeById(wb.handle, id);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 12.5);
  EXPECT_EQ(out.line.has_width_pt, 1);
  EXPECT_DOUBLE_EQ(out.line.width_pt, 2.25);

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  out = ShapeById(again.handle, id);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 12.5);
  EXPECT_EQ(out.line.has_width_pt, 1);
  EXPECT_DOUBLE_EQ(out.line.width_pt, 2.25);

  // A patch that sets only the rotation turns the shape and leaves the width.
  fm_shape_patch_t patch = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(patch.rotation_deg_op, patch.rotation_deg, 90.0);
  ASSERT_EQ(fm_sheet_update_shape(again.handle, 0, id, nullptr, &patch), 0) << fm_get_last_error_message();
  out = ShapeById(again.handle, id);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 90.0);
  EXPECT_EQ(out.line.has_width_pt, 1);
  EXPECT_DOUBLE_EQ(out.line.width_pt, 2.25);
}

TEST(FormulonCApiShapes, RotationIsNormalisedAndRangeChecked) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t spec = Callout();
  spec.rotation_deg = -90.0;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_DOUBLE_EQ(ShapeById(wb.handle, id).rotation_deg, 270.0);

  spec.rotation_deg = 450.0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_DOUBLE_EQ(ShapeById(wb.handle, id).rotation_deg, 90.0);

  for (const double bad :
       {3601.0, -3601.0, std::numeric_limits<double>::infinity(), std::numeric_limits<double>::quiet_NaN()}) {
    spec.rotation_deg = bad;
    EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument) << bad;
  }
}

TEST(FormulonCApiShapes, ThemeColorsRoundTripAndAreValidated) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t spec = Callout();
  spec.fill.color = Theme(4, 0.4);
  spec.line = fm_chart_line_t{};
  spec.line.color = Theme(1, -0.25);
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  const fm_shape_info_t out = ShapeById(wb.handle, id);
  EXPECT_EQ(out.fill.color.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(out.fill.color.theme, 4U);
  EXPECT_DOUBLE_EQ(out.fill.color.tint, 0.4);
  EXPECT_EQ(out.line.color.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(out.line.color.theme, 1U);
  EXPECT_DOUBLE_EQ(out.line.color.tint, -0.25);

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
    fm_shape_spec_t s = Callout();
    s.fill.color = fm_color_t{};
    s.fill.color.kind = b.kind;
    s.fill.color.theme = b.theme;
    s.fill.color.tint = b.tint;
    EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &s, &id), kInvalidArgument)
        << "kind=" << b.kind << " theme=" << b.theme << " tint=" << b.tint;
  }

  // A colour given with `none` is refused rather than dropped.
  fm_shape_spec_t none = Callout();
  none.fill.none = 1;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &none, &id), kInvalidArgument);
  none = Callout();
  none.line.color = Rgb(0xFF000000U);
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &none, &id), kInvalidArgument);
  // A colour left None states no colour.
  none.line.color = fm_color_t{};
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &none, &id), 0) << fm_get_last_error_message();
}

// Clearing every member of the fill and outline patches removes the shape's own
// fill and outline so the style's paint applies; it reads all-zero and survives
// a save and reload.
TEST(FormulonCApiShapes, ClearingEveryPaintMemberRestoresTheStylesPaint) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t spec = Callout();
  spec.line = fm_chart_line_t{};
  spec.line.color = Rgb(0xFF0000FFU);
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  fm_shape_patch_t reset = FM_STRUCT_INIT(fm_shape_patch_t);
  reset.fill.color_op = reset.fill.transparency_pct_op = reset.fill.none_op = FM_PATCH_OP_CLEAR;
  reset.line.color_op = reset.line.transparency_pct_op = reset.line.width_pt_op = FM_PATCH_OP_CLEAR;
  reset.line.dash_op = reset.line.none_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &reset), 0) << fm_get_last_error_message();
  fm_shape_info_t out = ShapeById(wb.handle, id);
  EXPECT_EQ(out.fill.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(out.fill.none, 0);
  EXPECT_EQ(out.line.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(out.line.none, 0);

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  out = ShapeById(again.handle, id);
  EXPECT_EQ(out.fill.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(out.line.color.kind, FM_COLOR_KIND_NONE);

  // No field outside the paint patches can be cleared, and a refused patch changes nothing.
  fm_shape_patch_t clear = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(clear.fill.color_op, clear.fill.color, Rgb(0xFF00FF00U));
  clear.rotation_deg_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_shape(again.handle, 0, id, nullptr, &clear), kInvalidArgument);
  fm_shape_patch_t clear_text = FM_STRUCT_INIT(fm_shape_patch_t);
  clear_text.text_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_shape(again.handle, 0, id, nullptr, &clear_text), kInvalidArgument);
  out = ShapeById(again.handle, id);
  EXPECT_EQ(out.fill.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 5.0);
}

// A paint patch merges member by member into the stored paint.
TEST(FormulonCApiShapes, PaintPatchMergesIntoTheStoredPaint) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_shape_spec_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  fm_shape_patch_t colour = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(colour.line.none_op, colour.line.none, 0);
  Set(colour.line.color_op, colour.line.color, Rgb(0xFF0000FFU));
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &colour), 0) << fm_get_last_error_message();
  fm_shape_info_t out = ShapeById(wb.handle, id);
  EXPECT_EQ(out.line.none, 0);
  EXPECT_EQ(out.line.color.argb, 0xFF0000FFU);

  fm_shape_patch_t width = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(width.line.width_pt_op, width.line.width_pt, 3.0);
  Set(width.fill.transparency_pct_op, width.fill.transparency_pct, 40.0);
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &width), 0) << fm_get_last_error_message();
  out = ShapeById(wb.handle, id);
  EXPECT_EQ(out.line.color.argb, 0xFF0000FFU);
  EXPECT_DOUBLE_EQ(out.line.width_pt, 3.0);
  EXPECT_EQ(out.fill.color.argb, 0xFFFFCC00U);
  EXPECT_EQ(out.fill.has_transparency_pct, 1);
  EXPECT_NEAR(out.fill.transparency_pct, 40.0, 1e-9);
}

// An anchor update through the common drawing patch merges field by field into
// the current anchor.
TEST(FormulonCApiShapes, DrawingPatchMergesAnchorFieldsIntoTheCurrentAnchor) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_shape_spec_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  fm_drawing_patch_t move = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(move.anchor.to_col_op, move.anchor.to_col, 7U);
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &move), 0) << fm_get_last_error_message();
  fm_drawing_object_info_t out = ObjectById(wb.handle, id);
  EXPECT_EQ(out.anchor.from_row, 2U);
  EXPECT_EQ(out.anchor.from_col, 1U);
  EXPECT_EQ(out.anchor.to_row, 6U);
  EXPECT_EQ(out.anchor.to_col, 7U);
  EXPECT_EQ(out.anchor.edit_as, FM_ANCHOR_EDIT_AS_ONE_CELL);

  // A kind equal to the current one is accepted.
  fm_drawing_patch_t same = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(same.anchor.kind_op, same.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &same), 0) << fm_get_last_error_message();
  // A size on a two-cell anchor is refused and changes nothing.
  fm_drawing_patch_t sized = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(sized.anchor.width_emu_op, sized.anchor.width_emu, 100000);
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &sized), kInvalidArgument);
  EXPECT_EQ(ObjectById(wb.handle, id).anchor.to_col, 7U);
}

// Every field of `a` as SET in an anchor patch: a read record passed back whole.
fm_drawing_patch_t WholeAnchorPatch(const fm_drawing_anchor_t& a) {
  fm_drawing_patch_t p = FM_STRUCT_INIT(fm_drawing_patch_t);
  fm_drawing_anchor_patch_t& q = p.anchor;
  Set(q.kind_op, q.kind, a.kind);
  Set(q.edit_as_op, q.edit_as, a.edit_as);
  Set(q.from_row_op, q.from_row, a.from_row);
  Set(q.from_col_op, q.from_col, a.from_col);
  Set(q.from_row_offset_emu_op, q.from_row_offset_emu, a.from_row_offset_emu);
  Set(q.from_col_offset_emu_op, q.from_col_offset_emu, a.from_col_offset_emu);
  Set(q.to_row_op, q.to_row, a.to_row);
  Set(q.to_col_op, q.to_col, a.to_col);
  Set(q.to_row_offset_emu_op, q.to_row_offset_emu, a.to_row_offset_emu);
  Set(q.to_col_offset_emu_op, q.to_col_offset_emu, a.to_col_offset_emu);
  Set(q.width_emu_op, q.width_emu, a.width_emu);
  Set(q.height_emu_op, q.height_emu, a.height_emu);
  return p;
}

void ExpectSameAnchor(const fm_drawing_anchor_t& a, const fm_drawing_anchor_t& b) {
  EXPECT_EQ(a.kind, b.kind);
  EXPECT_EQ(a.edit_as, b.edit_as);
  EXPECT_EQ(a.from_row, b.from_row);
  EXPECT_EQ(a.from_col, b.from_col);
  EXPECT_EQ(a.from_row_offset_emu, b.from_row_offset_emu);
  EXPECT_EQ(a.from_col_offset_emu, b.from_col_offset_emu);
  EXPECT_EQ(a.to_row, b.to_row);
  EXPECT_EQ(a.to_col, b.to_col);
  EXPECT_EQ(a.to_row_offset_emu, b.to_row_offset_emu);
  EXPECT_EQ(a.to_col_offset_emu, b.to_col_offset_emu);
  EXPECT_EQ(a.width_emu, b.width_emu);
  EXPECT_EQ(a.height_emu, b.height_emu);
}

// A shape's anchor converts between one-cell and two-cell through the drawing
// patch: one-cell keeps the box's size, two-cell takes its end marker from the
// current size, and the converted anchor survives a save and reload.
TEST(FormulonCApiShapes, DrawingPatchConvertsTheShapeAnchorKind) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_shape_spec_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();
  const fm_drawing_anchor_t box = ObjectById(wb.handle, id).anchor;

  fm_drawing_patch_t one = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(one.anchor.kind_op, one.anchor.kind, FM_ANCHOR_KIND_ONE_CELL);
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &one), 0) << fm_get_last_error_message();
  fm_drawing_anchor_t a = ObjectById(wb.handle, id).anchor;
  EXPECT_EQ(a.kind, FM_ANCHOR_KIND_ONE_CELL);
  EXPECT_EQ(a.from_row, 2U);
  EXPECT_EQ(a.from_col, 1U);
  EXPECT_EQ(a.width_emu, box.width_emu);
  EXPECT_EQ(a.height_emu, box.height_emu);
  EXPECT_EQ(ShapeById(wb.handle, id).fill.color.argb, 0xFFFFCC00U);

  fm_drawing_patch_t two = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(two.anchor.kind_op, two.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &two), 0) << fm_get_last_error_message();
  a = ObjectById(wb.handle, id).anchor;
  EXPECT_EQ(a.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(a.to_row, 6U);
  EXPECT_EQ(a.to_col, 4U);

  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &one), 0) << fm_get_last_error_message();
  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  a = ObjectById(again.handle, id).anchor;
  EXPECT_EQ(a.kind, FM_ANCHOR_KIND_ONE_CELL);
  EXPECT_EQ(a.width_emu, box.width_emu);
  EXPECT_EQ(ShapeById(again.handle, id).fill.color.argb, 0xFFFFCC00U);
}

// A read anchor passed back whole is accepted and changes nothing: a two-cell
// result takes its box from `to_*` and ignores the size, a one-cell result
// takes the size and ignores `to_*`.
TEST(FormulonCApiShapes, AReadAnchorFeedsBackIntoTheDrawingPatch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_shape_spec_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  const fm_drawing_anchor_t two_cell = ObjectById(wb.handle, id).anchor;
  fm_drawing_patch_t whole = WholeAnchorPatch(two_cell);
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &whole), 0) << fm_get_last_error_message();
  ExpectSameAnchor(ObjectById(wb.handle, id).anchor, two_cell);

  fm_drawing_patch_t one = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(one.anchor.kind_op, one.anchor.kind, FM_ANCHOR_KIND_ONE_CELL);
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &one), 0) << fm_get_last_error_message();
  const fm_drawing_anchor_t one_cell = ObjectById(wb.handle, id).anchor;
  whole = WholeAnchorPatch(one_cell);
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &whole), 0) << fm_get_last_error_message();
  ExpectSameAnchor(ObjectById(wb.handle, id).anchor, one_cell);

  // A one-cell read handed back with a two-cell kind takes the box from `to_*`, which the read holds as zero:
  // the end marker lies before the start.
  whole.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &whole), kInvalidArgument);
}

TEST(FormulonCApiShapes, Validation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  uint32_t id = 0;
  fm_shape_spec_t spec = Callout();
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 7, &spec, &id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  spec.placement.anchor.kind = 9;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.placement.anchor.kind = FM_ANCHOR_KIND_ABSOLUTE;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.placement.anchor = fm_drawing_anchor_t{};  // no box and no size
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.placement.struct_size = 0;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.placement.hidden = 3;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.placement.name = "\xC3";
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  spec = Callout();
  spec.text = "\xC3";
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  spec = Callout();
  spec.fill.other = 1;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.line.other = 1;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.fill.none = 9;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec = Callout();
  spec.placement.anchor.to_row = 0;  // two-cell box ends before it starts
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  EXPECT_EQ(ObjectCount(wb.handle), 0U);

  const fm_shape_spec_t valid = Callout();
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &valid, &id), 0) << fm_get_last_error_message();
  fm_shape_info_t out = FM_STRUCT_INIT(fm_shape_info_t);
  int32_t present = 1;
  EXPECT_EQ(fm_sheet_get_shape(wb.handle, 0, id, nullptr, &out), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_shape(wb.handle, 0, id, &present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_shape(wb.handle, 7, id, &present, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  ASSERT_EQ(fm_sheet_get_shape(wb.handle, 0, id + 5U, &present, &out), 0);
  EXPECT_EQ(present, 0);
  const fm_shape_patch_t keep = FM_STRUCT_INIT(fm_shape_patch_t);
  EXPECT_EQ(fm_sheet_update_shape(nullptr, 0, id, nullptr, &keep), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, nullptr), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 7, id, nullptr, &keep), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  fm_shape_patch_t bad_op = FM_STRUCT_INIT(fm_shape_patch_t);
  bad_op.rotation_deg_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &bad_op), kInvalidArgument);
  fm_shape_patch_t null_text = FM_STRUCT_INIT(fm_shape_patch_t);
  null_text.text_op = FM_PATCH_OP_SET;
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &null_text), kInvalidArgument);
}

TEST(FormulonCApiShapeUpdateAtomic, FailureChangesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_shape_spec_t spec = Callout();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  // A valid common half with a refused shape half writes neither.
  fm_drawing_patch_t common = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(common.name_op, common.name, "Renamed");
  Set(common.hidden_op, common.hidden, 1);
  fm_shape_patch_t wide = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(wide.rotation_deg_op, wide.rotation_deg, 45.0);
  Set(wide.line.none_op, wide.line.none, 0);
  Set(wide.line.width_pt_op, wide.line.width_pt, 5000.0);
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, &common, &wide), kInvalidArgument);
  fm_drawing_object_info_t frame = ObjectById(wb.handle, id);
  EXPECT_STREQ(frame.name, "Callout");
  EXPECT_EQ(frame.hidden, 0);
  fm_shape_info_t out = ShapeById(wb.handle, id);
  EXPECT_DOUBLE_EQ(out.rotation_deg, 5.0);
  EXPECT_EQ(out.line.none, 1);

  // Text on a connector is refused with the name it came with.
  fm_shape_spec_t connector = FM_STRUCT_INIT(fm_shape_spec_t);
  connector.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  connector.placement.name = "Link";
  connector.preset_geometry = "straightConnector1";
  connector.placement.anchor.kind = FM_ANCHOR_KIND_ONE_CELL;
  connector.placement.anchor.width_emu = 914400;
  uint32_t line_id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &connector, &line_id), 0) << fm_get_last_error_message();
  fm_shape_patch_t text = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(text.text_op, text.text, "label");
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, line_id, &common, &text), kInvalidArgument);
  EXPECT_STREQ(ObjectById(wb.handle, line_id).name, "Link");

  // A refused common half keeps the shape half unapplied.
  fm_drawing_patch_t empty_name = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(empty_name.name_op, empty_name.name, "");
  fm_shape_patch_t turn = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(turn.rotation_deg_op, turn.rotation_deg, 30.0);
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, &empty_name, &turn), kInvalidArgument);
  EXPECT_DOUBLE_EQ(ShapeById(wb.handle, id).rotation_deg, 5.0);

  // Both halves valid: one call applies both.
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, &common, &turn), 0) << fm_get_last_error_message();
  EXPECT_DOUBLE_EQ(ShapeById(wb.handle, id).rotation_deg, 30.0);
  frame = ObjectById(wb.handle, id);
  EXPECT_STREQ(frame.name, "Renamed");
  EXPECT_EQ(frame.hidden, 1);
}

TEST(FormulonCApiShapePaintRecord, FillLineRecords) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t spec = Callout();
  spec.fill.has_transparency_pct = 1;
  spec.fill.transparency_pct = 25.0;
  spec.line = fm_chart_line_t{};
  spec.line.color = Theme(1, 0.0);
  spec.line.has_width_pt = 1;
  spec.line.width_pt = 2.0;
  spec.line.has_dash = 1;
  spec.line.dash = FM_CHART_LINE_DASH_DASH;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  const fm_shape_info_t out = ShapeById(again.handle, id);
  EXPECT_EQ(out.fill.color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(out.fill.color.argb, 0xFFFFCC00U);
  EXPECT_EQ(out.fill.has_transparency_pct, 1);
  EXPECT_NEAR(out.fill.transparency_pct, 25.0, 1e-9);
  EXPECT_EQ(out.fill.none, 0);
  EXPECT_EQ(out.fill.other, 0);
  EXPECT_EQ(out.line.color.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(out.line.color.theme, 1U);
  EXPECT_EQ(out.line.has_width_pt, 1);
  EXPECT_DOUBLE_EQ(out.line.width_pt, 2.0);
  EXPECT_EQ(out.line.has_dash, 1);
  EXPECT_EQ(out.line.dash, FM_CHART_LINE_DASH_DASH);
  EXPECT_EQ(out.line.none, 0);
  EXPECT_EQ(out.line.other, 0);

  // A read record feeds back into a spec unchanged.
  fm_shape_spec_t copy = Callout();
  copy.fill = out.fill;
  copy.line = out.line;
  uint32_t copy_id = 0;
  ASSERT_EQ(fm_sheet_add_shape(again.handle, 0, &copy, &copy_id), 0) << fm_get_last_error_message();
  const fm_shape_info_t copied = ShapeById(again.handle, copy_id);
  EXPECT_EQ(copied.fill.color.argb, 0xFFFFCC00U);
  EXPECT_NEAR(copied.fill.transparency_pct, 25.0, 1e-9);
  EXPECT_DOUBLE_EQ(copied.line.width_pt, 2.0);
  EXPECT_EQ(copied.line.dash, FM_CHART_LINE_DASH_DASH);

  // A line width beyond 1584 pt is refused.
  fm_shape_spec_t wide = Callout();
  wide.line = fm_chart_line_t{};
  wide.line.has_width_pt = 1;
  wide.line.width_pt = 1585.0;
  EXPECT_EQ(fm_sheet_add_shape(again.handle, 0, &wide, &copy_id), kInvalidArgument);
}

TEST(FormulonCApiShapePresetGeometry, UnknownIsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t spec = Callout();
  spec.preset_geometry = "notAShape";
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec.preset_geometry = "Rect";  // names are case-sensitive
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), kInvalidArgument);
  EXPECT_EQ(ObjectCount(wb.handle), 0U);

  spec.preset_geometry = nullptr;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_STREQ(ShapeById(wb.handle, id).preset_geometry, "rect");

  fm_shape_patch_t unknown = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(unknown.preset_geometry_op, unknown.preset_geometry, "roundedBox");
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &unknown), kInvalidArgument);
  EXPECT_STREQ(ShapeById(wb.handle, id).preset_geometry, "rect");
  fm_shape_patch_t known = FM_STRUCT_INIT(fm_shape_patch_t);
  Set(known.preset_geometry_op, known.preset_geometry, "ellipse");
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &known), 0) << fm_get_last_error_message();
  EXPECT_STREQ(ShapeById(wb.handle, id).preset_geometry, "ellipse");
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
  g.line_width_pt = 1.25;
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
  const char* added_id = nullptr;
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 2, &added_id), 0) << fm_get_last_error_message();
  ASSERT_NE(added_id, nullptr);
  const std::string first_id = added_id;
  EXPECT_FALSE(first_id.empty());
  g.type = FM_SPARKLINE_TYPE_LINE;
  g.markers = 0;
  g.has_manual_min = 0;
  g.color_series = fm_color_t{};
  const fm_sparkline_t more[] = {{"Sheet1!A9:D9", "G9"}};
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, more, 1, &added_id), 0) << fm_get_last_error_message();
  const std::string second_id = added_id;
  EXPECT_NE(second_id, first_id);
  // Both groups read back through their ids, in the session that added them.
  fm_sparkline_group_t by_id = FM_STRUCT_INIT(fm_sparkline_group_t);
  int32_t id_present = 0;
  ASSERT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, second_id.c_str(), &id_present, &by_id), 0);
  ASSERT_EQ(id_present, 1);
  EXPECT_STREQ(by_id.id, second_id.c_str());
  EXPECT_EQ(by_id.type, FM_SPARKLINE_TYPE_LINE);
  ASSERT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, "no such group", &id_present, &by_id), 0);
  EXPECT_EQ(id_present, 0);

  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  ASSERT_EQ(fm_sheet_get_sparkline_group_count(again.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  fm_sparkline_group_t out = FM_STRUCT_INIT(fm_sparkline_group_t);
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(again.handle, 0, 0, &out), 0);
  ASSERT_NE(out.id, nullptr);
  const std::string reloaded_id = out.id;  // a load assigns fresh ids
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
  EXPECT_DOUBLE_EQ(out.line_width_pt, 1.25);
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
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, reloaded_id.c_str(), &patch), 0)
      << fm_get_last_error_message();
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
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, reloaded_id.c_str(), &replace), 0)
      << fm_get_last_error_message();
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
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, reloaded_id.c_str(), &clear), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(again.handle, 0, 0, &changed), 0);
  EXPECT_EQ(changed.date_axis_ref, nullptr);

  // CLEAR removes a colour, which then reads as no colour.
  fm_sparkline_group_patch_t uncolour = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  uncolour.color_high_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_sparkline_group(again.handle, 0, reloaded_id.c_str(), &uncolour), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(again.handle, 0, 0, &changed), 0);
  EXPECT_EQ(changed.color_high.kind, FM_COLOR_KIND_NONE);

  ASSERT_EQ(fm_sheet_remove_sparkline_group(again.handle, 0, reloaded_id.c_str()), 0);
  ASSERT_EQ(fm_sheet_get_sparkline_group_count(again.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  // The removed group's id stays unused, and the other one keeps its own.
  ASSERT_EQ(fm_sheet_get_sparkline_group(again.handle, 0, reloaded_id.c_str(), &id_present, &by_id), 0);
  EXPECT_EQ(id_present, 0);
  EXPECT_EQ(fm_sheet_remove_sparkline_group(again.handle, 0, reloaded_id.c_str()), FM_STATUS_NOT_FOUND);
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(again.handle, 0, 0, &changed), 0);
  ASSERT_NE(changed.id, nullptr);
  const std::string remaining_id = changed.id;
  EXPECT_NE(remaining_id, reloaded_id);
  ASSERT_EQ(fm_sheet_remove_sparkline_group(again.handle, 0, remaining_id.c_str()), 0);
  ASSERT_EQ(fm_sheet_get_sparkline_group_count(again.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiSparklines, Validation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t good[] = {{"Sheet1!A1:D1", "F1"}};
  fm_sparkline_group_t g = ColumnGroup();
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, nullptr, good, 1, ignored_id()), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, nullptr, 1, ignored_id()), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, good, 0, ignored_id()), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 3, &g, good, 1, ignored_id()), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  const fm_sparkline_t range_location[] = {{"Sheet1!A1:D1", "F1:F2"}};
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, range_location, 1, ignored_id()), kInvalidArgument);
  const fm_sparkline_t no_formula[] = {{"", "F1"}};
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, no_formula, 1, ignored_id()), kInvalidArgument);
  fm_sparkline_group_t bad = g;
  bad.type = 7;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &bad, good, 1, ignored_id()), kInvalidArgument);
  bad = g;
  bad.line_width_pt = -1.0;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &bad, good, 1, ignored_id()), kInvalidArgument);
  // A group always states how it draws empty cells.
  bad = g;
  bad.empty_cells_as = FM_BLANKS_AS_ABSENT;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &bad, good, 1, ignored_id()), kInvalidArgument);
  bad.empty_cells_as = FM_BLANKS_AS_SPAN + 1;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &bad, good, 1, ignored_id()), kInvalidArgument);
  size_t count = 5;
  ASSERT_EQ(fm_sheet_get_sparkline_group_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  const fm_sparkline_group_patch_t keep = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, "7", &keep), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, "7", nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, nullptr, &keep), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 4, "7", &keep), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, "\xC3", &keep), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  EXPECT_EQ(fm_sheet_remove_sparkline_group(wb.handle, 0, "7"), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_remove_sparkline_group(wb.handle, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_remove_sparkline_group(wb.handle, 4, "7"), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  fm_sparkline_group_t out = FM_STRUCT_INIT(fm_sparkline_group_t);
  int32_t present = 1;
  EXPECT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, nullptr, &present, &out), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, "7", nullptr, &out), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, "7", &present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_sparkline_group(wb.handle, 4, "7", &present, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, "\xC3", &present, &out), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  EXPECT_EQ(fm_sheet_get_sparkline_group_at(wb.handle, 0, 0, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_sparkline_group_at(wb.handle, 0, 0, nullptr), kBindingNullPointer);
}

TEST(FormulonCApiSparklines, UpdateWithItemsIsAllOrNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}, {"Sheet1!A2:D2", "F2"}};
  fm_sparkline_group_t g = ColumnGroup();
  const char* added = nullptr;
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 2, &added), 0) << fm_get_last_error_message();
  const std::string gid = added;

  // A group cannot hold zero sparklines (CT_Sparklines requires one), so an
  // empty SET items is rejected like an invalid item, and no field lands.
  fm_sparkline_group_patch_t patch = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  Set(patch.type_op, patch.type, FM_SPARKLINE_TYPE_LINE);
  patch.items_op = FM_PATCH_OP_SET;
  const fm_sparkline_t none[] = {{"Sheet1!A1:D1", "F1"}};
  patch.items = none;
  patch.items_count = 0;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, gid.c_str(), &patch), kInvalidArgument);
  const fm_sparkline_t bad[] = {{"Sheet1!A1:D1", "F1:F2"}};
  patch.items = bad;
  patch.items_count = 1;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, gid.c_str(), &patch), kInvalidArgument);
  patch.items = nullptr;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, gid.c_str(), &patch), kBindingNullPointer);
  // An invalid field with valid items does not replace the items either.
  const fm_sparkline_t one[] = {{"Sheet1!B1:B5", "H1"}};
  patch.items = one;
  Set(patch.line_width_pt_op, patch.line_width_pt, -1.0);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, gid.c_str(), &patch), kInvalidArgument);
  // Out-of-domain values and ops are refused before anything is written.
  fm_sparkline_group_patch_t bad_field = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  Set(bad_field.type_op, bad_field.type, 7);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, gid.c_str(), &bad_field), kInvalidArgument);
  bad_field = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  bad_field.markers_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, gid.c_str(), &bad_field), kInvalidArgument);
  bad_field = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  Set(bad_field.date_axis_ref_op, bad_field.date_axis_ref, nullptr);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, gid.c_str(), &bad_field), kInvalidArgument);
  bad_field = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  Set(bad_field.empty_cells_as_op, bad_field.empty_cells_as, FM_BLANKS_AS_ABSENT);
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, gid.c_str(), &bad_field), kInvalidArgument);

  fm_sparkline_group_t read = FM_STRUCT_INIT(fm_sparkline_group_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, gid.c_str(), &present, &read), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(read.type, FM_SPARKLINE_TYPE_COLUMN);
  EXPECT_EQ(read.sparkline_count, 2U);
}

// An out-of-range position names the parameter it came from in the error context; an
// unknown group id names the id.
TEST(FormulonCApiSparklines, ErrorContextNamesTheParameter) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}};
  fm_sparkline_group_t g = ColumnGroup();
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, ignored_id()), 0) << fm_get_last_error_message();
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
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, "5", &keep), FM_STATUS_NOT_FOUND);
  EXPECT_TRUE(context_has("group_id=5"));
  EXPECT_EQ(fm_sheet_remove_sparkline_group(wb.handle, 0, "6"), FM_STATUS_NOT_FOUND);
  EXPECT_TRUE(context_has("group_id=6"));
}

std::string SavedSheetXml(fm_workbook_t* wb) {
  BufferGuard buf;
  EXPECT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &buf.data, &buf.len), 0) << fm_get_last_error_message();
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
    ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, ignored_id()), 0) << fm_get_last_error_message();
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
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 2, ignored_id()), 0) << fm_get_last_error_message();
  g.date_axis_ref = "Sheet1!A6:D6";
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 2, ignored_id()), 0) << fm_get_last_error_message();

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

TEST(FormulonCApiSparklineLineWidth, RoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}};
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  g.line_width_pt = 2.5;
  const char* added = nullptr;
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, &added), 0) << fm_get_last_error_message();
  ASSERT_NE(added, nullptr);
  const std::string gid = added;
  fm_sparkline_group_t out = FM_STRUCT_INIT(fm_sparkline_group_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, gid.c_str(), &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_DOUBLE_EQ(out.line_width_pt, 2.5);

  fm_sparkline_group_patch_t patch = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  Set(patch.line_width_pt_op, patch.line_width_pt, 0.5);
  ASSERT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, gid.c_str(), &patch), 0) << fm_get_last_error_message();
  WorkbookGuard again;
  SaveAndReload(wb.handle, &again);
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(again.handle, 0, 0, &out), 0);
  EXPECT_DOUBLE_EQ(out.line_width_pt, 0.5);
}

}  // namespace
