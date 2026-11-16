// Empty strings in the DRAWING / SHAPE / SLICER / CHART / SPARKLINE setters (api-conventions §5 rule 30): where the
// file holds one state for "" and the absent value they mean the same; where "" is not a legal value it is refused
// and the absent value clears or derives.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

using Bytes = std::vector<std::uint8_t>;

Bytes Png(std::uint32_t w, std::uint32_t h) {
  Bytes b = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  for (std::uint32_t v : {w, h}) {
    for (int shift = 24; shift >= 0; shift -= 8) {
      b.push_back(static_cast<std::uint8_t>(v >> shift));
    }
  }
  b.insert(b.end(), {8, 6, 0, 0, 0, 0x1F, 0x15, 0xC4, 0x89});
  return b;
}

/// A two-cell box at B2:E6, the size every kind accepts.
fm_drawing_placement_spec_t Placement() {
  fm_drawing_placement_spec_t p = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  p.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  p.anchor.from_row = 1;
  p.anchor.from_col = 1;
  p.anchor.to_row = 5;
  p.anchor.to_col = 4;
  return p;
}

size_t ObjectCount(const fm_workbook_t* wb) {
  size_t count = 99;
  EXPECT_EQ(fm_sheet_get_drawing_object_count(wb, 0, &count), 0);
  return count;
}

fm_drawing_object_info_t Frame(const fm_workbook_t* wb, std::uint32_t id) {
  fm_drawing_object_info_t out = FM_STRUCT_INIT(fm_drawing_object_info_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_drawing_object(wb, 0, id, &present, &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return out;
}

bool NoText(const char* text) {
  return text == nullptr || text[0] == '\0';
}

void FillGrid(fm_workbook_t* wb) {
  for (std::uint32_t c = 0; c < 3; ++c) {
    ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, c + 1, "S"), 0);
    for (std::uint32_t r = 0; r < 4; ++r) {
      ASSERT_EQ(fm_sheet_set_number(wb, 0, r + 1, c + 1, static_cast<double>((r + 1) * (c + 2))), 0);
    }
  }
  for (std::uint32_t r = 0; r < 4; ++r) {
    ASSERT_EQ(fm_sheet_set_text(wb, 0, r + 1, 0, "a"), 0);
  }
}

// ---------------------------------------------------------------------------
// Placement: description is one state with the absent value, name is not a legal "" value
// ---------------------------------------------------------------------------

TEST(DrawingEmptyStrings, PictureDescriptionEmptyAndAbsentAreOneState) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(4, 4);
  fm_drawing_placement_spec_t empty = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  empty.description = "";
  fm_drawing_placement_spec_t absent = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  uint32_t empty_id = 0;
  uint32_t absent_id = 0;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &empty, &empty_id), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &absent, &absent_id), 0)
      << fm_get_last_error_message();
  EXPECT_TRUE(NoText(Frame(wb.handle, empty_id).description));
  EXPECT_TRUE(NoText(Frame(wb.handle, absent_id).description));
}

TEST(DrawingEmptyStrings, PlacementNameEmptyIsRefusedAndAbsentDerivesOnEveryKind) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const Bytes png = Png(4, 4);

  fm_drawing_placement_spec_t picture = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  picture.name = "";
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &picture, &id), kInvalidArgument);

  fm_shape_spec_t shape = FM_STRUCT_INIT(fm_shape_spec_t);
  shape.placement = Placement();
  shape.placement.name = "";
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &shape, &id), kInvalidArgument);

  fm_chart_spec_t chart = FM_STRUCT_INIT(fm_chart_spec_t);
  chart.type = FM_CHART_PLOT_TYPE_LINE;
  chart.bar_direction = FM_CHART_BAR_DIRECTION_COLUMN;
  chart.source = "Sheet1!$A$1:$D$5";
  chart.placement = Placement();
  chart.placement.name = "";
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &chart, &id), kInvalidArgument);

  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 8, 0, "Product"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 9, 0, "Widget"), 0);
  const char* columns[] = {"Product"};
  std::size_t table = 0;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{8, 0, 9, 0}), "Sales", columns, 1, "TableStyleMedium2", 1,
                     0, &table),
            0);
  fm_slicer_spec_t slicer = FM_STRUCT_INIT(fm_slicer_spec_t);
  slicer.placement = Placement();
  slicer.placement.name = "";
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, table, 0, &slicer, &id), kInvalidArgument);

  EXPECT_EQ(ObjectCount(wb.handle), 0U);

  // The absent value derives a name on each kind.
  picture.name = nullptr;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &picture, &id), 0);
  EXPECT_FALSE(NoText(Frame(wb.handle, id).name));
  shape.placement.name = nullptr;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &shape, &id), 0) << fm_get_last_error_message();
  EXPECT_FALSE(NoText(Frame(wb.handle, id).name));
  chart.placement.name = nullptr;
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &chart, &id), 0) << fm_get_last_error_message();
  EXPECT_FALSE(NoText(Frame(wb.handle, id).name));
  slicer.placement.name = nullptr;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, table, 0, &slicer, &id), 0) << fm_get_last_error_message();
  EXPECT_FALSE(NoText(Frame(wb.handle, id).name));
}

TEST(DrawingEmptyStrings, DrawingPatchDescriptionEmptyClearsLikeTheClearOp) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t shape = FM_STRUCT_INIT(fm_shape_spec_t);
  shape.placement = Placement();
  shape.placement.description = "alt";
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &shape, &id), 0) << fm_get_last_error_message();
  EXPECT_STREQ(Frame(wb.handle, id).description, "alt");
  fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.description_op = FM_PATCH_OP_SET;
  patch.description = "";
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, &patch, nullptr), 0) << fm_get_last_error_message();
  EXPECT_TRUE(NoText(Frame(wb.handle, id).description));
}

// ---------------------------------------------------------------------------
// Cell picture alt text: "" is none
// ---------------------------------------------------------------------------

TEST(DrawingEmptyStrings, CellPictureEmptyAltTextIsNone) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(2, 2);
  fm_cell_picture_t picture = FM_STRUCT_INIT(fm_cell_picture_t);
  picture.bytes = png.data();
  picture.len = png.size();
  picture.alt_text = "old";
  ASSERT_EQ(fm_sheet_set_cell_picture(wb.handle, 0, 0, 0, &picture), 0);
  picture.alt_text = "";
  ASSERT_EQ(fm_sheet_set_cell_picture(wb.handle, 0, 0, 0, &picture), 0) << fm_get_last_error_message();
  int32_t present = 0;
  fm_cell_picture_t out = FM_STRUCT_INIT(fm_cell_picture_t);
  ASSERT_EQ(fm_sheet_get_cell_picture(wb.handle, 0, 0, 0, &present, &out), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(out.alt_text, nullptr);
}

// ---------------------------------------------------------------------------
// Shape: presetGeometry "" is refused, text "" is none
// ---------------------------------------------------------------------------

TEST(DrawingEmptyStrings, ShapePresetGeometryEmptyIsRefusedAndAbsentIsRect) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t shape = FM_STRUCT_INIT(fm_shape_spec_t);
  shape.placement = Placement();
  shape.preset_geometry = "";
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &shape, &id), kInvalidArgument);
  EXPECT_EQ(ObjectCount(wb.handle), 0U);

  shape.preset_geometry = nullptr;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &shape, &id), 0) << fm_get_last_error_message();
  fm_shape_info_t info = FM_STRUCT_INIT(fm_shape_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_shape(wb.handle, 0, id, &present, &info), 0);
  ASSERT_NE(info.preset_geometry, nullptr);
  EXPECT_STREQ(info.preset_geometry, "rect");
}

TEST(DrawingEmptyStrings, ShapeTextEmptyAndAbsentAreOneState) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_shape_spec_t shape = FM_STRUCT_INIT(fm_shape_spec_t);
  shape.placement = Placement();
  shape.text = "";
  uint32_t empty_id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &shape, &empty_id), 0) << fm_get_last_error_message();
  shape.text = nullptr;
  uint32_t absent_id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &shape, &absent_id), 0) << fm_get_last_error_message();
  for (const uint32_t id : {empty_id, absent_id}) {
    fm_shape_info_t info = FM_STRUCT_INIT(fm_shape_info_t);
    int32_t present = 0;
    ASSERT_EQ(fm_sheet_get_shape(wb.handle, 0, id, &present, &info), 0);
    EXPECT_TRUE(NoText(info.text));
  }
}

// ---------------------------------------------------------------------------
// Slicer caption: "" takes the column name, as the absent value does
// ---------------------------------------------------------------------------

TEST(DrawingEmptyStrings, SlicerCaptionEmptyAndAbsentBothTakeTheColumnName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Product"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 1, 0, "Widget"), 0);
  const char* columns[] = {"Product"};
  std::size_t table = 0;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 1, 0}), "Sales", columns, 1, "TableStyleMedium2", 1,
                     0, &table),
            0);
  std::string captions[2];
  const char* inputs[] = {"", nullptr};
  for (std::uint32_t i = 0; i < 2U; ++i) {
    fm_slicer_spec_t spec = FM_STRUCT_INIT(fm_slicer_spec_t);
    spec.placement = Placement();
    spec.placement.anchor.from_col = 3U + 5U * i;
    spec.placement.anchor.to_col = 6U + 5U * i;
    spec.caption = inputs[i];
    uint32_t id = 0;
    ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, table, 0, &spec, &id), 0) << fm_get_last_error_message();
    fm_slicer_info_t info = FM_STRUCT_INIT(fm_slicer_info_t);
    int32_t present = 0;
    ASSERT_EQ(fm_sheet_get_slicer(wb.handle, 0, id, &present, &info), 0);
    ASSERT_NE(info.caption, nullptr);
    captions[i] = info.caption;
  }
  EXPECT_EQ(captions[0], "Product");
  EXPECT_EQ(captions[1], captions[0]);
}

// ---------------------------------------------------------------------------
// Sparkline: formula and date_axis_ref "" are refused, the absent value is a deleted source / no date axis
// ---------------------------------------------------------------------------

fm_sparkline_group_t Group() {
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  g.type = FM_SPARKLINE_TYPE_LINE;
  g.line_width_pt = 0.75;
  g.empty_cells_as = FM_BLANKS_AS_GAP;
  return g;
}

size_t GroupCount(const fm_workbook_t* wb) {
  size_t count = 99;
  EXPECT_EQ(fm_sheet_get_sparkline_group_count(wb, 0, &count), 0);
  return count;
}

TEST(DrawingEmptyStrings, SparklineFormulaEmptyIsRefusedAndNullIsADeletedSource) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sparkline_group_t g = Group();
  const char* id = nullptr;
  const fm_sparkline_t empty[] = {{"", "F1"}};
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, empty, 1, &id), kInvalidArgument);
  EXPECT_EQ(GroupCount(wb.handle), 0U);

  const fm_sparkline_t absent[] = {{nullptr, "F1"}};
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, absent, 1, &id), 0) << fm_get_last_error_message();
  fm_sparkline_t read = {};
  ASSERT_EQ(fm_sheet_get_sparkline_at(wb.handle, 0, 0, 0, &read), 0);
  EXPECT_TRUE(NoText(read.formula));
}

TEST(DrawingEmptyStrings, SparklineLocationEmptyIsRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sparkline_group_t g = Group();
  const char* id = nullptr;
  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", ""}};
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, &id), kInvalidArgument);
  EXPECT_EQ(GroupCount(wb.handle), 0U);
}

TEST(DrawingEmptyStrings, SparklineDateAxisRefEmptyIsRefusedOnAddAndPatch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}};
  const char* id = nullptr;
  fm_sparkline_group_t g = Group();
  g.date_axis_ref = "";
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, &id), kInvalidArgument);
  EXPECT_EQ(GroupCount(wb.handle), 0U);

  g.date_axis_ref = nullptr;
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, &id), 0) << fm_get_last_error_message();
  fm_sparkline_group_t read = FM_STRUCT_INIT(fm_sparkline_group_t);
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(wb.handle, 0, 0, &read), 0);
  EXPECT_EQ(read.date_axis_ref, nullptr);

  const std::string group_id = id;
  fm_sparkline_group_patch_t patch = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  patch.date_axis_ref_op = FM_PATCH_OP_SET;
  patch.date_axis_ref = "";
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, group_id.c_str(), &patch), kInvalidArgument);
}

TEST(DrawingEmptyStrings, SparklineItemsReplacementRefusesAnEmptyFormula) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}};
  const char* id = nullptr;
  fm_sparkline_group_t g = Group();
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, &id), 0) << fm_get_last_error_message();
  const std::string group_id = id;
  const fm_sparkline_t empty[] = {{"", "F1"}};
  fm_sparkline_group_patch_t patch = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  patch.items_op = FM_PATCH_OP_SET;
  patch.items = empty;
  patch.items_count = 1;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, group_id.c_str(), &patch), kInvalidArgument);
  fm_sparkline_t read = {};
  ASSERT_EQ(fm_sheet_get_sparkline_at(wb.handle, 0, 0, 0, &read), 0);
  EXPECT_STREQ(read.formula, "=Sheet1!A1:D1");
}

// ---------------------------------------------------------------------------
// Chart: series reference "" is refused, null (CLEAR / NONE) removes; title text "" and absent are one state
// ---------------------------------------------------------------------------

fm_chart_spec_t LineChart() {
  fm_chart_spec_t o = FM_STRUCT_INIT(fm_chart_spec_t);
  o.type = FM_CHART_PLOT_TYPE_LINE;
  o.bar_direction = FM_CHART_BAR_DIRECTION_COLUMN;
  o.source = "Sheet1!$A$1:$D$5";
  o.placement = Placement();
  return o;
}

TEST(DrawingEmptyStrings, ChartSeriesEmptyReferenceIsRefusedForEveryInput) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const fm_chart_spec_t spec = LineChart();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  const auto patch_with = [&](fm_patch_op_t fm_chart_series_patch_t::*op,
                              fm_chart_data_source_spec_t fm_chart_series_patch_t::*src) {
    fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
    patch.*op = FM_PATCH_OP_SET;
    patch.*src = fm_chart_data_source_spec_t{FM_CHART_SOURCE_KIND_REF, "", nullptr, 0};
    return fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch);
  };
  EXPECT_EQ(patch_with(&fm_chart_series_patch_t::name_op, &fm_chart_series_patch_t::name), kInvalidArgument);
  EXPECT_EQ(patch_with(&fm_chart_series_patch_t::categories_op, &fm_chart_series_patch_t::categories),
            kInvalidArgument);
  EXPECT_EQ(patch_with(&fm_chart_series_patch_t::x_op, &fm_chart_series_patch_t::x), kInvalidArgument);
  EXPECT_EQ(patch_with(&fm_chart_series_patch_t::values_op, &fm_chart_series_patch_t::values), kInvalidArgument);

  // CLEAR is the absent value: it removes the series name.
  fm_chart_series_patch_t clear = FM_STRUCT_INIT(fm_chart_series_patch_t);
  clear.name_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &clear), 0) << fm_get_last_error_message();
}

TEST(DrawingEmptyStrings, ChartTitleTextEmptyAndAbsentAreOneState) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  fm_chart_spec_t spec = LineChart();
  spec.has_title = 1;
  uint32_t id = 0;
  spec.title = fm_chart_title_t{FM_CHART_TITLE_KIND_TEXT, "", nullptr};
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &spec, &id), kInvalidArgument);
  spec.title = fm_chart_title_t{FM_CHART_TITLE_KIND_TEXT, nullptr, nullptr};
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &spec, &id), kInvalidArgument);
  EXPECT_EQ(ObjectCount(wb.handle), 0U);
}

}  // namespace
