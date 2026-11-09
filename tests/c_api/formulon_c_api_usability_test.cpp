//
// C ABI contracts shared by the bindings' usability surface: the dirty-cell
// query, the process-scope image probe, the drawing and chart id statuses
// (absent, of another kind, past the end), the drawing patches, the chart
// absence flags, CF rules and protected ranges addressed by their keys, the
// keys and indices adds return, the chart insert title, and the leading `=` on
// formula text read back from cells, rules, tables and pivot caches.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_setters.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kRuleId = "{0A1B2C3D-0000-4000-8000-00000000AB01}";

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

fm_drawing_placement_spec_t PictureOptions() {
  fm_drawing_placement_spec_t opts = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  opts.anchor.kind = FM_ANCHOR_KIND_ONE_CELL;
  opts.anchor.edit_as = FM_ANCHOR_EDIT_AS_TWO_CELL;
  return opts;
}

std::uint32_t AddPicture(fm_workbook_t* wb, const Bytes& png) {
  const fm_drawing_placement_spec_t opts = PictureOptions();
  std::uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_picture(wb, 0, png.data(), png.size(), &opts, &id), 0) << fm_get_last_error_message();
  return id;
}

// A shape boxed B2:D5, two-cell unless `kind` says otherwise.
fm_shape_spec_t ShapeSpec(std::int32_t kind = FM_ANCHOR_KIND_TWO_CELL) {
  fm_shape_spec_t spec = FM_STRUCT_INIT(fm_shape_spec_t);
  spec.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  spec.placement.anchor.kind = kind;
  spec.placement.anchor.from_row = 1;
  spec.placement.anchor.from_col = 1;
  spec.placement.anchor.to_row = 4;
  spec.placement.anchor.to_col = 3;
  return spec;
}

std::uint32_t AddShape(fm_workbook_t* wb) {
  const fm_shape_spec_t spec = ShapeSpec();
  std::uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_shape(wb, 0, &spec, &id), 0) << fm_get_last_error_message();
  return id;
}

fm_shape_info_t ReadShape(const fm_workbook_t* wb, std::uint32_t id) {
  fm_shape_info_t shape = FM_STRUCT_INIT(fm_shape_info_t);
  std::int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_shape(wb, 0, id, &present, &shape), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return shape;
}

fm_drawing_object_info_t ReadDrawing(const fm_workbook_t* wb, std::uint32_t id) {
  fm_drawing_object_info_t object = FM_STRUCT_INIT(fm_drawing_object_info_t);
  std::int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_drawing_object(wb, 0, id, &present, &object), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return object;
}

// A1:B3: a header row over two text categories and their numbers.
void FillChartSource(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, 0, "k"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, 1, "v"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 1, 0, "a"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 2, 0, "b"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, 0, 1, 1, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, 0, 2, 1, 2.0), 0);
}

fm_chart_spec_t ChartOptions() {
  fm_chart_spec_t opts = FM_STRUCT_INIT(fm_chart_spec_t);
  opts.type = FM_CHART_PLOT_TYPE_BAR;
  opts.source = "Sheet1!$A$1:$B$3";
  opts.legend = FM_CHART_LEGEND_POSITION_RIGHT;
  opts.placement = PictureOptions();
  opts.placement.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  opts.placement.anchor.from_row = 6;
  return opts;
}

fm_chart_info_t ReadChartInfo(const fm_workbook_t* wb, std::uint32_t id) {
  fm_chart_info_t info = FM_STRUCT_INIT(fm_chart_info_t);
  std::int32_t present = 0;
  fm_chart_t* chart = nullptr;
  EXPECT_EQ(fm_sheet_get_chart(wb, 0, id, &present, &chart), 0) << fm_get_last_error_message();
  if (chart != nullptr) {
    EXPECT_EQ(fm_chart_get_info(chart, &info), 0);
    fm_chart_destroy(chart);
  }
  return info;
}

std::int32_t Dirty(const fm_workbook_t* wb) {
  std::int32_t dirty = -1;
  EXPECT_EQ(fm_workbook_has_dirty_cells(wb, &dirty), 0) << fm_get_last_error_message();
  return dirty;
}

void SaveAndReload(const fm_workbook_t* wb, WorkbookGuard* out) {
  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &buf.data, &buf.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &out->handle), 0) << fm_get_last_error_message();
}

std::string Str(const char* text) {
  return text != nullptr ? std::string(text) : std::string("<null>");
}

fm_cf_rule_t ExpressionRule(const fm_cell_rect_t* sqref, const char* formula, const char* id = nullptr) {
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.id = id;
  rule.type = 0;  // Expression
  rule.formula1 = formula;
  rule.ranges = sqref;
  rule.range_count = 1;
  return rule;
}

std::size_t CfRuleCount(const fm_workbook_t* wb) {
  std::size_t count = 99;
  EXPECT_EQ(fm_sheet_get_cf_rule_count(wb, 0, &count), 0);
  return count;
}

fm_data_validation_t WholeBetween(const fm_cell_rect_t* sqref, const char* low, const char* high) {
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = sqref;
  v.range_count = 1;
  v.type = 1;  // whole
  v.op = 0;    // between
  v.formula1 = low;
  v.formula2 = high;
  return v;
}

// A progress callback that queries the dirty set of the workbook it reports on.
fm_status_t g_reentrant_status = 0;

extern "C" int32_t query_dirty_from_progress(uint32_t /*iteration*/, double /*max_residual*/,
                                             uint32_t /*max_iterations*/, void* user_data) {
  std::int32_t dirty = -1;
  g_reentrant_status = fm_workbook_has_dirty_cells(static_cast<const fm_workbook_t*>(user_data), &dirty);
  return 0;
}

}  // namespace

// ---------------------------------------------------------------------------
// fm_workbook_has_dirty_cells
// ---------------------------------------------------------------------------

TEST(FormulonCApiUsability, HasDirtyCellsIsFalseOnANewWorkbookAndAfterRecalc) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(Dirty(wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 5.0), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(Dirty(wb.handle), 0);
}

TEST(FormulonCApiUsability, HasDirtyCellsRejectsNullArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::int32_t dirty = 0;
  EXPECT_EQ(fm_workbook_has_dirty_cells(nullptr, &dirty), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_has_dirty_cells(wb.handle, nullptr), FM_STATUS_BINDING_NULL_POINTER);
}

// A loaded workbook holds the file's cached results; its formulas stay pending
// until the first full recalc re-evaluates them.
TEST(FormulonCApiUsability, HasDirtyCellsAfterLoadReflectsThePendingFormulas) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 2.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  WorkbookGuard literals;
  SaveAndReload(wb.handle, &literals);
  EXPECT_EQ(Dirty(literals.handle), 0);

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=A1*2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  WorkbookGuard formulas;
  SaveAndReload(wb.handle, &formulas);
  EXPECT_EQ(Dirty(formulas.handle), 1);
  ASSERT_EQ(fm_workbook_recalc(formulas.handle), 0);
  EXPECT_EQ(Dirty(formulas.handle), 0);
}

TEST(FormulonCApiUsability, HasDirtyCellsAfterAReferencedDefinedNameChanges) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "3"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=Rate*2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(Dirty(wb.handle), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "4"), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
}

TEST(FormulonCApiUsability, HasDirtyCellsAfterInsertRows) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 0, "=SUM(A1:A1)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(Dirty(wb.handle), 0);
  ASSERT_EQ(fm_sheet_insert_rows(wb.handle, 0, 0, 1), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
}

TEST(FormulonCApiUsability, HasDirtyCellsAfterRemovingAReferencedSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t second = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Data", append_position(wb.handle), &second), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, second, 0, 0, 7.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=Data!A1+1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(Dirty(wb.handle), 0);
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, second), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
}

TEST(FormulonCApiUsability, HasDirtyCellsFollowsAProfileChangeOnlyWhenItChanges) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=LEN(\"abc\")"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  const char* current = nullptr;
  ASSERT_EQ(fm_workbook_get_profile_id(wb.handle, &current), 0);
  const std::string same = Str(current);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, same.c_str()), 0);
  EXPECT_EQ(Dirty(wb.handle), 0);
  const char* other = same == "win-365-ja_JP" ? "win-365-en_US" : "win-365-ja_JP";
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, other), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
}

TEST(FormulonCApiUsability, HasDirtyCellsLeavesVolatileCellsOut) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=NOW()"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(Dirty(wb.handle), 0);
}

TEST(FormulonCApiUsability, HasDirtyCellsStaysTrueAfterAViewportRecalc) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=A1+1"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 500, 1, "=A1+2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 2.0), 0);
  const fm_cell_rect_t viewport{0, 0, 9, 9};
  ASSERT_EQ(fm_sheet_recalc_viewport(wb.handle, 0, &viewport, nullptr), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
}

TEST(FormulonCApiUsability, HasDirtyCellsFromTheProgressCallbackIsReentrant) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetIterative(wb.handle, 1, 100, 0.001), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=(A1+10)/2"), 0);
  g_reentrant_status = 0;
  ASSERT_EQ(fm_workbook_set_iterative_progress(wb.handle, &query_dirty_from_progress, wb.handle), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(g_reentrant_status, FM_STATUS_GRAPH_RECALC_REENTRANT);
}

// ---------------------------------------------------------------------------
// fm_get_image_info
// ---------------------------------------------------------------------------

TEST(FormulonCApiUsability, ImageInfoNeedsNoWorkbook) {
  const Bytes png = Png(16, 8);
  fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
  ASSERT_EQ(fm_get_image_info(png.data(), png.size(), &info), 0) << fm_get_last_error_message();
  EXPECT_EQ(info.format, FM_IMAGE_FORMAT_PNG);
  EXPECT_EQ(info.width_px, 16U);
  EXPECT_EQ(info.height_px, 8U);

  const Bytes junk = {'n', 'o', 't', ' ', 'a', 'n', ' ', 'i', 'm', 'a', 'g', 'e'};
  EXPECT_EQ(fm_get_image_info(junk.data(), junk.size(), &info), FM_STATUS_IO_IMAGE_UNSUPPORTED);
  EXPECT_EQ(fm_get_image_info(nullptr, 0, &info), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_get_image_info(png.data(), png.size(), nullptr), FM_STATUS_BINDING_NULL_POINTER);
}

// ---------------------------------------------------------------------------
// Absent drawing and chart ids
// ---------------------------------------------------------------------------

TEST(FormulonCApiUsability, AbsentDrawingAndChartIdsAreNotFound) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_chart_data_t* data = nullptr;
  std::int32_t present = 99;
  const fm_chart_preview_overrides_spec_t overrides = FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
  // A sheet with no drawing at all: a remove is NotFound, a chart read a miss.
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 0, 42U), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_chart_data(wb.handle, 0, 42U, &present, &data), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 0);
  EXPECT_EQ(data, nullptr);
  present = 99;
  EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, 42U, &overrides, &present, &data), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(present, 0);
  EXPECT_EQ(data, nullptr);

  // A drawing whose only object is a picture: its id names no chart.
  const Bytes png = Png(4, 4);
  const std::uint32_t picture = AddPicture(wb.handle, png);
  present = 99;
  EXPECT_EQ(fm_sheet_get_chart_data(wb.handle, 0, picture, &present, &data), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 0);
  present = 99;
  EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, picture, &overrides, &present, &data), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(present, 0);
  EXPECT_EQ(data, nullptr);
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 0, picture + 100U), FM_STATUS_NOT_FOUND);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);

  ASSERT_EQ(fm_sheet_remove_drawing(wb.handle, 0, picture), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 0, picture), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 9, picture), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

// The picture options are a record: a struct_size below its first layout is
// refused, on its own and embedded in a chart insert, and nothing is inserted.
TEST(FormulonCApiUsability, PictureOptionsWithoutAStructSizeAreRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(4, 4);
  fm_drawing_placement_spec_t opts = PictureOptions();
  opts.struct_size = 0;
  std::uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &id), FM_STATUS_INVALID_ARGUMENT);

  FillChartSource(wb.handle);
  fm_chart_spec_t chart = ChartOptions();
  chart.placement.struct_size = 0;
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &chart, &id), FM_STATUS_INVALID_ARGUMENT);
  std::size_t count = 99;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

// An id that exists but belongs to an object of another kind is a wrong
// argument, not a missing one.
TEST(FormulonCApiUsability, AnIdOfAnotherKindOfObjectIsAnInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(4, 4);
  const std::uint32_t picture = AddPicture(wb.handle, png);
  const std::uint32_t shape = AddShape(wb.handle);

  fm_shape_patch_t patch = FM_STRUCT_INIT(fm_shape_patch_t);
  patch.text_box_op = FM_PATCH_OP_SET;
  patch.text_box = 1;
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, picture, nullptr, &patch), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(ReadShape(wb.handle, shape).text_box, 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 2U);
}

// A position at the object or plot count is out of range; a series c:idx no
// series holds is a key that matches nothing.
TEST(FormulonCApiUsability, DrawingAndChartPositionsPastTheEndAreOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(4, 4);
  const std::uint32_t picture = AddPicture(wb.handle, png);
  FillChartSource(wb.handle);
  const fm_chart_spec_t opts = ChartOptions();
  std::uint32_t chart = 0;
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &opts, &chart), 0) << fm_get_last_error_message();

  EXPECT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, picture, 2), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, picture, 1), 0) << fm_get_last_error_message();

  const fm_chart_plot_patch_t plot = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  EXPECT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, chart, 1, &plot), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, chart, 0, &plot), 0) << fm_get_last_error_message();

  fm_chart_series_patch_t series = FM_STRUCT_INIT(fm_chart_series_patch_t);
  series.smooth_op = FM_PATCH_OP_SET;
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, chart, 0, 999U, &series), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_remove_chart_series(wb.handle, 0, chart, 0, 999U), FM_STATUS_NOT_FOUND);
}

// A shape patch writes only its SET fields; a refused field (CLEAR outside the
// paint fields, an op outside 0..2) leaves the whole shape unchanged.
TEST(FormulonCApiUsability, ShapePatchWritesOnlyItsSetFieldsAndRefusesAtomically) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t id = AddShape(wb.handle);
  const fm_shape_info_t before = ReadShape(wb.handle, id);

  fm_shape_patch_t patch = FM_STRUCT_INIT(fm_shape_patch_t);
  patch.rotation_deg_op = FM_PATCH_OP_SET;
  patch.rotation_deg = 90.0;
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &patch), 0) << fm_get_last_error_message();
  fm_shape_info_t after = ReadShape(wb.handle, id);
  EXPECT_EQ(after.rotation_deg, 90.0);
  EXPECT_EQ(after.text_box, before.text_box);
  EXPECT_EQ(std::string(after.preset_geometry), std::string(before.preset_geometry));

  fm_shape_patch_t bad = FM_STRUCT_INIT(fm_shape_patch_t);
  bad.text_box_op = FM_PATCH_OP_SET;
  bad.text_box = 1;
  bad.rotation_deg_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &bad), FM_STATUS_INVALID_ARGUMENT);
  bad.rotation_deg_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &bad), FM_STATUS_INVALID_ARGUMENT);
  after = ReadShape(wb.handle, id);
  EXPECT_EQ(after.text_box, before.text_box);
  EXPECT_EQ(after.rotation_deg, 90.0);

  const fm_shape_patch_t keep = FM_STRUCT_INIT(fm_shape_patch_t);
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &keep), 0) << fm_get_last_error_message();
}

// A shape's own fill is set to no fill and removed by clearing every fill
// member, which returns it to the style's fill.
TEST(FormulonCApiUsability, ShapeFillClearedReturnsToTheStylesFill) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t id = AddShape(wb.handle);

  fm_shape_patch_t none = FM_STRUCT_INIT(fm_shape_patch_t);
  none.fill.none_op = FM_PATCH_OP_SET;
  none.fill.none = 1;
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &none), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadShape(wb.handle, id).fill.none, 1);

  fm_shape_patch_t style = FM_STRUCT_INIT(fm_shape_patch_t);
  style.fill.color_op = FM_PATCH_OP_CLEAR;
  style.fill.transparency_pct_op = FM_PATCH_OP_CLEAR;
  style.fill.none_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &style), 0) << fm_get_last_error_message();
  const fm_shape_info_t shape = ReadShape(wb.handle, id);
  EXPECT_EQ(shape.fill.none, 0);
  EXPECT_EQ(shape.fill.other, 0);
  EXPECT_EQ(shape.fill.color.kind, FM_COLOR_KIND_NONE);
}

// A shape added with only its anchor takes Excel's defaults for the rest, and
// an absolute anchor kind is refused.
TEST(FormulonCApiUsability, ShapeAddedFromASpecTakesTheDefaultsItLeavesOut) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t id = AddShape(wb.handle);
  const fm_shape_info_t shape = ReadShape(wb.handle, id);
  EXPECT_EQ(std::string(shape.preset_geometry), "rect");
  // All-zero paint records: the shape draws its style's fill and outline.
  EXPECT_EQ(shape.fill.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(shape.fill.none, 0);
  EXPECT_EQ(shape.fill.other, 0);
  EXPECT_EQ(shape.line.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(shape.line.has_width_pt, 0);
  EXPECT_EQ(shape.line.none, 0);
  const fm_drawing_object_info_t object = ReadDrawing(wb.handle, id);
  EXPECT_EQ(std::string(object.name), "Shape " + std::to_string(id));
  EXPECT_EQ(object.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(object.anchor.edit_as, FM_ANCHOR_EDIT_AS_TWO_CELL);
  EXPECT_EQ(object.hidden, 0);

  const fm_shape_spec_t absolute = ShapeSpec(FM_ANCHOR_KIND_ABSOLUTE);
  std::uint32_t other = 0;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &absolute, &other), FM_STATUS_INVALID_ARGUMENT);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

// A zeroed anchor kind asks for the operation's default: one-cell for a
// picture, two-cell for a chart and a shape.
TEST(FormulonCApiUsability, ADefaultAnchorKindTakesEachOperationsDefault) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(4, 4);
  const fm_drawing_placement_spec_t picture_opts = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  std::uint32_t picture = 0;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &picture_opts, &picture), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(ReadDrawing(wb.handle, picture).anchor.kind, FM_ANCHOR_KIND_ONE_CELL);

  FillChartSource(wb.handle);
  fm_chart_spec_t chart_opts = ChartOptions();
  chart_opts.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  chart_opts.placement.anchor.from_row = 6;
  std::uint32_t chart = 0;
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &chart_opts, &chart), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t chart_object = ReadDrawing(wb.handle, chart);
  EXPECT_EQ(chart_object.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_GT(chart_object.anchor.to_row, chart_object.anchor.from_row);

  const fm_shape_spec_t shape_spec = ShapeSpec(FM_ANCHOR_KIND_DEFAULT);
  std::uint32_t shape = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &shape_spec, &shape), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t shape_object = ReadDrawing(wb.handle, shape);
  EXPECT_EQ(shape_object.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(shape_object.anchor.to_row, 4U);
  EXPECT_EQ(shape_object.anchor.to_col, 3U);
}

// One drawing patch edits the common fields of every kind; an anchor patch
// merges into the current anchor and keeps the kind it does not name.
TEST(FormulonCApiUsability, DrawingPatchMergesTheAnchorAndSetsTheCommonFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillChartSource(wb.handle);
  const fm_chart_spec_t opts = ChartOptions();
  std::uint32_t chart = 0;
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &opts, &chart), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t before = ReadDrawing(wb.handle, chart);
  ASSERT_GT(before.anchor.to_row, 7U);

  fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.anchor.from_row_op = FM_PATCH_OP_SET;
  patch.anchor.from_row = 7;
  patch.description_op = FM_PATCH_OP_SET;
  patch.description = "sales";
  patch.hidden_op = FM_PATCH_OP_SET;
  patch.hidden = 1;
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, chart, &patch), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t after = ReadDrawing(wb.handle, chart);
  EXPECT_EQ(after.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(after.anchor.from_row, 7U);
  EXPECT_EQ(after.anchor.from_col, before.anchor.from_col);
  EXPECT_EQ(after.anchor.to_row, before.anchor.to_row);
  EXPECT_EQ(after.anchor.to_col, before.anchor.to_col);
  EXPECT_EQ(std::string(after.description), "sales");
  EXPECT_EQ(after.hidden, 1);
  EXPECT_EQ(std::string(after.name), std::string(before.name));

  const std::uint32_t shape = AddShape(wb.handle);
  fm_drawing_patch_t rename = FM_STRUCT_INIT(fm_drawing_patch_t);
  rename.name_op = FM_PATCH_OP_SET;
  rename.name = "Box";
  rename.anchor.to_col_op = FM_PATCH_OP_SET;
  rename.anchor.to_col = 6;
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, shape, &rename), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t box = ReadDrawing(wb.handle, shape);
  EXPECT_EQ(std::string(box.name), "Box");
  EXPECT_EQ(box.anchor.from_row, 1U);
  EXPECT_EQ(box.anchor.from_col, 1U);
  EXPECT_EQ(box.anchor.to_col, 6U);

  fm_drawing_patch_t size_on_two_cell = FM_STRUCT_INIT(fm_drawing_patch_t);
  size_on_two_cell.anchor.width_emu_op = FM_PATCH_OP_SET;
  size_on_two_cell.anchor.width_emu = 9525;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, shape, &size_on_two_cell), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(ReadDrawing(wb.handle, shape).anchor.to_col, 6U);
}

// The one drawing remover takes a shape as it takes a picture or chart.
TEST(FormulonCApiUsability, RemoveDrawingRemovesAShape) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t shape = AddShape(wb.handle);
  ASSERT_EQ(fm_sheet_remove_drawing(wb.handle, 0, shape), 0) << fm_get_last_error_message();
  std::size_t count = 99;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 0, shape), FM_STATUS_NOT_FOUND);
}

// A sparkline group is addressed by the id its add returns: a patch clears a
// clearable bound, keeps the sparklines it does not name, and refuses CLEAR on
// any other field; a removed id is not found again.
TEST(FormulonCApiUsability, SparklineGroupAddressedByIdClearsAManualBoundAndKeepsTheRest) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sparkline_group_t group = FM_STRUCT_INIT(fm_sparkline_group_t);
  group.line_width_pt = 0.75;
  group.min_axis_type = FM_SPARKLINE_AXIS_TYPE_CUSTOM;
  group.has_manual_min = 1;
  group.manual_min = -5.0;
  const fm_sparkline_t items[] = {{"Sheet1!A1:C1", "D1"}, {"Sheet1!A2:C2", "D2"}};
  const char* added = nullptr;
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &group, items, 2, &added), 0) << fm_get_last_error_message();
  ASSERT_NE(added, nullptr);
  const std::string id = added;

  fm_sparkline_group_patch_t patch = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  patch.manual_min_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, id.c_str(), &patch), 0) << fm_get_last_error_message();
  fm_sparkline_group_t read = FM_STRUCT_INIT(fm_sparkline_group_t);
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, id.c_str(), &present, &read), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  EXPECT_EQ(std::string(read.id), id);
  EXPECT_EQ(read.has_manual_min, 0);
  EXPECT_EQ(read.sparkline_count, 2U);
  EXPECT_EQ(read.date_axis_ref, nullptr);

  fm_sparkline_group_patch_t bad = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  bad.line_width_pt_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, id.c_str(), &bad), FM_STATUS_INVALID_ARGUMENT);
  bad = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  bad.items_op = FM_PATCH_OP_SET;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, id.c_str(), &bad), FM_STATUS_INVALID_ARGUMENT);
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(wb.handle, 0, 0, &read), 0) << fm_get_last_error_message();
  EXPECT_EQ(std::string(read.id), id);
  EXPECT_EQ(read.sparkline_count, 2U);

  ASSERT_EQ(fm_sheet_remove_sparkline_group(wb.handle, 0, id.c_str()), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_remove_sparkline_group(wb.handle, 0, id.c_str()), FM_STATUS_NOT_FOUND);
  present = 99;
  EXPECT_EQ(fm_sheet_get_sparkline_group(wb.handle, 0, id.c_str(), &present, &read), 0);
  EXPECT_EQ(present, 0);
}

// A legend patch that only sets the hidden entries keeps the position and the
// overlay; neither of those can be cleared.
TEST(FormulonCApiUsability, LegendPatchOfHiddenEntriesKeepsThePositionAndOverlay) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillChartSource(wb.handle);
  fm_chart_spec_t opts = ChartOptions();
  opts.legend = FM_CHART_LEGEND_POSITION_TOP;
  std::uint32_t chart = 0;
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &opts, &chart), 0) << fm_get_last_error_message();

  const std::uint32_t hidden[] = {0U};
  fm_chart_legend_patch_t patch = FM_STRUCT_INIT(fm_chart_legend_patch_t);
  patch.hidden_entries_op = FM_PATCH_OP_SET;
  patch.hidden_entries = hidden;
  patch.hidden_entries_count = 1;
  ASSERT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, chart, &patch), 0) << fm_get_last_error_message();
  fm_chart_info_t info = ReadChartInfo(wb.handle, chart);
  EXPECT_EQ(info.legend_position, FM_CHART_LEGEND_POSITION_TOP);
  EXPECT_EQ(info.legend_overlay, 0);
  EXPECT_EQ(info.legend_hidden_count, 1U);

  for (int field = 0; field < 2; ++field) {
    fm_chart_legend_patch_t bad = FM_STRUCT_INIT(fm_chart_legend_patch_t);
    (field == 0 ? bad.position_op : bad.overlay_op) = FM_PATCH_OP_CLEAR;
    EXPECT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, chart, &bad), FM_STATUS_INVALID_ARGUMENT) << field;
  }
  info = ReadChartInfo(wb.handle, chart);
  EXPECT_EQ(info.legend_position, FM_CHART_LEGEND_POSITION_TOP);
}

// An absent optional number reads its has_ flag as 0, never NaN: an axis
// without a minimum, and a text category that has no value.
TEST(FormulonCApiUsability, AbsentChartNumbersReadAsHasFlagsNotNaN) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillChartSource(wb.handle);
  const fm_chart_spec_t opts = ChartOptions();
  std::uint32_t chart = 0;
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &opts, &chart), 0) << fm_get_last_error_message();

  std::int32_t present = 0;
  fm_chart_t* model = nullptr;
  ASSERT_EQ(fm_sheet_get_chart(wb.handle, 0, chart, &present, &model), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  fm_chart_info_t info = FM_STRUCT_INIT(fm_chart_info_t);
  ASSERT_EQ(fm_chart_get_info(model, &info), 0);
  ASSERT_GT(info.axis_count, 0U);
  for (std::size_t i = 0; i < info.axis_count; ++i) {
    fm_chart_axis_info_t axis = FM_STRUCT_INIT(fm_chart_axis_info_t);
    ASSERT_EQ(fm_chart_get_axis_at(model, i, &axis), 0);
    EXPECT_EQ(axis.has_min, 0) << i;
    EXPECT_EQ(axis.has_max, 0) << i;
  }
  fm_chart_destroy(model);

  fm_chart_data_t* data = nullptr;
  ASSERT_EQ(fm_sheet_get_chart_data(wb.handle, 0, chart, &present, &data), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  const char* label = nullptr;
  std::int32_t has_value = 99;
  double value = -1.0;
  ASSERT_EQ(fm_chart_data_get_category_at(data, 0, &label, &has_value, &value), 0);
  EXPECT_EQ(std::string(label), "a");
  EXPECT_EQ(has_value, 0);
  EXPECT_EQ(value, 0.0);
  fm_chart_data_destroy(data);
}

// ---------------------------------------------------------------------------
// Conditional-format rule ids
// ---------------------------------------------------------------------------

// A brace-wrapped GUID with upper-case hex digits, the shape the engine generates.
bool IsUpperGuid(const std::string& id) {
  static const char kShape[] = "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}";
  if (id.size() != sizeof(kShape) - 1U) {
    return false;
  }
  for (std::size_t i = 0; i < id.size(); ++i) {
    const char c = id[i];
    if (kShape[i] != 'X') {
      if (c != kShape[i]) {
        return false;
      }
    } else if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F'))) {
      return false;
    }
  }
  return true;
}

std::string Lower(std::string text) {
  for (char& c : text) {
    if (c >= 'A' && c <= 'F') {
      c = static_cast<char>(c - 'A' + 'a');
    }
  }
  return text;
}

std::string AddCfRule(fm_workbook_t* wb, std::size_t sheet, const fm_cf_rule_t& rule) {
  const char* id = nullptr;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb, sheet, &rule, &id), 0) << fm_get_last_error_message();
  return id != nullptr ? std::string(id) : std::string();
}

// A read rule re-added to its own sheet collides with itself; the copy takes a
// generated id, and the add returns the id actually stored.
TEST(FormulonCApiUsability, CfRuleWithAnExistingIdGetsAFreshId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule = ExpressionRule(&sqref, "A1>0", kRuleId);
  const std::string first = AddCfRule(wb.handle, 0, rule);
  const std::string second = AddCfRule(wb.handle, 0, rule);
  const std::string lower = Lower(kRuleId);
  rule.id = lower.c_str();
  const std::string third = AddCfRule(wb.handle, 0, rule);
  ASSERT_EQ(CfRuleCount(wb.handle), 3U);
  EXPECT_EQ(first, kRuleId);
  EXPECT_TRUE(IsUpperGuid(second)) << second;
  EXPECT_TRUE(IsUpperGuid(third)) << third;
  EXPECT_NE(second, first);
  EXPECT_NE(third, first);
  EXPECT_NE(third, second);

  // The same id on another sheet keeps its id.
  std::size_t other = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", append_position(wb.handle), &other), 0);
  rule.id = kRuleId;
  EXPECT_EQ(AddCfRule(wb.handle, other, rule), kRuleId);
}

TEST(FormulonCApiUsability, CfRuleIdIsGeneratedWhenOmittedAndMustOtherwiseBeAGuid) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  EXPECT_TRUE(IsUpperGuid(AddCfRule(wb.handle, 0, ExpressionRule(&sqref, "A1>0"))));
  EXPECT_TRUE(IsUpperGuid(AddCfRule(wb.handle, 0, ExpressionRule(&sqref, "A1>0", ""))));
  const fm_cf_rule_t bad = ExpressionRule(&sqref, "A1>0", "abc");
  const char* id = nullptr;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &bad, &id), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(CfRuleCount(wb.handle), 2U);
}

// The id an add returns reads, updates and removes the rule, in either case
// of its hex digits; the positional read stays.
TEST(FormulonCApiUsability, CfRuleIsAddressedByItsId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  AddCfRule(wb.handle, 0, ExpressionRule(&sqref, "A1>1"));
  const std::string id = AddCfRule(wb.handle, 0, ExpressionRule(&sqref, "A1>0", kRuleId));
  ASSERT_EQ(id, kRuleId);

  std::int32_t present = 99;
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, Lower(id).c_str(), &present, &out), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  EXPECT_EQ(Str(out.formula1), "=A1>0");
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 1, &out), 0);
  EXPECT_EQ(Str(out.id), kRuleId);

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.formula1_op = FM_PATCH_OP_SET;
  patch.formula1 = "A1>5";
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, id.c_str(), &patch), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, id.c_str(), &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(Str(out.formula1), "=A1>5");

  ASSERT_EQ(fm_sheet_remove_cf_rule(wb.handle, 0, Lower(id).c_str()), 0) << fm_get_last_error_message();
  EXPECT_EQ(CfRuleCount(wb.handle), 1U);
  present = 99;
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, id.c_str(), &present, &out), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_sheet_remove_cf_rule(wb.handle, 0, id.c_str()), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, id.c_str(), &patch), FM_STATUS_NOT_FOUND);
}

TEST(FormulonCApiUsability, CfRuleReadBackFeedsAddWithoutDoublingTheEquals) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  const std::string id = AddCfRule(wb.handle, 0, ExpressionRule(&sqref, "=A1>0", kRuleId));

  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=A1>0");
  // The read record goes back in once its own rule is gone.
  ASSERT_EQ(fm_sheet_remove_cf_rule(wb.handle, 0, id.c_str()), 0);
  EXPECT_EQ(AddCfRule(wb.handle, 0, out), kRuleId);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=A1>0");
  EXPECT_EQ(Str(out.id), kRuleId);
}

// ---------------------------------------------------------------------------
// Protected ranges by name
// ---------------------------------------------------------------------------

TEST(FormulonCApiUsability, ProtectedRangeIsAddressedByItsName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t rect{0, 0, 1, 1};
  fm_protected_range_t range = FM_STRUCT_INIT(fm_protected_range_t);
  range.name = "Edit";
  range.ranges = &rect;
  range.range_count = 1;
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &range, nullptr), 0) << fm_get_last_error_message();
  range.name = "EDIT";
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &range, nullptr), FM_STATUS_ALREADY_EXISTS);

  std::int32_t present = 99;
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  ASSERT_EQ(fm_sheet_get_protected_range(wb.handle, 0, "edit", &present, &out), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  EXPECT_EQ(Str(out.name), "Edit");

  fm_protected_range_patch_t patch = FM_STRUCT_INIT(fm_protected_range_patch_t);
  patch.name_op = FM_PATCH_OP_SET;
  patch.name = "Renamed";
  ASSERT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Edit", &patch), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Edit", &patch), FM_STATUS_NOT_FOUND);
  ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(Str(out.name), "Renamed");

  ASSERT_EQ(fm_sheet_remove_protected_range(wb.handle, 0, "renamed"), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_remove_protected_range(wb.handle, 0, "Renamed"), FM_STATUS_NOT_FOUND);
  present = 99;
  ASSERT_EQ(fm_sheet_get_protected_range(wb.handle, 0, "Renamed", &present, &out), 0);
  EXPECT_EQ(present, 0);
}

// ---------------------------------------------------------------------------
// Keys and indices returned by adds
// ---------------------------------------------------------------------------

TEST(FormulonCApiUsability, AddPersonReturnsTheIdAsStored) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::string lower = Lower(kRuleId);
  const fm_person_t person{sizeof(fm_person_t), lower.c_str(), "Alice", nullptr, nullptr};
  const char* id = nullptr;
  ASSERT_EQ(fm_workbook_add_person(wb.handle, &person, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(Str(id), kRuleId);
  EXPECT_EQ(fm_workbook_add_person(wb.handle, &person, nullptr), FM_STATUS_BINDING_NULL_POINTER);
}

TEST(FormulonCApiUsability, SharedItemAddsReturnTheNewItemsIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0U;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), 0) << fm_get_last_error_message();
  std::size_t field = 99U;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, "Region", &field), 0);
  ASSERT_EQ(field, 0U);
  const std::size_t before = 0U;

  std::size_t index = 99U;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache_id, 0U, "Q", &index), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(index, before);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(wb.handle, cache_id, 0U, 7.0, &index), 0);
  EXPECT_EQ(index, before + 1U);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0U, 45292.0, &index), 0);
  EXPECT_EQ(index, before + 2U);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_bool(wb.handle, cache_id, 0U, 1, &index), 0);
  EXPECT_EQ(index, before + 3U);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_blank(wb.handle, cache_id, 0U, &index), 0);
  EXPECT_EQ(index, before + 4U);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_error(wb.handle, cache_id, 0U, 1, &index), 0);  // #DIV/0!
  EXPECT_EQ(index, before + 5U);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache_id, 0U, "R", nullptr),
            FM_STATUS_BINDING_NULL_POINTER);
}

// ---------------------------------------------------------------------------
// Chart insert title
// ---------------------------------------------------------------------------

// The insert title is the setter's record: no flag is the automatic title,
// `NONE` no title, an empty `TEXT` a refusal.
TEST(FormulonCApiUsability, ChartInsertTitleTakesTheSettersRecord) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillChartSource(wb.handle);
  std::uint32_t id = 0;

  fm_chart_spec_t opts = ChartOptions();
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &opts, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadChartInfo(wb.handle, id).title.kind, FM_CHART_TITLE_KIND_AUTO);

  opts.has_title = 1;
  opts.title = fm_chart_title_t{FM_CHART_TITLE_KIND_NONE, nullptr, nullptr};
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &opts, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadChartInfo(wb.handle, id).title.kind, FM_CHART_TITLE_KIND_NONE);

  opts.title = fm_chart_title_t{FM_CHART_TITLE_KIND_TEXT, "Sales", nullptr};
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &opts, &id), 0) << fm_get_last_error_message();
  // The title text is a view into the chart snapshot, so read it before the snapshot is released.
  std::int32_t present = 0;
  fm_chart_t* chart = nullptr;
  ASSERT_EQ(fm_sheet_get_chart(wb.handle, 0, id, &present, &chart), 0) << fm_get_last_error_message();
  fm_chart_info_t info = FM_STRUCT_INIT(fm_chart_info_t);
  ASSERT_EQ(fm_chart_get_info(chart, &info), 0);
  EXPECT_EQ(info.title.kind, FM_CHART_TITLE_KIND_TEXT);
  EXPECT_EQ(Str(info.title.text), "Sales");
  fm_chart_destroy(chart);

  std::size_t before = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &before), 0);
  opts.title = fm_chart_title_t{FM_CHART_TITLE_KIND_TEXT, "", nullptr};
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &opts, &id), FM_STATUS_INVALID_ARGUMENT);
  std::size_t after = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &after), 0);
  EXPECT_EQ(after, before);
}

// ---------------------------------------------------------------------------
// Formula text: the leading `=`
// ---------------------------------------------------------------------------

TEST(FormulonCApiUsability, CellFormulaTextCarriesTheEqualsInBothNotations) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "1+1"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 1, "=A1*2"), 0) << fm_get_last_error_message();
  const char* text = nullptr;
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_formula(wb.handle, 0, 0, 0, &present, &text), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(Str(text), "=1+1");
  ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 0, 0, &present, &text), 0);
  EXPECT_EQ(Str(text), "=1+1");
  ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 1, 1, &present, &text), 0);
  EXPECT_EQ(Str(text), "=R[-1]C[-1]*2");
  // A cell without a formula reports no formula in either notation.
  ASSERT_EQ(fm_sheet_get_formula(wb.handle, 0, 5, 5, &present, &text), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(text, nullptr);
  ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 5, 5, &present, &text), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(text, nullptr);
}

TEST(FormulonCApiUsability, ArrayFormulaTextCarriesTheEqualsWhicheverWayItWasEntered) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (std::uint32_t r = 0; r < 3; ++r) {
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, r, 0, r + 1.0), 0);
  }
  const fm_cell_rect_t block{0, 2, 2, 2};
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, &block, "A1:A3*2"), 0) << fm_get_last_error_message();
  fm_array_formula_info_t info = FM_STRUCT_INIT(fm_array_formula_info_t);
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 1, 2, &present, &info), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(Str(info.formula), "=A1:A3*2");
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, &block, "=A1:A3*3"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 1, 2, &present, &info), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(Str(info.formula), "=A1:A3*3");
}

TEST(FormulonCApiUsability, CfFormulasCarryTheEqualsWhicheverWayTheyWereEntered) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  const char* id = nullptr;
  fm_cf_rule_t bare = FM_STRUCT_INIT(fm_cf_rule_t);
  bare.type = 1;  // CellIs
  bare.has_op = 1;
  bare.op = FM_COMPARISON_OPERATOR_BETWEEN;
  bare.formula1 = "1";
  bare.formula2 = "=10";
  bare.ranges = &sqref;
  bare.range_count = 1;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &bare, &id), 0) << fm_get_last_error_message();
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=1");
  EXPECT_EQ(Str(out.formula2), "=10");

  // A derived formula reads back with the `=` as well.
  fm_cf_rule_t text = FM_STRUCT_INIT(fm_cf_rule_t);
  text.type = 7;  // ContainsText
  text.text = "ab";
  text.ranges = &sqref;
  text.range_count = 1;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &text, &id), 0) << fm_get_last_error_message();
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, id, &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(Str(out.formula1), "=NOT(ISERROR(SEARCH(\"ab\",A1)))");

  // A doubled `=` is still refused.
  fm_cf_rule_t doubled = ExpressionRule(&sqref, "==A1>0");
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &doubled, &id), FM_STATUS_PARSER_UNEXPECTED_TOKEN);
}

TEST(FormulonCApiUsability, ValidationFormulasCarryTheEqualsWhicheverWayTheyWereEntered) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  const fm_data_validation_t v = WholeBetween(&sqref, "1", "=B1*2");
  const char* id = nullptr;
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, &id), 0) << fm_get_last_error_message();
  fm_data_validation_t out = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=1");
  EXPECT_EQ(Str(out.formula2), "=B1*2");

  // An absent formula reads NULL.
  const fm_data_validation_t one = WholeBetween(&sqref, "5", nullptr);
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &one, &id), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_validation_at(wb.handle, 0, 1, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=5");
  EXPECT_EQ(out.formula2, nullptr);
}

TEST(FormulonCApiUsability, TableTotalsFormulaCarriesTheEqualsWhicheverWayItWasEntered) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Product"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "Amount"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 10.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2, 1, 20.0), 0);
  const fm_cell_rect_t ref{0, 0, 3, 1};
  const char* columns[] = {"Product", "Amount"};
  std::size_t table = 99;
  ASSERT_EQ(AddTable(wb.handle, 0, &ref, "Sales", columns, 2, "TableStyleMedium2", 1, 1, &table), 0)
      << fm_get_last_error_message();
  fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
  ASSERT_EQ(fm_sheet_get_table_column_at(wb.handle, 0, table, 1, &column), 0);
  EXPECT_EQ(column.totals_row_formula, nullptr);

  ASSERT_EQ(fm_sheet_set_table_column_totals(wb.handle, 0, table, 1, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr,
                                             "SUM(Sales[Amount])*2"),
            0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_table_column_at(wb.handle, 0, table, 1, &column), 0);
  EXPECT_EQ(Str(column.totals_row_formula), "=SUM([Amount])*2");
  ASSERT_EQ(fm_sheet_set_table_column_totals(wb.handle, 0, table, 1, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr,
                                             "=SUM(Sales[Amount])*3"),
            0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_table_column_at(wb.handle, 0, table, 1, &column), 0);
  EXPECT_EQ(Str(column.totals_row_formula), "=SUM([Amount])*3");
}

TEST(FormulonCApiUsability, PivotCacheFieldFormulaCarriesTheEqualsWhicheverWayItWasEntered) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache), 0) << fm_get_last_error_message();
  for (const char* name : {"Amt", "Qty"}) {
    std::size_t field = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache, name, &field), 0) << fm_get_last_error_message();
  }
  std::size_t price = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_field(wb.handle, cache, "Price", "Amt/Qty", &price), 0)
      << fm_get_last_error_message();
  std::int32_t present = 0;
  const char* formula = nullptr;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, price, &present, &formula), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(Str(formula), "=Amt/Qty");
  ASSERT_EQ(fm_workbook_set_pivot_cache_calculated_field_formula(wb.handle, cache, price, "=Amt*2"), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, price, &present, &formula), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(Str(formula), "=Amt*2");
}

TEST(FormulonCApiUsability, PivotCacheFieldFormulaIsAbsentForADatabaseField) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache), 0) << fm_get_last_error_message();
  std::size_t field = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache, "Amt", &field), 0) << fm_get_last_error_message();
  std::int32_t present = -1;
  const char* formula = "unchanged";
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, field, &present, &formula), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(formula, nullptr);
}

TEST(FormulonCApiUsability, SaveAndReloadDoesNotDoubleTheEquals) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "A1+1"), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  const fm_cf_rule_t rule = ExpressionRule(&sqref, "=A1>0");
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, ignored_id()), 0) << fm_get_last_error_message();
  const fm_data_validation_t v = WholeBetween(&sqref, "=1", "10");
  const char* validation_id = nullptr;
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, &validation_id), 0) << fm_get_last_error_message();
  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "=0.5"), 0);

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  const char* text = nullptr;
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_formula(reloaded.handle, 0, 0, 1, &present, &text), 0);
  EXPECT_EQ(Str(text), "=A1+1");
  ASSERT_EQ(fm_sheet_get_formula_r1c1(reloaded.handle, 0, 0, 1, &present, &text), 0);
  EXPECT_EQ(Str(text), "=RC[-1]+1");
  fm_cf_rule_t cf = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(reloaded.handle, 0, 0, &cf), 0);
  EXPECT_EQ(Str(cf.formula1), "=A1>0");
  fm_data_validation_t dv = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(reloaded.handle, 0, 0, &dv), 0);
  EXPECT_EQ(Str(dv.formula1), "=1");
  EXPECT_EQ(Str(dv.formula2), "=10");
  fm_defined_name_t name = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(reloaded.handle, 0, &name), 0);
  EXPECT_EQ(Str(name.formula), "=0.5");
}
