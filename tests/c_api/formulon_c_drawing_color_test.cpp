//
// C ABI tests for the paint records of shapes and charts: a DrawingML colour
// with its transforms and rendered result, line ends, the full dash preset
// set, and the chart and plot area paint.

#include <cstddef>
#include <cstdint>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

struct ChartGuard {
  fm_chart_t* handle = nullptr;
  ~ChartGuard() { fm_chart_destroy(handle); }
};

fm_color_t Rgb(uint32_t argb) {
  fm_color_t c{};
  c.kind = FM_COLOR_KIND_RGB;
  c.argb = argb;
  return c;
}

fm_color_transform_t Transform(const char* name, int32_t value) {
  return fm_color_transform_t{name, 1, value};
}

fm_shape_spec_t RectSpec(const char* prst = "rect") {
  fm_shape_spec_t s = FM_STRUCT_INIT(fm_shape_spec_t);
  s.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  s.placement.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  s.placement.anchor.from_row = 1;
  s.placement.anchor.from_col = 1;
  s.placement.anchor.to_row = 5;
  s.placement.anchor.to_col = 4;
  s.preset_geometry = prst;
  return s;
}

uint32_t AddShape(fm_workbook_t* wb, const fm_shape_spec_t& spec) {
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_shape(wb, 0, &spec, &id), FM_STATUS_OK) << fm_get_last_error_message();
  return id;
}

fm_shape_info_t ReadShape(const fm_workbook_t* wb, uint32_t id) {
  fm_shape_info_t out = FM_STRUCT_INIT(fm_shape_info_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_shape(wb, 0, id, &present, &out), FM_STATUS_OK) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return out;
}

void SaveAndReload(const WorkbookGuard& source, WorkbookGuard* out) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(source.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), FM_STATUS_OK)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &out->handle), FM_STATUS_OK)
      << fm_get_last_error_message();
}

/// Sheet1: B1:D1 names, A2:A5 text categories, B2:D5 numbers.
void FillGrid(fm_workbook_t* wb) {
  const char* names[] = {"S1", "S2", "S3"};
  const char* cats[] = {"a", "b", "c", "d"};
  for (std::uint32_t c = 0; c < 3; ++c) {
    ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, c + 1, names[c]), FM_STATUS_OK);
  }
  for (std::uint32_t r = 0; r < 4; ++r) {
    ASSERT_EQ(fm_sheet_set_text(wb, 0, r + 1, 0, cats[r]), FM_STATUS_OK);
    for (std::uint32_t c = 0; c < 3; ++c) {
      ASSERT_EQ(fm_sheet_set_number(wb, 0, r + 1, c + 1, static_cast<double>((r + 1) * (c + 2))), FM_STATUS_OK);
    }
  }
}

std::uint32_t InsertLineChart(fm_workbook_t* wb) {
  fm_chart_spec_t o = FM_STRUCT_INIT(fm_chart_spec_t);
  o.type = FM_CHART_PLOT_TYPE_LINE;
  o.source = "Sheet1!$A$1:$D$5";
  o.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  o.placement.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  o.placement.anchor.from_row = 6;
  std::uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_chart(wb, 0, &o, &id), FM_STATUS_OK) << fm_get_last_error_message();
  return id;
}

void OpenChart(fm_workbook_t* wb, std::uint32_t id, ChartGuard* out) {
  fm_chart_destroy(out->handle);
  out->handle = nullptr;
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_chart(wb, 0, id, &present, &out->handle), FM_STATUS_OK) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
}

fm_chart_series_info_t Series(const ChartGuard& chart, std::size_t series) {
  fm_chart_series_info_t s = FM_STRUCT_INIT(fm_chart_series_info_t);
  EXPECT_EQ(fm_chart_get_series_at(chart.handle, 0, series, &s), FM_STATUS_OK) << fm_get_last_error_message();
  return s;
}

fm_chart_info_t Info(const ChartGuard& chart) {
  fm_chart_info_t info = FM_STRUCT_INIT(fm_chart_info_t);
  EXPECT_EQ(fm_chart_get_info(chart.handle, &info), FM_STATUS_OK);
  return info;
}

fm_status_t UpdateFill(fm_workbook_t* wb, uint32_t id, const fm_chart_fill_patch_t& fill) {
  fm_shape_patch_t patch = FM_STRUCT_INIT(fm_shape_patch_t);
  patch.fill = fill;
  return fm_sheet_update_shape(wb, 0, id, nullptr, &patch);
}

TEST(FormulonCApiDrawingColor, EnumOrdinals) {
  EXPECT_EQ(FM_CHART_LINE_END_TYPE_NONE, 0);
  EXPECT_EQ(FM_CHART_LINE_END_TYPE_ARROW, 5);
  EXPECT_EQ(FM_CHART_LINE_END_SIZE_MEDIUM, 0);
  EXPECT_EQ(FM_CHART_LINE_END_SIZE_SMALL, 1);
  EXPECT_EQ(FM_CHART_LINE_END_SIZE_LARGE, 2);
  EXPECT_EQ(FM_CHART_LINE_DASH_OTHER, 4);
  EXPECT_EQ(FM_CHART_LINE_DASH_SYS_DOT, 5);
  EXPECT_EQ(FM_CHART_LINE_DASH_LG_DASH_DOT_DOT, 9);
  EXPECT_EQ(FM_CHART_LINE_DASH_SYS_DASH_DOT_DOT, 11);
}

// A colour's transforms are written in order and read back with the colour Excel renders.
TEST(FormulonCApiDrawingColor, TransformsRoundTripWithTheRenderedColour) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const fm_color_transform_t transforms[] = {Transform("lumMod", 50000)};
  fm_shape_spec_t spec = RectSpec();
  spec.fill.color.base = Rgb(0xFF336699U);
  spec.fill.color.transforms = transforms;
  spec.fill.color.transform_count = 1;
  // The rendered fields are ignored on input.
  spec.fill.color.has_resolved = 1;
  spec.fill.color.resolved = 0xFF000000U;
  const uint32_t id = AddShape(wb.handle, spec);

  WorkbookGuard again;
  SaveAndReload(wb, &again);
  for (const fm_workbook_t* book :
       {static_cast<const fm_workbook_t*>(wb.handle), static_cast<const fm_workbook_t*>(again.handle)}) {
    const fm_shape_info_t out = ReadShape(book, id);
    EXPECT_EQ(out.fill.color.base.kind, FM_COLOR_KIND_RGB);
    EXPECT_EQ(out.fill.color.base.argb & 0xFFFFFFU, 0x336699U);
    ASSERT_EQ(out.fill.color.transform_count, 1U);
    EXPECT_STREQ(out.fill.color.transforms[0].name, "lumMod");
    EXPECT_EQ(out.fill.color.transforms[0].has_value, 1);
    EXPECT_EQ(out.fill.color.transforms[0].value, 50000);
    ASSERT_EQ(out.fill.color.has_resolved, 1);
    EXPECT_EQ(out.fill.color.resolved >> 24, 0xFFU);
    EXPECT_NE(out.fill.color.resolved, 0xFF336699U);
    EXPECT_NE(out.fill.color.resolved, 0xFF000000U);
    // No alpha transform: the colour renders opaque.
    EXPECT_EQ(out.fill.color.has_resolved_transparency_pct, 1);
    EXPECT_DOUBLE_EQ(out.fill.color.resolved_transparency_pct, 0.0);
  }
}

// A plain colour reads without transforms and renders as itself; no colour reads base None.
TEST(FormulonCApiDrawingColor, PlainColourAndAbsentColour) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  fm_shape_spec_t spec = RectSpec();
  spec.fill.color.base = Rgb(0xFF102030U);
  const uint32_t id = AddShape(wb.handle, spec);
  const fm_shape_info_t out = ReadShape(wb.handle, id);
  EXPECT_EQ(out.fill.color.transform_count, 0U);
  EXPECT_EQ(out.fill.color.transforms, nullptr);
  ASSERT_EQ(out.fill.color.has_resolved, 1);
  EXPECT_EQ(out.fill.color.resolved, 0xFF102030U);
  EXPECT_EQ(out.line.color.base.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(out.line.color.has_resolved, 0);
  EXPECT_EQ(out.line.color.has_resolved_transparency_pct, 0);
}

// A transform the engine does not model keeps the colour unrendered.
TEST(FormulonCApiDrawingColor, UnmodelledTransformLeavesTheRenderedColourUnset) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const fm_color_transform_t transforms[] = {Transform("fooBar", 7)};
  fm_shape_spec_t spec = RectSpec();
  spec.fill.color.base = Rgb(0xFF336699U);
  spec.fill.color.transforms = transforms;
  spec.fill.color.transform_count = 1;
  const uint32_t id = AddShape(wb.handle, spec);
  const fm_shape_info_t out = ReadShape(wb.handle, id);
  ASSERT_EQ(out.fill.color.transform_count, 1U);
  EXPECT_STREQ(out.fill.color.transforms[0].name, "fooBar");
  EXPECT_EQ(out.fill.color.has_resolved, 0);
  EXPECT_EQ(out.fill.color.has_resolved_transparency_pct, 0);
}

// alphaMod and alphaOff scale and offset the rendered opacity in order; the colour is unchanged.
TEST(FormulonCApiDrawingColor, AlphaModAndOffRenderTheTransparency) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const fm_color_transform_t transforms[] = {Transform("alphaMod", 50000), Transform("alphaOff", -20000)};
  fm_shape_spec_t spec = RectSpec();
  spec.fill.color.base = Rgb(0xFF336699U);
  spec.fill.color.transforms = transforms;
  spec.fill.color.transform_count = 2;
  const uint32_t id = AddShape(wb.handle, spec);
  const fm_shape_info_t out = ReadShape(wb.handle, id);
  ASSERT_EQ(out.fill.color.has_resolved, 1);
  EXPECT_EQ(out.fill.color.resolved, 0xFF336699U);
  ASSERT_EQ(out.fill.color.has_resolved_transparency_pct, 1);
  EXPECT_DOUBLE_EQ(out.fill.color.resolved_transparency_pct, 70.0);
}

// An `alpha` transform and `transparency_pct` in one write must agree; either alone states the opacity.
TEST(FormulonCApiDrawingColor, AlphaTransformAndTransparencyMustAgree) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const uint32_t id = AddShape(wb.handle, RectSpec());
  const fm_color_transform_t alpha[] = {Transform("alpha", 40000)};

  fm_chart_fill_patch_t both{};
  both.color_op = FM_PATCH_OP_SET;
  both.color.base = Rgb(0xFF336699U);
  both.color.transforms = alpha;
  both.color.transform_count = 1;
  both.transparency_pct_op = FM_PATCH_OP_SET;
  both.transparency_pct = 50.0;
  EXPECT_EQ(UpdateFill(wb.handle, id, both), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(ReadShape(wb.handle, id).fill.color.base.kind, FM_COLOR_KIND_NONE);

  both.transparency_pct = 60.0;
  ASSERT_EQ(UpdateFill(wb.handle, id, both), FM_STATUS_OK) << fm_get_last_error_message();
  fm_shape_info_t out = ReadShape(wb.handle, id);
  EXPECT_EQ(out.fill.has_transparency_pct, 1);
  EXPECT_DOUBLE_EQ(out.fill.transparency_pct, 60.0);
  ASSERT_EQ(out.fill.color.transform_count, 1U);
  EXPECT_STREQ(out.fill.color.transforms[0].name, "alpha");
  EXPECT_EQ(out.fill.color.transforms[0].value, 40000);

  fm_chart_fill_patch_t alone{};
  alone.color_op = FM_PATCH_OP_SET;
  alone.color.base = Rgb(0xFF336699U);
  const fm_color_transform_t quarter[] = {Transform("alpha", 75000)};
  alone.color.transforms = quarter;
  alone.color.transform_count = 1;
  ASSERT_EQ(UpdateFill(wb.handle, id, alone), FM_STATUS_OK) << fm_get_last_error_message();
  out = ReadShape(wb.handle, id);
  EXPECT_EQ(out.fill.has_transparency_pct, 1);
  EXPECT_DOUBLE_EQ(out.fill.transparency_pct, 25.0);
}

// Colour input is checked before anything is written.
TEST(FormulonCApiDrawingColor, MalformedColourInputIsRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const uint32_t id = AddShape(wb.handle, RectSpec());

  const auto with = [](const fm_color_transform_t* transforms, uint32_t count) {
    fm_chart_fill_patch_t p{};
    p.color_op = FM_PATCH_OP_SET;
    p.color.base = Rgb(0xFF336699U);
    p.color.transforms = transforms;
    p.color.transform_count = count;
    return p;
  };
  // NULL transforms with a count.
  EXPECT_EQ(UpdateFill(wb.handle, id, with(nullptr, 1)), FM_STATUS_INVALID_ARGUMENT);
  // A transform without a name, with a malformed name, or with a has_value outside 0..1.
  const fm_color_transform_t unnamed[] = {{nullptr, 1, 50000}};
  EXPECT_EQ(UpdateFill(wb.handle, id, with(unnamed, 1)), FM_STATUS_INVALID_ARGUMENT);
  const fm_color_transform_t malformed[] = {Transform("1lumMod", 50000)};
  EXPECT_EQ(UpdateFill(wb.handle, id, with(malformed, 1)), FM_STATUS_INVALID_ARGUMENT);
  const fm_color_transform_t empty[] = {Transform("", 50000)};
  EXPECT_EQ(UpdateFill(wb.handle, id, with(empty, 1)), FM_STATUS_INVALID_ARGUMENT);
  const fm_color_transform_t flag[] = {{"lumMod", 2, 50000}};
  EXPECT_EQ(UpdateFill(wb.handle, id, with(flag, 1)), FM_STATUS_INVALID_ARGUMENT);
  const fm_color_transform_t bad_utf8[] = {Transform("lum\xC3", 50000)};
  EXPECT_EQ(UpdateFill(wb.handle, id, with(bad_utf8, 1)), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  // More transforms than a colour may carry.
  const std::vector<fm_color_transform_t> many(65, Transform("lumMod", 50000));
  EXPECT_EQ(UpdateFill(wb.handle, id, with(many.data(), 65)), FM_STATUS_SECURITY_RESOURCE_LIMIT);
  // A tint is the shorthand for transforms and cannot sit beside them.
  const fm_color_transform_t lum[] = {Transform("lumMod", 50000)};
  fm_chart_fill_patch_t tinted = with(lum, 1);
  tinted.color.base = fm_color_t{};
  tinted.color.base.kind = FM_COLOR_KIND_THEME;
  tinted.color.base.theme = 4;
  tinted.color.base.tint = 0.4;
  EXPECT_EQ(UpdateFill(wb.handle, id, tinted), FM_STATUS_INVALID_ARGUMENT);
  // A set colour needs a base.
  fm_chart_fill_patch_t baseless = with(nullptr, 0);
  baseless.color.base = fm_color_t{};
  EXPECT_EQ(UpdateFill(wb.handle, id, baseless), FM_STATUS_INVALID_ARGUMENT);

  EXPECT_EQ(ReadShape(wb.handle, id).fill.color.base.kind, FM_COLOR_KIND_NONE);
}

// Only CLEAR removes a colour: a SET whose base is None is refused on a shape outline and on chart label text.
TEST(FormulonCApiDrawingColor, SetColourWithoutBaseIsRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  fm_shape_spec_t spec = RectSpec();
  spec.line.color.base = Rgb(0xFFFF0000U);
  const uint32_t id = AddShape(wb.handle, spec);
  fm_shape_patch_t line = FM_STRUCT_INIT(fm_shape_patch_t);
  line.line.color_op = FM_PATCH_OP_SET;
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, id, nullptr, &line), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(ReadShape(wb.handle, id).line.color.base.argb, 0xFFFF0000U);

  FillGrid(wb.handle);
  const std::uint32_t chart_id = InsertLineChart(wb.handle);
  fm_chart_series_patch_t labels = FM_STRUCT_INIT(fm_chart_series_patch_t);
  labels.labels_op = FM_PATCH_OP_SET;
  labels.labels.show_value_op = FM_PATCH_OP_SET;
  labels.labels.show_value = 1;
  labels.labels.font_color_op = FM_PATCH_OP_SET;
  labels.labels.font_color = Rgb(0xFFFF0000U);
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, chart_id, 0, 0, &labels), FM_STATUS_OK)
      << fm_get_last_error_message();
  labels.labels.font_color = fm_color_t{};
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, chart_id, 0, 0, &labels), FM_STATUS_INVALID_ARGUMENT);
  ChartGuard chart;
  OpenChart(wb.handle, chart_id, &chart);
  EXPECT_EQ(Series(chart, 0).labels.font_color.argb, 0xFFFF0000U);
}

// Line ends are read, set whole, cleared and range-checked.
TEST(FormulonCApiDrawingColor, LineEndsReadSetClearAndRefuse) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  fm_shape_spec_t spec = RectSpec("straightConnector1");
  spec.line.color.base = Rgb(0xFF000000U);
  spec.line.has_head_end = 1;
  spec.line.head_end = fm_chart_line_end_t{FM_CHART_LINE_END_TYPE_ARROW, 1, FM_CHART_LINE_END_SIZE_LARGE, 0, 0};
  spec.line.has_tail_end = 1;
  spec.line.tail_end = fm_chart_line_end_t{FM_CHART_LINE_END_TYPE_OVAL, 0, 0, 1, FM_CHART_LINE_END_SIZE_SMALL};
  const uint32_t id = AddShape(wb.handle, spec);

  WorkbookGuard again;
  SaveAndReload(wb, &again);
  fm_shape_info_t out = ReadShape(again.handle, id);
  ASSERT_EQ(out.line.has_head_end, 1);
  EXPECT_EQ(out.line.head_end.type, FM_CHART_LINE_END_TYPE_ARROW);
  EXPECT_EQ(out.line.head_end.has_width, 1);
  EXPECT_EQ(out.line.head_end.width, FM_CHART_LINE_END_SIZE_LARGE);
  EXPECT_EQ(out.line.head_end.has_length, 0);
  ASSERT_EQ(out.line.has_tail_end, 1);
  EXPECT_EQ(out.line.tail_end.type, FM_CHART_LINE_END_TYPE_OVAL);
  EXPECT_EQ(out.line.tail_end.has_width, 0);
  EXPECT_EQ(out.line.tail_end.has_length, 1);
  EXPECT_EQ(out.line.tail_end.length, FM_CHART_LINE_END_SIZE_SMALL);

  // A set end replaces the stored one whole; CLEAR removes it.
  fm_shape_patch_t patch = FM_STRUCT_INIT(fm_shape_patch_t);
  patch.line.head_end_op = FM_PATCH_OP_CLEAR;
  patch.line.tail_end_op = FM_PATCH_OP_SET;
  patch.line.tail_end = fm_chart_line_end_t{FM_CHART_LINE_END_TYPE_TRIANGLE, 0, 0, 0, 0};
  ASSERT_EQ(fm_sheet_update_shape(again.handle, 0, id, nullptr, &patch), FM_STATUS_OK) << fm_get_last_error_message();
  out = ReadShape(again.handle, id);
  EXPECT_EQ(out.line.has_head_end, 0);
  ASSERT_EQ(out.line.has_tail_end, 1);
  EXPECT_EQ(out.line.tail_end.type, FM_CHART_LINE_END_TYPE_TRIANGLE);
  EXPECT_EQ(out.line.tail_end.has_length, 0);

  const fm_chart_line_end_t bad[] = {
      {6, 0, 0, 0, 0},
      {-1, 0, 0, 0, 0},
      {FM_CHART_LINE_END_TYPE_ARROW, 2, 0, 0, 0},
      {FM_CHART_LINE_END_TYPE_ARROW, 1, 3, 0, 0},
      {FM_CHART_LINE_END_TYPE_ARROW, 0, 0, 1, -1},
  };
  for (const fm_chart_line_end_t& end : bad) {
    fm_shape_patch_t p = FM_STRUCT_INIT(fm_shape_patch_t);
    p.line.head_end_op = FM_PATCH_OP_SET;
    p.line.head_end = end;
    EXPECT_EQ(fm_sheet_update_shape(again.handle, 0, id, nullptr, &p), FM_STATUS_INVALID_ARGUMENT)
        << end.type << " " << end.has_width << " " << end.width << " " << end.has_length << " " << end.length;
    fm_shape_spec_t s = RectSpec("straightConnector1");
    s.line.has_tail_end = 1;
    s.line.tail_end = end;
    uint32_t ignored = 0;
    EXPECT_EQ(fm_sheet_add_shape(again.handle, 0, &s, &ignored), FM_STATUS_INVALID_ARGUMENT);
  }
  fm_shape_spec_t flag = RectSpec("straightConnector1");
  flag.line.has_head_end = 2;
  uint32_t ignored = 0;
  EXPECT_EQ(fm_sheet_add_shape(again.handle, 0, &flag, &ignored), FM_STATUS_INVALID_ARGUMENT);
}

// Every preset dash reads as itself; Other stays a read-only value and 12 is out of range.
TEST(FormulonCApiDrawingColor, EveryPresetDashRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertLineChart(wb.handle);
  for (int32_t dash = FM_CHART_LINE_DASH_SOLID; dash <= FM_CHART_LINE_DASH_SYS_DASH_DOT_DOT; ++dash) {
    fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
    patch.line.dash_op = FM_PATCH_OP_SET;
    patch.line.dash = dash;
    const fm_status_t rc = fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch);
    if (dash == FM_CHART_LINE_DASH_OTHER) {
      EXPECT_EQ(rc, FM_STATUS_INVALID_ARGUMENT);
      continue;
    }
    ASSERT_EQ(rc, FM_STATUS_OK) << dash << " " << fm_get_last_error_message();
    ChartGuard chart;
    OpenChart(wb.handle, id, &chart);
    const fm_chart_series_info_t s = Series(chart, 0);
    EXPECT_EQ(s.line.has_dash, 1);
    EXPECT_EQ(s.line.dash, dash);
  }
  fm_chart_series_patch_t past = FM_STRUCT_INIT(fm_chart_series_patch_t);
  past.line.dash_op = FM_PATCH_OP_SET;
  past.line.dash = FM_CHART_LINE_DASH_SYS_DASH_DOT_DOT + 1;
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &past), FM_STATUS_INVALID_ARGUMENT);
}

// A chart read's transform views stay valid, and stable, until the handle is destroyed.
TEST(FormulonCApiDrawingColor, ChartTransformViewsLiveAsLongAsTheHandle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertLineChart(wb.handle);
  const fm_color_transform_t transforms[] = {Transform("lumMod", 75000), Transform("lumOff", 25000)};
  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.line.color_op = FM_PATCH_OP_SET;
  patch.line.color.base.kind = FM_COLOR_KIND_THEME;
  patch.line.color.base.theme = 5;
  patch.line.color.transforms = transforms;
  patch.line.color.transform_count = 2;
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_OK) << fm_get_last_error_message();

  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  const fm_chart_series_info_t first = Series(chart, 0);
  const fm_chart_series_info_t other = Series(chart, 1);
  const fm_chart_series_info_t again = Series(chart, 0);
  (void)other;
  ASSERT_EQ(first.line.color.transform_count, 2U);
  EXPECT_EQ(first.line.color.transforms, again.line.color.transforms);
  EXPECT_EQ(first.line.color.base.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(first.line.color.base.theme, 5U);
  EXPECT_STREQ(first.line.color.transforms[0].name, "lumMod");
  EXPECT_EQ(first.line.color.transforms[0].value, 75000);
  EXPECT_STREQ(first.line.color.transforms[1].name, "lumOff");
  EXPECT_EQ(first.line.color.transforms[1].value, 25000);
  EXPECT_EQ(first.line.color.has_resolved, 1);
}

// The chart and plot area paint is read, set through the display patch, kept by KEEP and cleared by CLEAR.
TEST(FormulonCApiDrawingColor, ChartAndPlotAreaPaint) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertLineChart(wb.handle);

  fm_chart_display_patch_t patch = FM_STRUCT_INIT(fm_chart_display_patch_t);
  patch.chart_area.fill.color_op = FM_PATCH_OP_SET;
  patch.chart_area.fill.color.base = Rgb(0xFFF0E0D0U);
  patch.plot_area.line.width_pt_op = FM_PATCH_OP_SET;
  patch.plot_area.line.width_pt = 2.0;
  patch.plot_area.line.dash_op = FM_PATCH_OP_SET;
  patch.plot_area.line.dash = FM_CHART_LINE_DASH_LG_DASH;
  patch.plot_area.line.color_op = FM_PATCH_OP_SET;
  patch.plot_area.line.color.base = Rgb(0xFF203040U);
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &patch), FM_STATUS_OK) << fm_get_last_error_message();

  WorkbookGuard again;
  SaveAndReload(wb, &again);
  ChartGuard chart;
  OpenChart(again.handle, id, &chart);
  fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.chart_area.fill.color.base.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(info.chart_area.fill.color.base.argb & 0xFFFFFFU, 0xF0E0D0U);
  EXPECT_EQ(info.plot_area.line.color.base.argb & 0xFFFFFFU, 0x203040U);
  EXPECT_EQ(info.plot_area.line.has_width_pt, 1);
  EXPECT_DOUBLE_EQ(info.plot_area.line.width_pt, 2.0);
  EXPECT_EQ(info.plot_area.line.dash, FM_CHART_LINE_DASH_LG_DASH);

  // An all-KEEP patch changes nothing.
  const fm_chart_display_patch_t keep = FM_STRUCT_INIT(fm_chart_display_patch_t);
  ASSERT_EQ(fm_sheet_update_chart_display(again.handle, 0, id, &keep), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(again.handle, id, &chart);
  EXPECT_EQ(Info(chart).chart_area.fill.color.base.argb & 0xFFFFFFU, 0xF0E0D0U);

  // Clearing every member removes the paint the file states.
  fm_chart_display_patch_t clear = FM_STRUCT_INIT(fm_chart_display_patch_t);
  for (fm_chart_area_paint_patch_t* area : {&clear.chart_area, &clear.plot_area}) {
    area->fill.color_op = area->fill.transparency_pct_op = area->fill.none_op = FM_PATCH_OP_CLEAR;
    area->line.color_op = area->line.transparency_pct_op = area->line.width_pt_op = FM_PATCH_OP_CLEAR;
    area->line.dash_op = area->line.none_op = area->line.head_end_op = area->line.tail_end_op = FM_PATCH_OP_CLEAR;
  }
  ASSERT_EQ(fm_sheet_update_chart_display(again.handle, 0, id, &clear), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(again.handle, id, &chart);
  info = Info(chart);
  EXPECT_EQ(info.chart_area.fill.color.base.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(info.plot_area.line.color.base.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(info.plot_area.line.has_width_pt, 0);
  EXPECT_EQ(info.plot_area.line.has_dash, 0);

  // A member that fails its rule refuses the whole patch.
  const fm_chart_info_t before = Info(chart);
  fm_chart_display_patch_t bad = FM_STRUCT_INIT(fm_chart_display_patch_t);
  bad.style_op = FM_PATCH_OP_SET;
  bad.style = 3;
  bad.plot_area.fill.color_op = FM_PATCH_OP_SET;
  bad.plot_area.fill.color.base.kind = FM_COLOR_KIND_INDEXED;
  EXPECT_EQ(fm_sheet_update_chart_display(again.handle, 0, id, &bad), FM_STATUS_INVALID_ARGUMENT);
  bad.plot_area.fill.color_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_update_chart_display(again.handle, 0, id, &bad), FM_STATUS_INVALID_ARGUMENT);
  OpenChart(again.handle, id, &chart);
  EXPECT_EQ(Info(chart).has_style, before.has_style);
  EXPECT_EQ(Info(chart).style, before.style);
}

}  // namespace
