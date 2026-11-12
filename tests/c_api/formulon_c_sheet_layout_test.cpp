//
// Stable C ABI smoke tests for the sheet view / layout entry points
// added in `formulon_c.h`. The driver is C++ for gtest convenience; the
// layout mutations and observations stay on the C surface (the core model is
// used only to seed a styled input fixture for the overlay regression).

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <iterator>
#include <limits>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "formulon_c_out_index.h"
#include "formulon_c_sheet_view_patch.h"
#include "gtest/gtest.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

static_assert(offsetof(fm_column_layout_info_t, has_width_chars) == 36U);
static_assert(offsetof(fm_column_layout_info_t, has_xf_index) == 40U);
static_assert(offsetof(fm_column_layout_info_t, xf_index) == 44U);
static_assert(sizeof(fm_column_layout_info_t) == 48U);
static_assert(offsetof(fm_row_layout_info_t, has_xf_index) == 28U);
static_assert(offsetof(fm_row_layout_info_t, xf_index) == 32U);
static_assert(offsetof(fm_row_layout_info_t, has_height_pt) == 36U);
static_assert(offsetof(fm_row_layout_info_t, custom_height) == 40U);
static_assert(sizeof(fm_row_layout_info_t) == 48U);
// `fm_sheet_get_view` writes through a caller-supplied pointer, so a caller
// compiled against a narrower definition of this struct is overwritten past
// the end of its own storage rather than getting a diagnosable error. Pin the
// size here: changing the struct must break the build and be acknowledged,
// not merely be reflected in the bindings afterwards.
static_assert(sizeof(fm_sheet_view_info_t) == 80U);
static_assert(offsetof(fm_sheet_view_info_t, tab_selected) == 32U);
static_assert(offsetof(fm_sheet_view_info_t, view_mode) == 36U);
static_assert(offsetof(fm_sheet_view_info_t, top_left_row) == 40U);

/// The status every documented rejection on this surface returns. Pinning
/// the value rather than "non-zero" is what makes the assertions below
/// distinguish a refused argument from an unrelated failure that happens to
/// be non-zero too.
constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();
constexpr double kInfinity = std::numeric_limits<double>::infinity();

/// Last addressable column / row of Excel's grid. Spelled from the model
/// constants so widening the grid widens these with it.
constexpr std::uint32_t kLastColumn = formulon::Sheet::kMaxCols - 1U;
constexpr std::uint32_t kLastRow = formulon::Sheet::kMaxRows - 1U;

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

TEST(FormulonCApiSheetLayout, GetViewDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_view_info_t v = FM_STRUCT_INIT(fm_sheet_view_info_t);
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.zoom_pct, 100U);
  EXPECT_EQ(v.freeze_rows, 0U);
  EXPECT_EQ(v.freeze_cols, 0U);
  EXPECT_EQ(v.show_grid_lines, 1);
  EXPECT_EQ(v.show_row_column_headers, 1);
  EXPECT_EQ(v.show_zeros, 1);
  EXPECT_EQ(v.right_to_left, 0);
  EXPECT_EQ(v.tab_selected, 1);  // a new workbook's only sheet is its selected tab
  EXPECT_EQ(v.view_mode, FM_SHEET_VIEW_MODE_NORMAL);
}

TEST(FormulonCApiSheetLayout, NewWorkbookSelectsOnlyItsFirstTab) {
  for (const char* profile : {static_cast<const char*>(nullptr), "mac-365-ja_JP"}) {
    SCOPED_TRACE(profile == nullptr ? "default" : profile);
    WorkbookGuard wb;
    ASSERT_EQ(profile == nullptr ? fm_workbook_create_default(&wb.handle)
                                 : fm_workbook_create_for_profile(profile, &wb.handle),
              0);
    ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second", 1U, ignored_index()), 0);
    fm_sheet_view_info_t first = FM_STRUCT_INIT(fm_sheet_view_info_t);
    fm_sheet_view_info_t second = FM_STRUCT_INIT(fm_sheet_view_info_t);
    ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &first), 0);
    ASSERT_EQ(fm_sheet_get_view(wb.handle, 1, &second), 0);
    EXPECT_EQ(first.tab_selected, 1);
    EXPECT_EQ(second.tab_selected, 0);
  }
}

TEST(FormulonCApiSheetLayout, ViewPatchSetsZoomAndFreeze) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetViewZoom(wb.handle, 0, 175U), 0);
  ASSERT_EQ(SetViewFreeze(wb.handle, 0, 4U, 2U), 0);
  fm_sheet_view_info_t v = FM_STRUCT_INIT(fm_sheet_view_info_t);
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.zoom_pct, 175U);
  EXPECT_EQ(v.freeze_rows, 4U);
  EXPECT_EQ(v.freeze_cols, 2U);
}

int32_t ReadVisibility(const WorkbookGuard& wb, std::size_t sheet) {
  int32_t visibility = -1;
  EXPECT_EQ(fm_sheet_get_visibility(wb.handle, sheet, &visibility), 0);
  return visibility;
}

TEST(FormulonCApiSheetLayout, GetVisibilityReadsVisibleByDefault) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(ReadVisibility(wb, 0), FM_SHEET_VISIBILITY_VISIBLE);
  EXPECT_EQ(fm_sheet_get_visibility(wb.handle, 0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  int32_t out = 0;
  EXPECT_EQ(fm_sheet_get_visibility(wb.handle, 99U, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiSheetLayout, SetVisibilityStatesAllThreeTabStates) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // A workbook keeps one visible sheet, so sheet 0 can only be hidden beside another.
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Visible", 1U, ignored_index()), 0);

  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 0, FM_SHEET_VISIBILITY_VERY_HIDDEN), 0);
  EXPECT_EQ(ReadVisibility(wb, 0), FM_SHEET_VISIBILITY_VERY_HIDDEN);

  // Demotion to plain hidden.
  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 0, FM_SHEET_VISIBILITY_HIDDEN), 0);
  EXPECT_EQ(ReadVisibility(wb, 0), FM_SHEET_VISIBILITY_HIDDEN);

  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 0, FM_SHEET_VISIBILITY_VISIBLE), 0);
  EXPECT_EQ(ReadVisibility(wb, 0), FM_SHEET_VISIBILITY_VISIBLE);
}

TEST(FormulonCApiSheetLayout, SetVisibilityRejectsUnknownStateAndLeavesTheSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // A workbook keeps one visible sheet, so sheet 0 can only be hidden beside another.
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Visible", 1U, ignored_index()), 0);
  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 0, FM_SHEET_VISIBILITY_VERY_HIDDEN), 0);

  EXPECT_EQ(fm_sheet_set_visibility(wb.handle, 0, 3), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_visibility(wb.handle, 0, -1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_visibility(wb.handle, 99U, FM_SHEET_VISIBILITY_HIDDEN), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_visibility(nullptr, 0, FM_SHEET_VISIBILITY_HIDDEN), FM_STATUS_BINDING_NULL_POINTER);

  EXPECT_EQ(ReadVisibility(wb, 0), FM_SHEET_VISIBILITY_VERY_HIDDEN);
}

TEST(FormulonCApiSheetLayout, SetVisibilityRefusesToHideTheOnlyVisibleSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_visibility(wb.handle, 0, FM_SHEET_VISIBILITY_HIDDEN), FM_STATUS_PRECONDITION_FAILED);
  EXPECT_EQ(ReadVisibility(wb, 0), FM_SHEET_VISIBILITY_VISIBLE);
}

TEST(FormulonCApiSheetLayout, ViewPatchSetsDisplayFlags) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(
      SetViewFlag(wb.handle, 0, &fm_sheet_view_patch_t::show_grid_lines_op, &fm_sheet_view_patch_t::show_grid_lines, 0),
      0);
  ASSERT_EQ(SetViewFlag(wb.handle, 0, &fm_sheet_view_patch_t::show_row_column_headers_op,
                        &fm_sheet_view_patch_t::show_row_column_headers, 0),
            0);
  ASSERT_EQ(SetViewFlag(wb.handle, 0, &fm_sheet_view_patch_t::show_zeros_op, &fm_sheet_view_patch_t::show_zeros, 0), 0);
  ASSERT_EQ(
      SetViewFlag(wb.handle, 0, &fm_sheet_view_patch_t::right_to_left_op, &fm_sheet_view_patch_t::right_to_left, 1), 0);
  ASSERT_EQ(SetViewFlag(wb.handle, 0, &fm_sheet_view_patch_t::tab_selected_op, &fm_sheet_view_patch_t::tab_selected, 1),
            0);
  ASSERT_EQ(SetViewMode(wb.handle, 0, FM_SHEET_VIEW_MODE_PAGE_BREAK_PREVIEW), 0);
  fm_sheet_view_info_t v = FM_STRUCT_INIT(fm_sheet_view_info_t);
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.show_grid_lines, 0);
  EXPECT_EQ(v.show_row_column_headers, 0);
  EXPECT_EQ(v.show_zeros, 0);
  EXPECT_EQ(v.right_to_left, 1);
  EXPECT_EQ(v.tab_selected, 1);
  EXPECT_EQ(v.view_mode, FM_SHEET_VIEW_MODE_PAGE_BREAK_PREVIEW);
}

TEST(FormulonCApiSheetLayout, AutoFilterXmlRoundTripsAndClears) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  const char* xml =
      "<autoFilter ref=\"A1:C10\"><filterColumn colId=\"1\"><filters><filter val=\"east\"/>"
      "</filters></filterColumn></autoFilter>";
  ASSERT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, xml), 0);

  const char* actual = nullptr;
  ASSERT_EQ(fm_sheet_get_auto_filter_xml(wb.handle, 0, &actual), 0);
  ASSERT_NE(actual, nullptr);
  EXPECT_STREQ(actual, xml);

  ASSERT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, ""), 0);
  ASSERT_EQ(fm_sheet_get_auto_filter_xml(wb.handle, 0, &actual), 0);
  ASSERT_NE(actual, nullptr);
  EXPECT_STREQ(actual, "");
}

TEST(FormulonCApiSheetLayout, AutoFilterXmlSurvivesXlsxSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* xml =
      "<autoFilter ref=\"A1:C10\"><filterColumn colId=\"1\"><filters><filter val=\"east\"/>"
      "</filters></filterColumn></autoFilter>";
  ASSERT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, xml), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  const char* actual = nullptr;
  ASSERT_EQ(fm_sheet_get_auto_filter_xml(reloaded.handle, 0, &actual), 0);
  ASSERT_NE(actual, nullptr);
  EXPECT_STREQ(actual, xml);

  pugi::xml_document parsed;
  const pugi::xml_parse_result parse = parsed.load_string(actual, pugi::parse_full);
  ASSERT_TRUE(parse) << parse.description();
  ASSERT_TRUE(parsed.first_child());
  EXPECT_FALSE(parsed.first_child().next_sibling());
  EXPECT_STREQ(parsed.first_child().name(), "autoFilter");
  EXPECT_TRUE(parsed.first_child().child("filterColumn").child("filters").child("filter"));
}

TEST(FormulonCApiSheetLayout, AutoFilterXmlRejectsWrongRoot) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<filterColumn colId=\"0\"/>"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, nullptr), 0);
  // A longer element name that merely starts with "autoFilter".
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilterColumn colId=\"0\"/>"), 0);
  // An unterminated element would break the whole worksheet part.
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter ref=\"A1:C10\">"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter ref=\"A1:C10\"/><autoFilter/>"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter ref=\"A1:C10\"/><filterColumn/>"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter ref=\"A1:C10\""), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter ref=\"A1:C10/>"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<?xml version=\"1.0\"?><autoFilter/>"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<!-- comment --><autoFilter/>"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<?processing?><autoFilter/>"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<!DOCTYPE autoFilter><autoFilter/>"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter"), 0);
  EXPECT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter ref=\"A1:C10\"/>"), 0);
}

TEST(FormulonCApiSheetLayout, AutoFilterXmlFailureIncludesParseContext) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter ref=\"A1:C10\""), 0);
  EXPECT_NE(std::string(fm_get_last_error_context()).find("pugixml"), std::string::npos);
}

TEST(FormulonCApiSheetLayout, SetViewModeRejectsUnknownModeAndKeepsTheView) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetViewMode(wb.handle, 0, FM_SHEET_VIEW_MODE_PAGE_LAYOUT), 0);
  EXPECT_EQ(SetViewMode(wb.handle, 0, -1), kInvalidArgument);
  EXPECT_EQ(SetViewMode(wb.handle, 0, 3), kInvalidArgument);
  fm_sheet_view_info_t v = FM_STRUCT_INIT(fm_sheet_view_info_t);
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.view_mode, FM_SHEET_VIEW_MODE_PAGE_LAYOUT);

  ASSERT_EQ(SetViewMode(wb.handle, 0, FM_SHEET_VIEW_MODE_NORMAL), 0);
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.view_mode, FM_SHEET_VIEW_MODE_NORMAL);
}

TEST(FormulonCApiSheetLayout, ColumnWidthAndOutline) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 0U, 4U, 18.5), 0);
  ASSERT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 6U, 8U, 2U), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 10U, 10U, 1), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 12U, 12U, 0.0), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 4U);

  fm_column_layout_info_t a = FM_STRUCT_INIT(fm_column_layout_info_t);
  fm_column_layout_info_t b = FM_STRUCT_INIT(fm_column_layout_info_t);
  fm_column_layout_info_t c = FM_STRUCT_INIT(fm_column_layout_info_t);
  fm_column_layout_info_t d = FM_STRUCT_INIT(fm_column_layout_info_t);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0, &a), 0);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 1, &b), 0);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 2, &c), 0);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 3, &d), 0);
  EXPECT_EQ(a.first, 0U);
  EXPECT_EQ(a.last, 4U);
  EXPECT_DOUBLE_EQ(a.width_chars, 18.5);
  EXPECT_EQ(a.has_width_chars, 1);
  EXPECT_EQ(a.has_xf_index, 0);
  EXPECT_EQ(b.first, 6U);
  EXPECT_EQ(b.last, 8U);
  EXPECT_EQ(b.outline_level, 2U);
  EXPECT_EQ(b.has_width_chars, 0);
  EXPECT_EQ(c.first, 10U);
  EXPECT_EQ(c.last, 10U);
  EXPECT_EQ(c.hidden, 1);
  EXPECT_EQ(c.has_width_chars, 0);
  EXPECT_EQ(d.first, 12U);
  EXPECT_EQ(d.last, 12U);
  EXPECT_DOUBLE_EQ(d.width_chars, 0.0);
  EXPECT_EQ(d.has_width_chars, 1);
  EXPECT_EQ(d.has_xf_index, 0);
}

TEST(FormulonCApiSheetLayout, WidthOverlayPreservesAdjacentStylesAcrossSaveLoad) {
  // Seed two adjacent style-bearing spans through the core writer, then make
  // the actual mutation through the C ABI. This catches the old upsert path,
  // which inherited the first span's style across the entire target.
  formulon::Workbook source = formulon::Workbook::create();
  formulon::ColumnLayout first;
  first.first = 0U;
  first.last = 0U;
  first.has_style = true;
  first.style_xf = 1U;
  formulon::ColumnLayout second;
  second.first = 1U;
  second.last = 1U;
  second.has_style = true;
  second.style_xf = 2U;
  source.sheet(0).mutable_layout().columns = {first, second};
  auto source_bytes = source.save();
  ASSERT_TRUE(static_cast<bool>(source_bytes)) << source_bytes.error().message;

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(source_bytes.value().data(), source_bytes.value().size(), &wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 0U, 1U, 12.5), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  fm_column_layout_info_t col0 = FM_STRUCT_INIT(fm_column_layout_info_t);
  fm_column_layout_info_t col1 = FM_STRUCT_INIT(fm_column_layout_info_t);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0U, &col0), 0);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 1U, &col1), 0);
  EXPECT_EQ(col0.first, 0U);
  EXPECT_EQ(col0.last, 0U);
  EXPECT_DOUBLE_EQ(col0.width_chars, 12.5);
  EXPECT_EQ(col0.has_width_chars, 1);
  EXPECT_EQ(col0.has_xf_index, 1);
  EXPECT_EQ(col0.xf_index, 1U);
  EXPECT_EQ(col1.first, 1U);
  EXPECT_EQ(col1.last, 1U);
  EXPECT_DOUBLE_EQ(col1.width_chars, 12.5);
  EXPECT_EQ(col1.has_width_chars, 1);
  EXPECT_EQ(col1.has_xf_index, 1);
  EXPECT_EQ(col1.xf_index, 2U);

  // The other column setters use the same field-specific overlay and must
  // leave both width and style attributes on their intersections untouched.
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 0U, 1U, 1), 0);
  ASSERT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 0U, 1U, 2U), 0);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0U, &col0), 0);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 1U, &col1), 0);
  EXPECT_EQ(col0.hidden, 1);
  EXPECT_EQ(col1.hidden, 1);
  EXPECT_EQ(col0.outline_level, 2U);
  EXPECT_EQ(col1.outline_level, 2U);
  EXPECT_DOUBLE_EQ(col0.width_chars, 12.5);
  EXPECT_DOUBLE_EQ(col1.width_chars, 12.5);
  EXPECT_EQ(col0.xf_index, 1U);
  EXPECT_EQ(col1.xf_index, 2U);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  ASSERT_EQ(fm_sheet_get_column_layout_at(reloaded.handle, 0, 0U, &col0), 0);
  ASSERT_EQ(fm_sheet_get_column_layout_at(reloaded.handle, 0, 1U, &col1), 0);
  EXPECT_EQ(col0.has_xf_index, 1);
  EXPECT_EQ(col0.xf_index, 1U);
  EXPECT_EQ(col0.has_width_chars, 1);
  EXPECT_DOUBLE_EQ(col0.width_chars, 12.5);
  EXPECT_EQ(col1.has_xf_index, 1);
  EXPECT_EQ(col1.xf_index, 2U);
  EXPECT_EQ(col1.has_width_chars, 1);
  EXPECT_DOUBLE_EQ(col1.width_chars, 12.5);
}

TEST(FormulonCApiSheetLayout, WidthOverlayDoesNotSmearAttributesIntoTargetGaps) {
  formulon::Workbook source = formulon::Workbook::create();
  formulon::ColumnLayout left;
  left.first = 0U;
  left.last = 0U;
  left.hidden = true;
  left.has_style = true;
  left.style_xf = 1U;
  formulon::ColumnLayout right;
  right.first = 2U;
  right.last = 2U;
  right.outline_level = 3U;
  right.has_style = true;
  right.style_xf = 2U;
  source.sheet(0).mutable_layout().columns = {left, right};
  auto source_bytes = source.save();
  ASSERT_TRUE(static_cast<bool>(source_bytes)) << source_bytes.error().message;

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(source_bytes.value().data(), source_bytes.value().size(), &wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 0U, 3U, 9.0), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 4U);
  for (size_t i = 0; i < count; ++i) {
    fm_column_layout_info_t column = FM_STRUCT_INIT(fm_column_layout_info_t);
    ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, i, &column), 0);
    EXPECT_EQ(column.first, static_cast<std::uint32_t>(i));
    EXPECT_EQ(column.last, static_cast<std::uint32_t>(i));
    EXPECT_DOUBLE_EQ(column.width_chars, 9.0);
    EXPECT_EQ(column.has_width_chars, 1);
    if (i == 0U) {
      EXPECT_EQ(column.hidden, 1);
      EXPECT_EQ(column.has_xf_index, 1);
      EXPECT_EQ(column.xf_index, 1U);
    } else if (i == 2U) {
      EXPECT_EQ(column.hidden, 0);
      EXPECT_EQ(column.outline_level, 3U);
      EXPECT_EQ(column.has_xf_index, 1);
      EXPECT_EQ(column.xf_index, 2U);
    } else {
      EXPECT_EQ(column.hidden, 0);
      EXPECT_EQ(column.outline_level, 0U);
      EXPECT_EQ(column.has_xf_index, 0);
      EXPECT_EQ(column.xf_index, 0U);
    }
  }
}

TEST(FormulonCApiSheetLayout, GetterNormalizesLegacyNonZeroWidthPresence) {
  formulon::Workbook source = formulon::Workbook::create();
  formulon::ColumnLayout legacy;
  legacy.first = 4U;
  legacy.last = 4U;
  legacy.width = 12.5;
  EXPECT_FALSE(legacy.has_width);
  source.sheet(0).mutable_layout().columns.push_back(legacy);
  auto source_bytes = source.save();
  ASSERT_TRUE(static_cast<bool>(source_bytes)) << source_bytes.error().message;

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(source_bytes.value().data(), source_bytes.value().size(), &wb.handle), 0);
  fm_column_layout_info_t column = FM_STRUCT_INIT(fm_column_layout_info_t);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0U, &column), 0);
  EXPECT_EQ(column.first, 4U);
  EXPECT_EQ(column.last, 4U);
  EXPECT_DOUBLE_EQ(column.width_chars, 12.5);
  EXPECT_EQ(column.has_width_chars, 1);
}

TEST(FormulonCApiSheetLayout, RowOverridesUpsert) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 4U, 4U, 22.0), 0);
  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 9U, 9U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 4U, 4U, 3U), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  // Two distinct rows even though row 4 received two updates: the C
  // ABI's row setters upsert by row index.
  EXPECT_EQ(count, 2U);

  // Row overrides are listed in row order.
  fm_row_layout_info_t row4 = FM_STRUCT_INIT(fm_row_layout_info_t);
  fm_row_layout_info_t row9 = FM_STRUCT_INIT(fm_row_layout_info_t);
  ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, 0, &row4), 0);
  ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, 1, &row9), 0);
  EXPECT_EQ(row4.row, 4U);
  EXPECT_DOUBLE_EQ(row4.height_pt, 22.0);
  EXPECT_EQ(row4.outline_level, 3U);
  EXPECT_EQ(row4.hidden, 0);
  EXPECT_EQ(row4.has_xf_index, 0);
  EXPECT_EQ(row4.xf_index, 0U);
  EXPECT_EQ(row9.row, 9U);
  EXPECT_EQ(row9.hidden, 1);
  EXPECT_EQ(row9.has_xf_index, 0);
  EXPECT_EQ(row9.xf_index, 0U);
}

TEST(FormulonCApiSheetLayout, RowSettersTakeAnInclusiveSpan) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 2U, 3U, 22.0), 0);
  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 3U, 5U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 1U, 2U, 2U), 0);
  ASSERT_EQ(fm_sheet_set_row_collapsed(wb.handle, 0, 4U, 6U, 1), 0);

  struct Expected {
    uint32_t row;
    int32_t has_height_pt;
    int32_t hidden;
    uint32_t outline_level;
    int32_t collapsed;
  };
  const Expected expected[] = {
      {1U, 0, 0, 2U, 0}, {2U, 1, 0, 2U, 0}, {3U, 1, 1, 0U, 0}, {4U, 0, 1, 0U, 1}, {5U, 0, 1, 0U, 1}, {6U, 0, 0, 0U, 1},
  };
  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, std::size(expected));
  for (size_t i = 0; i < count; ++i) {
    fm_row_layout_info_t row = FM_STRUCT_INIT(fm_row_layout_info_t);
    ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, i, &row), 0);
    EXPECT_EQ(row.row, expected[i].row);
    EXPECT_EQ(row.has_height_pt, expected[i].has_height_pt) << row.row;
    EXPECT_EQ(row.hidden, expected[i].hidden) << row.row;
    EXPECT_EQ(row.outline_level, expected[i].outline_level) << row.row;
    EXPECT_EQ(row.collapsed, expected[i].collapsed) << row.row;
  }

  // Clearing `collapsed` over a span drops the rows it leaves with no state.
  ASSERT_EQ(fm_sheet_set_row_collapsed(wb.handle, 0, 0U, 10U, 0), 0);
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 5U);

  EXPECT_EQ(fm_sheet_set_row_height(wb.handle, 0, 4U, 3U, 10.0), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 4U, 3U, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 4U, 3U, 1U), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_collapsed(wb.handle, 0, 4U, 3U, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_auto_height(wb.handle, 0, 4U, 3U, 10.0), kInvalidArgument);
  EXPECT_EQ(fm_sheet_reset_row_height(wb.handle, 0, 4U, 3U), kInvalidArgument);
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 5U);

  // An automatic height over rows 4..5 keeps them automatic; clearing the
  // height over 1..5 then leaves only the hidden and outline state behind.
  ASSERT_EQ(fm_sheet_set_row_auto_height(wb.handle, 0, 4U, 5U, 18.0), 0);
  for (size_t i = 3; i < 5; ++i) {
    fm_row_layout_info_t row = FM_STRUCT_INIT(fm_row_layout_info_t);
    ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, i, &row), 0);
    EXPECT_EQ(row.has_height_pt, 1) << row.row;
    EXPECT_EQ(row.custom_height, 0) << row.row;
  }
  ASSERT_EQ(fm_sheet_reset_row_height(wb.handle, 0, 1U, 5U), 0);
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 5U);
  for (size_t i = 0; i < count; ++i) {
    fm_row_layout_info_t row = FM_STRUCT_INIT(fm_row_layout_info_t);
    ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, i, &row), 0);
    EXPECT_EQ(row.has_height_pt, 0) << row.row;
  }
}

TEST(FormulonCApiSheetLayout, RowHiddenDirtiesSubtotalOverTheRow) {
  // SUBTOTAL(109, ...) sums while skipping manually hidden rows; that
  // dependency on row-visibility state is invisible to the ordinary
  // dep-graph edges a cell-value write dirties, so `fm_sheet_set_row_hidden`
  // has to dirty it some other way.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 10.0), 0);                     // A1
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 0, 20.0), 0);                     // A2
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, 0, "=SUBTOTAL(109,A1:A2)"), 0);  // A3
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 2, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 30.0);

  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 0U, 0U, 1), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 2, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 20.0) << "A1's row is now hidden; SUBTOTAL(109,...) must exclude it";
}

TEST(FormulonCApiSheetLayout, ZeroHeightSettersAndClearsPreserveEffectiveVisibility) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(2U);
  wb.handle->workbook().sheet(0).mutable_format_defaults().zero_height = true;

  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 3U, 3U, 22.0), 0);
  EXPECT_EQ(RowOverride(wb, 3U).hidden, 1);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 4U, 4U, 2U), 0);
  EXPECT_EQ(RowOverride(wb, 4U).hidden, 1);
  ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 5U, 5U, 1U), 0);
  EXPECT_EQ(RowOverride(wb, 5U).hidden, 1);

  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 7U, 7U, 0), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 7U, 7U, 30.0), 0);
  ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 7U, 7U, 1U), 0);
  ASSERT_EQ(fm_sheet_reset_row_height(wb.handle, 0, 7U, 7U), 0);
  ASSERT_EQ(fm_sheet_reset_row_xf_index(wb.handle, 0, 7U, 7U), 0);

  const fm_row_layout_info_t visible = RowOverride(wb, 7U);
  EXPECT_EQ(visible.row, 7U);
  EXPECT_EQ(visible.hidden, 0);
  EXPECT_EQ(visible.has_height_pt, 0);
  EXPECT_EQ(visible.has_xf_index, 0);
}

TEST(FormulonCApiSheetLayout, ZeroHeightVisibilitySetterDirtiesOnlyEffectiveTransitions) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().sheet(0).mutable_format_defaults().zero_height = true;
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 0U, 10.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1U, 0U, 20.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2U, 0U, "=SUBTOTAL(109,A1:A2)"), 0);

  auto initial = wb.handle->workbook().recalc(formulon::eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(initial));
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 2U, 0U, &value), 0);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 0.0);

  // Row 0 is already effectively hidden through zeroHeight.
  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 0U, 0U, 1), 0);
  auto unchanged = wb.handle->workbook().recalc(formulon::eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(unchanged));
  EXPECT_EQ(unchanged.value().cells_evaluated, 0U);

  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 0U, 0U, 0), 0);
  auto changed = wb.handle->workbook().recalc(formulon::eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(changed));
  EXPECT_GT(changed.value().cells_evaluated, 0U);
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 2U, 0U, &value), 0);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 10.0);

  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 0U, 0U, 0), 0);
  unchanged = wb.handle->workbook().recalc(formulon::eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(unchanged));
  EXPECT_EQ(unchanged.value().cells_evaluated, 0U);
}

TEST(FormulonCApiSheetLayout, InvalidSheetIndexRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_view_info_t v = FM_STRUCT_INIT(fm_sheet_view_info_t);
  EXPECT_EQ(fm_sheet_get_view(wb.handle, 99U, &v), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetViewZoom(wb.handle, 99U, 100U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetViewFreeze(wb.handle, 99U, 1U, 1U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetViewVisibility(wb.handle, 99U, FM_SHEET_VISIBILITY_HIDDEN), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 99U, 0U, 1U, 10.0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

// The row setters take a bare row index rather than a span, so they reach
// `check_sheet_index` on a path of their own; a sheet-index guard added to
// the column setters alone would leave these three writing into whatever
// `sheet(99)` resolves to.
TEST(FormulonCApiSheetLayout, RowSettersRejectInvalidSheetIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_row_height(wb.handle, 99U, 0U, 0U, 20.0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_row_hidden(wb.handle, 99U, 0U, 0U, 1), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_row_outline(wb.handle, 99U, 0U, 0U, 2U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);

  // Sheet 0 exists and must not have absorbed any of the three.
  size_t count = 1U;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiSheetLayout, ColumnSetterRejectsInverseSpan) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 0, 5U, 3U, 10.0), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 5U, 3U, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 5U, 3U, 2U), kInvalidArgument);
}

// A refused call has to leave the sheet exactly as it was. Returning the
// error after the overlay had already run would report a rejection the model
// does not reflect, which is worse than either outcome on its own: the
// caller believes nothing happened and the file says otherwise.
TEST(FormulonCApiSheetLayout, RejectedColumnSpanLeavesTheLayoutUntouched) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 2U, 14.0), 0);

  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 0, 5U, 3U, 99.0), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 5U, 3U, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 5U, 3U, 7U), kInvalidArgument);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U) << "a refused span still created or split an entry";
  fm_column_layout_info_t only = FM_STRUCT_INIT(fm_column_layout_info_t);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0U, &only), 0);
  EXPECT_EQ(only.first, 1U);
  EXPECT_EQ(only.last, 2U);
  EXPECT_DOUBLE_EQ(only.width_chars, 14.0);
  EXPECT_EQ(only.hidden, 0);
  EXPECT_EQ(only.outline_level, 0U);
}

// A width the number writer cannot spell reaches `<col width="...">` as an
// empty attribute value, which is not a lexical `xsd:double`, so Excel
// repairs the worksheet instead of reading a wrong width. A negative width
// is outside the attribute's own domain. Zero stays valid - it is the width
// of a zero-width column, and `ColumnWidthAndOutline` above relies on it.
TEST(FormulonCApiSheetLayout, ColumnWidthRejectsNonFiniteOrNegativeWidth) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 2U, 14.0), 0);

  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 2U, kNaN), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 2U, kInfinity), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 2U, -kInfinity), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 2U, -1.0), kInvalidArgument);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U) << "a refused width still created or split an entry";
  fm_column_layout_info_t only = FM_STRUCT_INIT(fm_column_layout_info_t);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0U, &only), 0);
  EXPECT_EQ(only.first, 1U);
  EXPECT_EQ(only.last, 2U);
  EXPECT_DOUBLE_EQ(only.width_chars, 14.0);
}

// Same contract for `<row ht="...">`.
TEST(FormulonCApiSheetLayout, RowHeightRejectsNonFiniteOrNegativeHeight) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 3U, 3U, 22.0), 0);

  EXPECT_EQ(fm_sheet_set_row_height(wb.handle, 0, 3U, 3U, kNaN), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_height(wb.handle, 0, 3U, 3U, kInfinity), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_height(wb.handle, 0, 3U, 3U, -kInfinity), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_height(wb.handle, 0, 3U, 3U, -1.0), kInvalidArgument);
  // A refused height must not have created a second override either.
  EXPECT_EQ(fm_sheet_set_row_height(wb.handle, 0, 7U, 7U, kNaN), kInvalidArgument);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U) << "a refused height still created an override";
  fm_row_layout_info_t only = FM_STRUCT_INIT(fm_row_layout_info_t);
  ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, 0U, &only), 0);
  EXPECT_EQ(only.row, 3U);
  EXPECT_DOUBLE_EQ(only.height_pt, 22.0);

  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 3U, 3U, 0.0), 0);
}

// A span reaching past column XFD is written as a `<col min max>` pair
// outside the grid, which Excel reads as a damaged worksheet. The span is
// refused rather than clamped: clamping would silently move the override
// onto the last real column, which is not the sheet the caller described.
TEST(FormulonCApiSheetLayout, ColumnSettersRejectSpanPastTheGrid) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 2U, 14.0), 0);

  // `last` past the grid with `first` inside it, and both past it.
  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 0, kLastColumn, kLastColumn + 1U, 10.0), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, kLastColumn, kLastColumn + 1U, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_outline(wb.handle, 0, kLastColumn, kLastColumn + 1U, 2U), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 0, kLastColumn + 1U, kLastColumn + 5U, 10.0), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, kLastColumn + 1U, kLastColumn + 5U, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_outline(wb.handle, 0, kLastColumn + 1U, kLastColumn + 5U, 2U), kInvalidArgument);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U) << "a refused span still created or split an entry";
  fm_column_layout_info_t only = FM_STRUCT_INIT(fm_column_layout_info_t);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0U, &only), 0);
  EXPECT_EQ(only.first, 1U);
  EXPECT_EQ(only.last, 2U);
  EXPECT_DOUBLE_EQ(only.width_chars, 14.0);
}

// Same contract for `<row r="...">`.
TEST(FormulonCApiSheetLayout, RowSettersRejectRowPastTheGrid) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 3U, 3U, 22.0), 0);

  EXPECT_EQ(fm_sheet_set_row_height(wb.handle, 0, kLastRow + 1U, kLastRow + 1U, 10.0), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, kLastRow + 1U, kLastRow + 1U, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_outline(wb.handle, 0, kLastRow + 1U, kLastRow + 1U, 2U), kInvalidArgument);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U) << "a refused row index still created an override";
  fm_row_layout_info_t only = FM_STRUCT_INIT(fm_row_layout_info_t);
  ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, 0U, &only), 0);
  EXPECT_EQ(only.row, 3U);
  EXPECT_DOUBLE_EQ(only.height_pt, 22.0);
}

// The grid bound is inclusive of its last coordinate; a check written with
// the wrong comparison would refuse the one column and row Excel does
// address, and no other test on this surface reaches them.
TEST(FormulonCApiSheetLayout, LastGridCoordinateIsStillAccepted) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, kLastColumn, kLastColumn, 11.0), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, kLastColumn, kLastColumn, 1), 0);
  ASSERT_EQ(fm_sheet_set_column_outline(wb.handle, 0, kLastColumn, kLastColumn, 2U), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, kLastRow, kLastRow, 33.0), 0);
  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, kLastRow, kLastRow, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, kLastRow, kLastRow, 2U), 0);

  fm_column_layout_info_t column = FM_STRUCT_INIT(fm_column_layout_info_t);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0U, &column), 0);
  EXPECT_EQ(column.first, kLastColumn);
  EXPECT_EQ(column.last, kLastColumn);
  EXPECT_DOUBLE_EQ(column.width_chars, 11.0);
  EXPECT_EQ(column.hidden, 1);
  EXPECT_EQ(column.outline_level, 2U);

  fm_row_layout_info_t row = FM_STRUCT_INIT(fm_row_layout_info_t);
  ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, 0U, &row), 0);
  EXPECT_EQ(row.row, kLastRow);
  EXPECT_DOUBLE_EQ(row.height_pt, 33.0);
  EXPECT_EQ(row.hidden, 1);
  EXPECT_EQ(row.outline_level, 2U);
}

TEST(FormulonCApiSheetLayout, RowOverrideReportsHeightPresenceAndCustomFlag) {
  // An auto height that only caches `ht` arrives through the reader; the
  // setter produces an explicit override.
  formulon::Workbook source = formulon::Workbook::create();
  formulon::RowLayout auto_row;
  auto_row.row = 2U;
  auto_row.height = 18.0;
  auto_row.has_height = true;
  source.sheet(0).mutable_layout().row_overrides = {auto_row};
  auto source_bytes = source.save();
  ASSERT_TRUE(static_cast<bool>(source_bytes)) << source_bytes.error().message;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(source_bytes.value().data(), source_bytes.value().size(), &wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 5U, 5U, 30.0), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  for (size_t i = 0; i < count; ++i) {
    fm_row_layout_info_t r = FM_STRUCT_INIT(fm_row_layout_info_t);
    ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, i, &r), 0);
    EXPECT_EQ(r.has_height_pt, 1);
    EXPECT_EQ(r.custom_height, r.row == 5U ? 1 : 0) << "row " << r.row;
  }
}

namespace {

fm_row_layout_info_t ReloadedRowOverride(const WorkbookGuard& wb, std::uint32_t row) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  EXPECT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  return RowOverride(reloaded, row);
}

}  // namespace

TEST(FormulonCApiSheetLayout, AutoRowHeightRefreshesTheCacheAndStaysAutomatic) {
  formulon::Workbook source = formulon::Workbook::create();
  formulon::RowLayout auto_row;
  auto_row.row = 2U;
  auto_row.height = 18.0;
  auto_row.has_height = true;
  source.sheet(0).mutable_layout().row_overrides = {auto_row};
  auto source_bytes = source.save();
  ASSERT_TRUE(static_cast<bool>(source_bytes)) << source_bytes.error().message;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(source_bytes.value().data(), source_bytes.value().size(), &wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(2U);

  ASSERT_EQ(fm_sheet_set_row_auto_height(wb.handle, 0, 2U, 2U, 31.5), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 4U, 4U, 25.0), 0);
  // An auto height on a styled, hidden, outlined row keeps all of it.
  ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 6U, 6U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 6U, 6U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 6U, 6U, 2U), 0);
  ASSERT_EQ(fm_sheet_set_row_collapsed(wb.handle, 0, 6U, 6U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_auto_height(wb.handle, 0, 6U, 6U, 40.0), 0);

  for (const fm_row_layout_info_t& r : {RowOverride(wb, 2U), ReloadedRowOverride(wb, 2U)}) {
    EXPECT_EQ(r.has_height_pt, 1);
    EXPECT_EQ(r.custom_height, 0);
    EXPECT_DOUBLE_EQ(r.height_pt, 31.5);
  }
  for (const fm_row_layout_info_t& r : {RowOverride(wb, 4U), ReloadedRowOverride(wb, 4U)}) {
    EXPECT_EQ(r.has_height_pt, 1);
    EXPECT_EQ(r.custom_height, 1);
  }
  for (const fm_row_layout_info_t& r : {RowOverride(wb, 6U), ReloadedRowOverride(wb, 6U)}) {
    EXPECT_EQ(r.custom_height, 0);
    EXPECT_DOUBLE_EQ(r.height_pt, 40.0);
    EXPECT_EQ(r.has_xf_index, 1);
    EXPECT_EQ(r.xf_index, 1U);
    EXPECT_EQ(r.hidden, 1);
    EXPECT_EQ(r.outline_level, 2U);
    EXPECT_EQ(r.collapsed, 1);
  }

  // Dropping the height keeps the row's other attributes.
  ASSERT_EQ(fm_sheet_reset_row_height(wb.handle, 0, 6U, 6U), 0);
  const fm_row_layout_info_t cleared = RowOverride(wb, 6U);
  EXPECT_EQ(cleared.has_height_pt, 0);
  EXPECT_EQ(cleared.has_xf_index, 1);
  EXPECT_EQ(cleared.hidden, 1);

  EXPECT_EQ(fm_sheet_set_row_auto_height(wb.handle, 0, kLastRow + 1U, kLastRow + 1U, 10.0), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_auto_height(wb.handle, 7, 0U, 0U, 10.0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_row_auto_height(wb.handle, 0, 0U, 0U, -1.0), kInvalidArgument);
}

TEST(FormulonCApiSheetLayout, ClearRowHeightDropsBareOverride) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 3U, 3U, 25.0), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 4U, 4U, 25.0), 0);
  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 4U, 4U, 1), 0);

  ASSERT_EQ(fm_sheet_reset_row_height(wb.handle, 0, 3U, 3U), 0);
  ASSERT_EQ(fm_sheet_reset_row_height(wb.handle, 0, 4U, 4U), 0);
  // A row without an override is a no-op.
  ASSERT_EQ(fm_sheet_reset_row_height(wb.handle, 0, 9U, 9U), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_row_layout_info_t r = FM_STRUCT_INIT(fm_row_layout_info_t);
  ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, 0, &r), 0);
  EXPECT_EQ(r.row, 4U);
  EXPECT_EQ(r.hidden, 1);
  EXPECT_EQ(r.has_height_pt, 0);
  EXPECT_EQ(r.custom_height, 0);

  EXPECT_EQ(fm_sheet_reset_row_height(wb.handle, 0, kLastRow + 1U, kLastRow + 1U), kInvalidArgument);
  EXPECT_EQ(fm_sheet_reset_row_height(wb.handle, 7, 0U, 0U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

namespace {

std::vector<fm_column_layout_info_t> ReadColumns(const fm_workbook_t* wb) {
  size_t count = 0;
  EXPECT_EQ(fm_sheet_get_column_layout_count(wb, 0, &count), 0);
  std::vector<fm_column_layout_info_t> columns(count, FM_STRUCT_INIT(fm_column_layout_info_t));
  for (size_t i = 0; i < count; ++i) {
    EXPECT_EQ(fm_sheet_get_column_layout_at(wb, 0, i, &columns[i]), 0);
  }
  return columns;
}

double ColumnWidthPt(const fm_workbook_t* wb, uint32_t col) {
  double pt = 0.0;
  EXPECT_EQ(fm_sheet_get_column_width_pt(wb, 0, col, FM_GEOMETRY_MODE_PRINT, &pt), 0);
  return pt;
}

}  // namespace

TEST(FormulonCApiSheetLayout, ClearColumnWidthSplitsSpanAndKeepsOtherState) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 0U, 4U, 20.0), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 2U, 2U, 1), 0);
  ASSERT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 2U, 2U, 3U), 0);

  ASSERT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 1U, 3U), 0);

  const auto columns = ReadColumns(wb.handle);
  // [0,0] keeps width; [1,1] and [3,3] drop out; [2,2] survives for its
  // hidden/outline state; [4,4] keeps width.
  ASSERT_EQ(columns.size(), 3U);
  EXPECT_EQ(columns[0].first, 0U);
  EXPECT_EQ(columns[0].last, 0U);
  EXPECT_EQ(columns[0].has_width_chars, 1);
  EXPECT_DOUBLE_EQ(columns[0].width_chars, 20.0);
  EXPECT_EQ(columns[1].first, 2U);
  EXPECT_EQ(columns[1].last, 2U);
  EXPECT_EQ(columns[1].has_width_chars, 0);
  EXPECT_DOUBLE_EQ(columns[1].width_chars, 0.0);
  EXPECT_EQ(columns[1].hidden, 1);
  EXPECT_EQ(columns[1].outline_level, 3U);
  EXPECT_EQ(columns[2].first, 4U);
  EXPECT_EQ(columns[2].last, 4U);
  EXPECT_EQ(columns[2].has_width_chars, 1);
}

TEST(FormulonCApiSheetLayout, ClearColumnWidthKeepsColumnStyle) {
  formulon::Workbook source = formulon::Workbook::create();
  formulon::ColumnLayout styled;
  styled.first = 0U;
  styled.last = 1U;
  styled.width = 30.0;
  styled.has_width = true;
  styled.has_style = true;
  styled.style_xf = 1U;
  source.sheet(0).mutable_layout().columns = {styled};
  auto bytes = source.save();
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(bytes.value().data(), bytes.value().size(), &wb.handle), 0);
  ASSERT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 0U, 0U), 0);

  const auto columns = ReadColumns(wb.handle);
  ASSERT_EQ(columns.size(), 2U);
  EXPECT_EQ(columns[0].has_width_chars, 0);
  EXPECT_EQ(columns[0].has_xf_index, 1);
  EXPECT_EQ(columns[0].xf_index, 1U);
  EXPECT_EQ(columns[1].has_width_chars, 1);
  EXPECT_DOUBLE_EQ(columns[1].width_chars, 30.0);
}

TEST(FormulonCApiSheetLayout, ColumnStyleOverlayPreservesOtherAttributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(3U);

  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 0U, 4U, 20.0), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 2U, 2U, 1), 0);
  ASSERT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 2U, 2U, 3U), 0);
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 0U, 4U, 2U), 0);
  ASSERT_EQ(fm_sheet_reset_column_xf_index(wb.handle, 0, 1U, 3U), 0);

  const auto columns = ReadColumns(wb.handle);
  ASSERT_EQ(columns.size(), 5U);
  EXPECT_EQ(columns[0].first, 0U);
  EXPECT_EQ(columns[0].last, 0U);
  EXPECT_EQ(columns[0].has_xf_index, 1);
  EXPECT_EQ(columns[0].xf_index, 2U);
  EXPECT_EQ(columns[1].first, 1U);
  EXPECT_EQ(columns[1].last, 1U);
  EXPECT_EQ(columns[1].has_xf_index, 0);
  EXPECT_EQ(columns[1].has_width_chars, 1);
  EXPECT_EQ(columns[2].first, 2U);
  EXPECT_EQ(columns[2].last, 2U);
  EXPECT_EQ(columns[2].has_xf_index, 0);
  EXPECT_EQ(columns[2].hidden, 1);
  EXPECT_EQ(columns[2].outline_level, 3U);
  EXPECT_EQ(columns[3].first, 3U);
  EXPECT_EQ(columns[3].last, 3U);
  EXPECT_EQ(columns[3].has_xf_index, 0);
  EXPECT_EQ(columns[4].first, 4U);
  EXPECT_EQ(columns[4].last, 4U);
  EXPECT_EQ(columns[4].has_xf_index, 1);
  EXPECT_EQ(columns[4].xf_index, 2U);
}

TEST(FormulonCApiSheetLayout, RowStyleUpsertAndClearPreserveOtherAttributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(2U);

  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 4U, 4U, 22.0), 0);
  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 4U, 4U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 4U, 4U, 3U), 0);
  ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 4U, 5U, 1U), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);

  ASSERT_EQ(fm_sheet_reset_row_xf_index(wb.handle, 0, 4U, 5U), 0);
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);

  fm_row_layout_info_t row = FM_STRUCT_INIT(fm_row_layout_info_t);
  ASSERT_EQ(fm_sheet_get_row_layout_at(wb.handle, 0, 0U, &row), 0);
  EXPECT_EQ(row.row, 4U);
  EXPECT_DOUBLE_EQ(row.height_pt, 22.0);
  EXPECT_EQ(row.has_height_pt, 1);
  EXPECT_EQ(row.custom_height, 1);
  EXPECT_EQ(row.hidden, 1);
  EXPECT_EQ(row.outline_level, 3U);
  EXPECT_EQ(row.has_xf_index, 0);
  EXPECT_EQ(row.xf_index, 0U);
}

TEST(FormulonCApiSheetLayout, ColumnAndRowStyleRejectInvalidXfWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  EXPECT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 0U, 0U, 1U), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 0U, 0U, 1U), FM_STATUS_INDEX_OUT_OF_RANGE);

  size_t column_count = 0;
  size_t row_count = 0;
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &column_count), 0);
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &row_count), 0);
  EXPECT_EQ(column_count, 0U);
  EXPECT_EQ(row_count, 0U);
}

TEST(FormulonCApiSheetLayout, StyleOverridesSurviveSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(2U);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 0U, 2U, 18.0), 0);
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 0U, 2U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 3U, 3U, 24.0), 0);
  ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 3U, 3U, 1U), 0);

  for (const std::int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    BufferGuard saved;
    ASSERT_EQ(fm_workbook_save(wb.handle, format, &saved.data, &saved.len), 0) << "format=" << format;
    WorkbookGuard reloaded;
    ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << "format=" << format;

    const auto columns = ReadColumns(reloaded.handle);
    ASSERT_EQ(columns.size(), 1U) << "format=" << format;
    EXPECT_EQ(columns[0].first, 0U);
    EXPECT_EQ(columns[0].last, 2U);
    EXPECT_EQ(columns[0].has_width_chars, 1);
    EXPECT_EQ(columns[0].has_xf_index, 1);
    EXPECT_EQ(columns[0].xf_index, 1U);

    size_t row_count = 0;
    ASSERT_EQ(fm_sheet_get_row_layout_count(reloaded.handle, 0, &row_count), 0);
    ASSERT_EQ(row_count, 1U) << "format=" << format;
    fm_row_layout_info_t row = FM_STRUCT_INIT(fm_row_layout_info_t);
    ASSERT_EQ(fm_sheet_get_row_layout_at(reloaded.handle, 0, 0U, &row), 0);
    EXPECT_EQ(row.row, 3U);
    EXPECT_DOUBLE_EQ(row.height_pt, 24.0);
    EXPECT_EQ(row.has_height_pt, 1);
    EXPECT_EQ(row.has_xf_index, 1);
    EXPECT_EQ(row.xf_index, 1U);
  }
}

// Mac Excel 365 stores an entry into an unstyled cell with its row's style,
// else its column's; the row wins where a styled row and column cross.
TEST(FormulonCApiSheetLayout, EnteredValuesTakeTheAxisStyle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(3U);
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 1U, 1U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 2U, 2U, 2U), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 2U, 5U, 0U), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 4U, 1U, "column"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 5U, 1U, "=1+1"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2U, 3U, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_bool(wb.handle, 0, 2U, 1U, 1), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 6U, 4U, "plain"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2U, 5U, 2.0), 0);
  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, 7U, 1U), 0);

  const auto xf_at = [&](uint32_t row, uint32_t col) {
    std::uint32_t xf = 99U;
    std::int32_t present = 0;
    EXPECT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, row, col, &present, &xf), 0);
    return xf;
  };
  EXPECT_EQ(xf_at(4U, 1U), 1U);
  EXPECT_EQ(xf_at(5U, 1U), 1U);
  EXPECT_EQ(xf_at(2U, 3U), 2U);
  EXPECT_EQ(xf_at(2U, 1U), 2U);
  EXPECT_EQ(xf_at(6U, 4U), 0U);
  // An explicit XF 0 is the cell's own style.
  EXPECT_EQ(xf_at(2U, 5U), 0U);
  // Clearing the column style later leaves the entered cell's copy.
  ASSERT_EQ(fm_sheet_reset_column_xf_index(wb.handle, 0, 1U, 1U), 0);
  EXPECT_EQ(xf_at(4U, 1U), 1U);
  // A blank write is not an entry.
  const formulon::Cell* blank = wb.handle->workbook().sheet(0).cell_at(7U, 1U);
  EXPECT_TRUE(blank == nullptr || !blank->has_explicit_xf);
}

TEST(FormulonCApiSheetLayout, ClearColumnWidthDropsEmptySpansAndIgnoresUncoveredColumns) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // No overrides at all: nothing is created.
  ASSERT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 0U, 9U), 0);
  EXPECT_TRUE(ReadColumns(wb.handle).empty());

  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 3U, 4U, 14.0), 0);
  // A span covering the override plus uncovered columns on both sides.
  ASSERT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 0U, 9U), 0);
  EXPECT_TRUE(ReadColumns(wb.handle).empty());

  // Clearing outside an override leaves it intact.
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 3U, 4U, 14.0), 0);
  ASSERT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 6U, 8U), 0);
  const auto columns = ReadColumns(wb.handle);
  ASSERT_EQ(columns.size(), 1U);
  EXPECT_EQ(columns[0].first, 3U);
  EXPECT_EQ(columns[0].last, 4U);
  EXPECT_DOUBLE_EQ(columns[0].width_chars, 14.0);
}

TEST(FormulonCApiSheetLayout, ClearColumnWidthRejectsBadArgumentsAndLeavesTheLayout) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 2U, 14.0), 0);

  EXPECT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 5U, 3U), kInvalidArgument);
  EXPECT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 0U, kLastColumn + 1U), kInvalidArgument);
  EXPECT_EQ(fm_sheet_reset_column_width(wb.handle, 7, 0U, 0U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_NE(fm_sheet_reset_column_width(nullptr, 0, 0U, 0U), 0);

  const auto columns = ReadColumns(wb.handle);
  ASSERT_EQ(columns.size(), 1U);
  EXPECT_EQ(columns[0].first, 1U);
  EXPECT_EQ(columns[0].last, 2U);
  EXPECT_DOUBLE_EQ(columns[0].width_chars, 14.0);
  EXPECT_EQ(fm_sheet_reset_column_width(wb.handle, 0, kLastColumn, kLastColumn), 0);
}

TEST(FormulonCApiSheetLayout, ClearedColumnFollowsALaterSheetDefaultChange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_format_defaults_patch_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  d.default_column_width_chars_op = FM_PATCH_OP_SET;
  d.default_column_width_chars = 10.0;
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &d), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 0U, 0U, 10.0), 0);
  ASSERT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 0U, 0U), 0);
  d.default_column_width_chars = 20.0;
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &d), 0);

  EXPECT_DOUBLE_EQ(ColumnWidthPt(wb.handle, 0U), ColumnWidthPt(wb.handle, 1U));
}

TEST(FormulonCApiSheetLayout, ClearColumnWidthSurvivesXlsxAndXlsbSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 0U, 3U, 18.0), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 1U, 1U, 1), 0);
  ASSERT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 1U, 2U), 0);

  for (const int32_t format :
       {static_cast<int32_t>(FM_WORKBOOK_FORMAT_XLSX), static_cast<int32_t>(FM_WORKBOOK_FORMAT_XLSB)}) {
    BufferGuard saved;
    ASSERT_EQ(fm_workbook_save(wb.handle, format, &saved.data, &saved.len), 0) << "format=" << format;
    WorkbookGuard reloaded;
    ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << "format=" << format;
    const auto columns = ReadColumns(reloaded.handle);
    ASSERT_EQ(columns.size(), 3U) << "format=" << format;
    EXPECT_EQ(columns[0].first, 0U);
    EXPECT_EQ(columns[0].has_width_chars, 1);
    EXPECT_DOUBLE_EQ(columns[0].width_chars, 18.0);
    EXPECT_EQ(columns[1].first, 1U);
    EXPECT_EQ(columns[1].last, 1U);
    EXPECT_EQ(columns[1].has_width_chars, 0) << "format=" << format;
    EXPECT_EQ(columns[1].hidden, 1);
    EXPECT_EQ(columns[2].first, 3U);
    EXPECT_EQ(columns[2].has_width_chars, 1);
    EXPECT_DOUBLE_EQ(columns[2].width_chars, 18.0);
  }
}

TEST(FormulonCApiSheetLayout, EffectiveFormatDefaultsFollowTheStatedValues) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_format_defaults_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &d), 0);
  EXPECT_EQ(d.has_default_column_width_chars, 0);
  EXPECT_DOUBLE_EQ(d.effective_default_column_width_chars, 2340.0 / 256.0);  // Base 8: 64 px, shown as 8.43.
  EXPECT_DOUBLE_EQ(d.effective_default_row_height_pt, 102.0 / 7.0);

  // zeroHeight hides rows without an explicit override, so the effective
  // layout baseline is zero even though the nominal default is retained.
  wb.handle->workbook().sheet(0).mutable_format_defaults().zero_height = true;
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &d), 0);
  EXPECT_DOUBLE_EQ(d.effective_default_row_height_pt, 0.0);
  wb.handle->workbook().sheet(0).mutable_format_defaults().zero_height = false;

  fm_sheet_format_defaults_patch_t next = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  next.default_column_width_chars_op = FM_PATCH_OP_SET;
  next.default_column_width_chars = 12.5;
  next.default_row_height_pt_op = FM_PATCH_OP_SET;
  next.default_row_height_pt = 20.25;
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &next), 0);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &d), 0);
  EXPECT_DOUBLE_EQ(d.effective_default_column_width_chars, 12.5);
  EXPECT_DOUBLE_EQ(d.effective_default_row_height_pt, 20.25);

  next.default_column_width_chars_op = FM_PATCH_OP_CLEAR;
  next.default_row_height_pt_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &next), 0);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &d), 0);
  EXPECT_DOUBLE_EQ(d.effective_default_column_width_chars, 2340.0 / 256.0);  // Base 8: 64 px, shown as 8.43.
  EXPECT_DOUBLE_EQ(d.effective_default_row_height_pt, 102.0 / 7.0);
}

TEST(FormulonCApiSheetLayout, FormatDefaultsRoundTripThroughSave) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_format_defaults_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &d), 0);
  EXPECT_EQ(d.base_column_width_chars, 8.0);

  fm_sheet_format_defaults_patch_t next = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  next.default_column_width_chars_op = FM_PATCH_OP_SET;
  next.default_column_width_chars = 12.5;
  next.default_row_height_pt_op = FM_PATCH_OP_SET;
  next.default_row_height_pt = 20.25;
  next.base_column_width_chars_op = FM_PATCH_OP_SET;
  next.base_column_width_chars = 10.0;
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &next), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  fm_sheet_format_defaults_t got = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(reloaded.handle, 0, &got), 0);
  EXPECT_EQ(got.default_column_width_chars, 12.5);
  EXPECT_EQ(got.default_row_height_pt, 20.25);
  EXPECT_EQ(got.base_column_width_chars, 10.0);
  EXPECT_EQ(got.has_default_column_width_chars, 1);
  EXPECT_EQ(got.has_default_row_height_pt, 1);
}

TEST(FormulonCApiSheetLayout, FormatDefaultsIgnoreKeptValuesAndRejectBadNumbers) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_format_defaults_patch_t unstated = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  unstated.default_column_width_chars = kNaN;  // ignored: the op is KEEP
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &unstated), 0);
  fm_sheet_format_defaults_t got = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &got), 0);
  EXPECT_EQ(got.has_default_column_width_chars, 0);
  EXPECT_EQ(got.default_column_width_chars, 0.0);

  fm_sheet_format_defaults_patch_t bad = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  bad.default_row_height_pt_op = FM_PATCH_OP_SET;
  bad.default_row_height_pt = -1.0;
  EXPECT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &bad), kInvalidArgument);
  bad.default_row_height_pt = 15.0;
  bad.base_column_width_chars_op = FM_PATCH_OP_SET;
  bad.base_column_width_chars = kInfinity;
  EXPECT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &bad), kInvalidArgument);
  bad.base_column_width_chars = 8.0;
  for (const fm_patch_op_t op : {FM_PATCH_OP_CLEAR, static_cast<fm_patch_op_t>(3)}) {
    bad.base_column_width_chars_op = op;
    EXPECT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &bad), kInvalidArgument) << op;
  }
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &got), 0);
  EXPECT_EQ(got.has_default_row_height_pt, 0);
  EXPECT_NE(fm_sheet_update_format_defaults(wb.handle, 0, nullptr), 0);
  EXPECT_NE(fm_sheet_get_format_defaults(wb.handle, 0, nullptr), 0);
}

TEST(FormulonCApiSheetLayout, ResetXfIndexOnAnUntouchedAxisIsANoOp) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_reset_row_xf_index(wb.handle, 0, 3U, 8U), 0);
  ASSERT_EQ(fm_sheet_reset_column_xf_index(wb.handle, 0, 2U, 6U), 0);
  ASSERT_EQ(fm_sheet_reset_row_height(wb.handle, 0, 3U, 8U), 0);
  ASSERT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 2U, 6U), 0);
  size_t count = 99U;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiSheetLayout, ResetColumnXfIndexKeepsTheWidthAndResetRowXfIndexKeepsTheHeight) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(2U);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 1U, 17.0), 0);
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 1U, 1U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 2U, 2U, 31.5), 0);
  ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 2U, 2U, 1U), 0);

  ASSERT_EQ(fm_sheet_reset_column_xf_index(wb.handle, 0, 1U, 1U), 0);
  ASSERT_EQ(fm_sheet_reset_row_xf_index(wb.handle, 0, 2U, 2U), 0);

  fm_column_layout_info_t col = FM_STRUCT_INIT(fm_column_layout_info_t);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0U, &col), 0);
  EXPECT_EQ(col.has_xf_index, 0);
  EXPECT_EQ(col.has_width_chars, 1);
  EXPECT_DOUBLE_EQ(col.width_chars, 17.0);
  const fm_row_layout_info_t row = RowOverride(wb, 2U);
  EXPECT_EQ(row.has_xf_index, 0);
  EXPECT_EQ(row.has_height_pt, 1);
  EXPECT_DOUBLE_EQ(row.height_pt, 31.5);
}

TEST(FormulonCApiSheetLayout, ResetRowHeightAndColumnWidthKeepHiddenAndOutline) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 4U, 4U, 22.0), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 4U, 4U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 5U, 5U, 40.0), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 5U, 5U, 2U), 0);

  ASSERT_EQ(fm_sheet_reset_column_width(wb.handle, 0, 4U, 4U), 0);
  ASSERT_EQ(fm_sheet_reset_row_height(wb.handle, 0, 5U, 5U), 0);

  fm_column_layout_info_t col = FM_STRUCT_INIT(fm_column_layout_info_t);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0U, &col), 0);
  EXPECT_EQ(col.has_width_chars, 0);
  EXPECT_EQ(col.hidden, 1);
  const fm_row_layout_info_t row = RowOverride(wb, 5U);
  EXPECT_EQ(row.has_height_pt, 0);
  EXPECT_EQ(row.outline_level, 2U);
}

TEST(FormulonCApiSheetLayout, ColumnWidthCharsAndRowHeightPtRoundTripThroughTheLayoutRecords) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 0U, 1U, 13.5), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 7U, 7U, 18.75), 0);
  fm_column_layout_info_t col = FM_STRUCT_INIT(fm_column_layout_info_t);
  ASSERT_EQ(fm_sheet_get_column_layout_at(wb.handle, 0, 0U, &col), 0);
  EXPECT_DOUBLE_EQ(col.width_chars, 13.5);
  EXPECT_EQ(col.collapsed, 0);
  const fm_row_layout_info_t row = RowOverride(wb, 7U);
  EXPECT_DOUBLE_EQ(row.height_pt, 18.75);
  EXPECT_EQ(row.custom_height, 1);
  EXPECT_EQ(row.collapsed, 0);
}

TEST(FormulonCApiSheetLayout, FormatDefaultsUseUnitSuffixedFieldsAndReportEffectiveValues) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_format_defaults_patch_t defaults = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  defaults.default_column_width_chars_op = FM_PATCH_OP_SET;
  defaults.default_column_width_chars = 11.0;
  defaults.default_row_height_pt_op = FM_PATCH_OP_SET;
  defaults.default_row_height_pt = 19.5;
  defaults.base_column_width_chars_op = FM_PATCH_OP_SET;
  defaults.base_column_width_chars = 9.0;
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &defaults), 0);

  fm_sheet_format_defaults_t read = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.has_default_column_width_chars, 1);
  EXPECT_DOUBLE_EQ(read.default_column_width_chars, 11.0);
  EXPECT_EQ(read.has_default_row_height_pt, 1);
  EXPECT_DOUBLE_EQ(read.default_row_height_pt, 19.5);
  EXPECT_DOUBLE_EQ(read.base_column_width_chars, 9.0);
  EXPECT_DOUBLE_EQ(read.effective_default_column_width_chars, 11.0);
  EXPECT_DOUBLE_EQ(read.effective_default_row_height_pt, 19.5);
}

TEST(FormulonCApiSheetLayout, ZoomPctRoundTripsThroughTheView) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetViewZoom(wb.handle, 0, 150U), 0);
  fm_sheet_view_info_t view = FM_STRUCT_INIT(fm_sheet_view_info_t);
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &view), 0);
  EXPECT_EQ(view.zoom_pct, 150U);
  EXPECT_EQ(view.show_row_column_headers, 1);
}

namespace {

fm_sheet_view_info_t ReadView(const WorkbookGuard& wb) {
  fm_sheet_view_info_t v = FM_STRUCT_INIT(fm_sheet_view_info_t);
  EXPECT_EQ(fm_sheet_get_view(wb.handle, 0, &v), 0);
  return v;
}

}  // namespace

TEST(FormulonCApiSheetLayout, ViewPatchPaneOriginAppliesAfterThePanesChange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  patch.freeze_rows_op = FM_PATCH_OP_SET;
  patch.freeze_rows = 2U;
  patch.freeze_cols_op = FM_PATCH_OP_SET;
  patch.freeze_cols = 1U;
  patch.pane_top_left_row_op = FM_PATCH_OP_SET;
  patch.pane_top_left_row = 40U;
  patch.pane_top_left_col_op = FM_PATCH_OP_SET;
  patch.pane_top_left_col = 6U;
  ASSERT_EQ(fm_sheet_update_view(wb.handle, 0, &patch), 0) << fm_get_last_error_message();
  const fm_sheet_view_info_t v = ReadView(wb);
  EXPECT_EQ(v.freeze_rows, 2U);
  EXPECT_EQ(v.freeze_cols, 1U);
  EXPECT_EQ(v.pane_top_left_row, 40U);
  EXPECT_EQ(v.pane_top_left_col, 6U);

  // Re-stating the current freeze is a no-op and keeps the origin.
  ASSERT_EQ(SetViewFreeze(wb.handle, 0, 2U, 1U), 0);
  EXPECT_EQ(ReadView(wb).pane_top_left_row, 40U);
}

TEST(FormulonCApiSheetLayout, ViewPatchPaneOriginInsideTheNewFreezeWritesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  patch.zoom_pct_op = FM_PATCH_OP_SET;
  patch.zoom_pct = 150U;
  patch.freeze_rows_op = FM_PATCH_OP_SET;
  patch.freeze_rows = 2U;
  patch.pane_top_left_row_op = FM_PATCH_OP_SET;
  patch.pane_top_left_row = 1U;
  EXPECT_EQ(fm_sheet_update_view(wb.handle, 0, &patch), kInvalidArgument);
  const fm_sheet_view_info_t v = ReadView(wb);
  EXPECT_EQ(v.freeze_rows, 0U);
  EXPECT_EQ(v.zoom_pct, 100U);
}

TEST(FormulonCApiSheetLayout, ViewPatchSplitWithAZeroFreezeReplacesTheFreeze) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetViewFreeze(wb.handle, 0, 2U, 1U), 0);
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  patch.freeze_rows_op = FM_PATCH_OP_SET;
  patch.freeze_rows = 0U;
  patch.split_x_twips_op = FM_PATCH_OP_SET;
  patch.split_x_twips = 1200.0;
  ASSERT_EQ(fm_sheet_update_view(wb.handle, 0, &patch), 0) << fm_get_last_error_message();
  const fm_sheet_view_info_t v = ReadView(wb);
  EXPECT_EQ(v.freeze_rows, 0U);
  EXPECT_EQ(v.freeze_cols, 0U);
  EXPECT_EQ(v.split_x_twips, 1200.0);

  // Clearing the split of a frozen view leaves the freeze.
  ASSERT_EQ(SetViewFreeze(wb.handle, 0, 3U, 0U), 0);
  ASSERT_EQ(SetViewSplit(wb.handle, 0, 0.0, 0.0), 0);
  EXPECT_EQ(ReadView(wb).freeze_rows, 3U);
}

TEST(FormulonCApiSheetLayout, ViewPatchRejectsAFreezePastTheGrid) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetViewFreeze(wb.handle, 0, 2U, 0U), 0);
  // The unstated axis keeps its value, so the first unfrozen cell is (2, 16384).
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  patch.freeze_cols_op = FM_PATCH_OP_SET;
  patch.freeze_cols = formulon::Sheet::kMaxCols;
  EXPECT_EQ(fm_sheet_update_view(wb.handle, 0, &patch), kInvalidArgument);
  EXPECT_EQ(ReadView(wb).freeze_cols, 0U);
  EXPECT_EQ(fm_sheet_update_view(wb.handle, 0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
}
