// Sheet-kind guard across the C ABI. Every sheet-scoped function, called through
// the generated table in `sheet_index_calls.inc`, serves the sheet kinds its
// ledger entry states: a call that needs a worksheet refuses a chartsheet,
// dialogsheet, macrosheet and any other sheet with INVALID_ARGUMENT and a
// message leading `sheet_index is a <kind>`, before any other argument is read.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <iterator>
#include <limits>
#include <set>
#include <string>
#include <type_traits>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_sheet_index_table.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

#include "sheet_index_calls.inc"

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kIndexOutOfRange = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange);
constexpr const char* kRelationshipBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

// The sheet under test is sheet 1 of the chartsheet fixture (a worksheet then a chartsheet), retyped for the
// non-chartsheet kinds, or sheet 0 of a default workbook.
enum class Subject { kWorksheet, kChartsheetXlsx, kChartsheetXlsb, kDialogsheet, kMacrosheet, kOther };

std::vector<std::uint8_t> FixtureBytes(const char* ext) {
  std::string dir = __FILE__;
  dir = dir.substr(0, dir.find_last_of("/\\"));
  std::ifstream in(dir + "/../fixtures/excel/chartsheet_twins/tab_color_rgb_red." + ext, std::ios::binary);
  EXPECT_TRUE(in.good()) << ext;
  return std::vector<std::uint8_t>((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
}

void LoadChartsheet(const char* ext, WorkbookGuard* out) {
  const std::vector<std::uint8_t> bytes = FixtureBytes(ext);
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), nullptr, &out->handle), 0) << fm_get_last_error_message();
}

// Builds the workbook of `subject`; returns the index of the sheet under test.
std::size_t Build(Subject subject, WorkbookGuard* out) {
  if (subject == Subject::kWorksheet) {
    EXPECT_EQ(fm_workbook_create_default(&out->handle), 0);
    return 0;
  }
  LoadChartsheet(subject == Subject::kChartsheetXlsb ? "xlsb" : "xlsx", out);
  if (out->handle == nullptr) {
    return 1;
  }
  const char* retype = subject == Subject::kDialogsheet  ? "dialogsheet"
                       : subject == Subject::kMacrosheet ? "xlMacrosheet"
                       : subject == Subject::kOther      ? "unknownsheet"
                                                         : nullptr;
  if (retype != nullptr) {
    formulon::Sheet& sheet = out->handle->wb->sheet(1);
    sheet.set_opaque_ooxml_sheet(sheet.opaque_ooxml_part_path(), std::string(kRelationshipBase) + retype);
  }
  return 1;
}

const char* Name(Subject subject) {
  switch (subject) {
    case Subject::kWorksheet:
      return "worksheet";
    case Subject::kChartsheetXlsx:
      return "chartsheet xlsx";
    case Subject::kChartsheetXlsb:
      return "chartsheet xlsb";
    case Subject::kDialogsheet:
      return "dialogsheet";
    case Subject::kMacrosheet:
      return "macrosheet";
    case Subject::kOther:
      return "other";
  }
  return "";
}

bool IsGuardError(fm_status_t rc) {
  return rc == kInvalidArgument && std::strncmp(fm_get_last_error_message(), "sheet_index is a ", 17) == 0;
}

bool GuardExpected(Subject subject, Serves serves) {
  switch (subject) {
    case Subject::kWorksheet:
      return serves == Serves::kChartsheet;
    case Subject::kChartsheetXlsx:
    case Subject::kChartsheetXlsb:
      return serves == Serves::kWorksheet;
    default:
      return serves != Serves::kAll;
  }
}

constexpr Subject kSubjects[] = {Subject::kWorksheet,   Subject::kChartsheetXlsx, Subject::kChartsheetXlsb,
                                 Subject::kDialogsheet, Subject::kMacrosheet,     Subject::kOther};

}  // namespace

TEST(FormulonCApiSheetKindGuard, EveryEntryServesTheKindsItsLedgerEntryStates) {
  for (const Subject subject : kSubjects) {
    for (const SheetIndexCall& entry : kSheetIndexCalls) {
      WorkbookGuard wb;
      const std::size_t sheet = Build(subject, &wb);
      ASSERT_NE(wb.handle, nullptr) << Name(subject);
      ZeroSlots slots;
      const fm_status_t rc = entry.call(wb.handle, sheet, slots);
      const bool guard = IsGuardError(rc);
      if (entry.release != nullptr) {
        entry.release(slots);
      }
      EXPECT_EQ(guard, GuardExpected(subject, entry.serves)) << entry.name << " on " << Name(subject) << ": " << rc;
    }
  }
}

// The guard comes after the handle and the sheet index and before every other argument, so a poisoned record
// still reports the guard.
TEST(FormulonCApiSheetKindGuard, GuardPrecedesRecordArguments) {
  std::size_t checked = 0;
  for (const Subject subject : kSubjects) {
    for (const SheetIndexCall& entry : kSheetIndexCalls) {
      if (entry.poison == nullptr || !GuardExpected(subject, entry.serves)) {
        continue;
      }
      ++checked;
      WorkbookGuard wb;
      const std::size_t sheet = Build(subject, &wb);
      ASSERT_NE(wb.handle, nullptr);
      ZeroSlots slots;
      entry.poison(slots);
      const fm_status_t rc = entry.call(wb.handle, sheet, slots);
      EXPECT_TRUE(IsGuardError(rc)) << entry.name << " on " << Name(subject) << ": " << rc << " "
                                    << fm_get_last_error_message();
      if (entry.release != nullptr) {
        entry.release(slots);
      }
    }
  }
  EXPECT_GT(checked, 0U);
}

// Malformed text is an ordinary argument, so the guard still wins.
TEST(FormulonCApiSheetKindGuard, GuardPrecedesTextArguments) {
  std::size_t checked = 0;
  for (const Subject subject : {Subject::kChartsheetXlsx, Subject::kDialogsheet}) {
    for (const SheetIndexCall& entry : kSheetIndexCalls) {
      if (entry.text_mask == 0U || !GuardExpected(subject, entry.serves)) {
        continue;
      }
      ++checked;
      WorkbookGuard wb;
      const std::size_t sheet = Build(subject, &wb);
      ASSERT_NE(wb.handle, nullptr);
      ZeroSlots slots;
      for (std::size_t slot = 0; slot < ZeroSlots::kCount; ++slot) {
        if ((entry.text_mask & (1U << slot)) != 0U) {
          slots.storage[slot][0] = 0xFFU;
        }
      }
      const fm_status_t rc = entry.call(wb.handle, sheet, slots);
      EXPECT_TRUE(IsGuardError(rc)) << entry.name << " on " << Name(subject) << ": " << rc << " "
                                    << fm_get_last_error_message();
      if (entry.release != nullptr) {
        entry.release(slots);
      }
    }
  }
  EXPECT_GT(checked, 0U);
}

// The calls that serve every kind are exactly these: adding one is a visible edit here.
TEST(FormulonCApiSheetKindGuard, OnlyTheSheetIdentityCallsServeEveryKind) {
  std::set<std::string> served;
  for (const SheetIndexCall& entry : kSheetIndexCalls) {
    if (entry.serves == Serves::kAll) {
      served.insert(entry.name);
    }
  }
  const std::set<std::string> expected = {"fm_sheet_get_chartsheet_chart", "fm_sheet_get_chartsheet_protection",
                                          "fm_sheet_get_kind", "fm_sheet_get_visibility", "fm_sheet_set_visibility"};
  EXPECT_EQ(served, expected);
}

TEST(FormulonCApiSheetKindGuard, MessageNamesTheKindAndWhatTheCallNeeds) {
  struct Case {
    Subject subject;
    const char* message;
    const char* context;
  };
  const Case cases[] = {
      {Subject::kChartsheetXlsx, "sheet_index is a chartsheet, but this call needs a worksheet",
       "sheet_index=1 kind=chartsheet"},
      {Subject::kDialogsheet, "sheet_index is a dialogsheet, but this call needs a worksheet",
       "sheet_index=1 kind=dialogsheet"},
      {Subject::kMacrosheet, "sheet_index is a macrosheet, but this call needs a worksheet",
       "sheet_index=1 kind=macrosheet"},
      {Subject::kOther, "sheet_index is a sheet of another type, but this call needs a worksheet",
       "sheet_index=1 kind=sheet of another type"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(Name(c.subject));
    WorkbookGuard wb;
    const std::size_t sheet = Build(c.subject, &wb);
    ASSERT_NE(wb.handle, nullptr);
    fm_value_t value = fm_value_t{};
    EXPECT_EQ(fm_sheet_get_value(wb.handle, sheet, 0, 0, &value), kInvalidArgument);
    EXPECT_STREQ(fm_get_last_error_message(), c.message);
    EXPECT_STREQ(fm_get_last_error_context(), c.context);
  }
  {
    WorkbookGuard worksheet;
    const std::size_t index = Build(Subject::kWorksheet, &worksheet);
    EXPECT_EQ(fm_sheet_set_chartsheet_protection(worksheet.handle, index, nullptr), kInvalidArgument);
    EXPECT_STREQ(fm_get_last_error_message(), "sheet_index is a worksheet, but this call needs a chartsheet");
    EXPECT_STREQ(fm_get_last_error_context(), "sheet_index=0 kind=worksheet");
  }
  WorkbookGuard wb;
  const std::size_t sheet = Build(Subject::kDialogsheet, &wb);
  size_t count = 0;
  EXPECT_EQ(fm_sheet_get_cell_count(wb.handle, sheet, &count), kInvalidArgument);
  EXPECT_STREQ(fm_get_last_error_message(),
               "sheet_index is a dialogsheet, but this call needs a worksheet or chartsheet");
}

// A chartsheet answers a list read the way a worksheet with none does.
TEST(FormulonCApiSheetKindGuard, ChartsheetReadsAnEmptyCollection) {
  for (const Subject subject : {Subject::kChartsheetXlsx, Subject::kChartsheetXlsb}) {
    SCOPED_TRACE(Name(subject));
    WorkbookGuard wb;
    const std::size_t sheet = Build(subject, &wb);
    ASSERT_NE(wb.handle, nullptr);

    std::size_t count = 99;
    EXPECT_EQ(fm_sheet_get_cell_count(wb.handle, sheet, &count), 0) << fm_get_last_error_message();
    EXPECT_EQ(count, 0U);
    count = 99;
    EXPECT_EQ(fm_sheet_get_merge_count(wb.handle, sheet, &count), 0) << fm_get_last_error_message();
    EXPECT_EQ(count, 0U);
    count = 99;
    EXPECT_EQ(fm_sheet_get_comment_count(wb.handle, sheet, &count), 0) << fm_get_last_error_message();
    EXPECT_EQ(count, 0U);
    count = 99;
    EXPECT_EQ(fm_sheet_get_pivot_count(wb.handle, sheet, &count), 0) << fm_get_last_error_message();
    EXPECT_EQ(count, 0U);
    count = 99;
    EXPECT_EQ(fm_sheet_get_table_count(wb.handle, sheet, &count), 0) << fm_get_last_error_message();
    EXPECT_EQ(count, 0U);
    count = 99;
    EXPECT_EQ(fm_sheet_get_row_layout_count(wb.handle, sheet, &count), 0) << fm_get_last_error_message();
    EXPECT_EQ(count, 0U);

    fm_cell_rect_t merge{};
    EXPECT_EQ(fm_sheet_get_merge_at(wb.handle, sheet, 0, &merge), kIndexOutOfRange);
    int32_t present = -1;
    fm_auto_filter_t filter = FM_STRUCT_INIT(fm_auto_filter_t);
    EXPECT_EQ(fm_sheet_get_auto_filter(wb.handle, sheet, &present, &filter), 0) << fm_get_last_error_message();
    EXPECT_EQ(present, 0);

    // A cell address is worksheet content.
    fm_value_t value = fm_value_t{};
    EXPECT_TRUE(IsGuardError(fm_sheet_get_value(wb.handle, sheet, 0, 0, &value)));
    EXPECT_TRUE(IsGuardError(fm_sheet_set_number(wb.handle, sheet, 0, 0, 1.0)));
  }
}

// A list filtered by a range and a sheet-level optional value answer as on an empty worksheet; a read giving a
// value per position refuses.
TEST(FormulonCApiSheetKindGuard, ChartsheetAnswersRangeListsEmptyAndRefusesPositionalReads) {
  for (const Subject subject : {Subject::kChartsheetXlsx, Subject::kChartsheetXlsb}) {
    SCOPED_TRACE(Name(subject));
    WorkbookGuard wb;
    const std::size_t sheet = Build(subject, &wb);
    ASSERT_NE(wb.handle, nullptr);
    const fm_cell_rect_t rect{0, 0, 9, 9};

    fm_cell_range_t* range = nullptr;
    ASSERT_EQ(fm_sheet_get_cells_in_range(wb.handle, sheet, &rect, 0, 0, &range), 0) << fm_get_last_error_message();
    std::size_t count = 99;
    EXPECT_EQ(fm_cell_range_get_cell_count(range, &count), 0);
    EXPECT_EQ(count, 0U);
    fm_cell_range_destroy(range);

    fm_cell_rect_t merges[1];
    count = 99;
    EXPECT_EQ(fm_sheet_get_merges_in_range(wb.handle, sheet, &rect, merges, 1, &count), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(count, 0U);
    fm_cell_address_t phonetic[1];
    count = 99;
    EXPECT_EQ(fm_sheet_get_show_phonetic_cells(wb.handle, sheet, &rect, phonetic, 1, &count), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(count, 0U);
    fm_xf_cell_t xfs[1];
    count = 99;
    EXPECT_EQ(fm_sheet_get_cell_xf_indices(wb.handle, sheet, &rect, xfs, 1, &count), 0) << fm_get_last_error_message();
    EXPECT_EQ(count, 0U);

    fm_cell_rect_t area[1];
    count = 99;
    EXPECT_EQ(fm_sheet_get_print_area(wb.handle, sheet, area, 1, &count), 0) << fm_get_last_error_message();
    EXPECT_EQ(count, 0U);
    fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
    titles.has_rows = 1;
    titles.has_columns = 1;
    EXPECT_EQ(fm_sheet_get_print_titles(wb.handle, sheet, &titles), 0) << fm_get_last_error_message();
    EXPECT_EQ(titles.has_rows, 0);
    EXPECT_EQ(titles.has_columns, 0);

    std::uint32_t bases[1];
    count = 99;
    EXPECT_TRUE(IsGuardError(fm_sheet_get_range_xf_bases(wb.handle, sheet, &rect, bases, 1, &count)));
  }
}

// Neither a print area nor print titles can be scoped to a chartsheet; a chartsheet's is set in workbook scope.
TEST(FormulonCApiSheetKindGuard, ChartsheetCannotHoldAPrintAreaOrTitlesName) {
  WorkbookGuard wb;
  const std::size_t sheet = Build(Subject::kChartsheetXlsx, &wb);
  ASSERT_NE(wb.handle, nullptr);
  for (const char* name : {"_xlnm.Print_Area", "_xlnm.Print_Titles"}) {
    SCOPED_TRACE(name);
    fm_defined_name_t input = FM_STRUCT_INIT(fm_defined_name_t);
    input.name = name;
    input.formula = "$A$1:$B$2";
    input.local_sheet_id = static_cast<int32_t>(sheet);
    EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, &input), kInvalidArgument);
    EXPECT_NE(std::strstr(fm_get_last_error_message(), "needs a worksheet scope"), nullptr)
        << fm_get_last_error_message();
    input.local_sheet_id = 0;
    EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, &input), 0) << fm_get_last_error_message();
  }
  fm_cell_rect_t area[1];
  std::size_t count = 99;
  EXPECT_EQ(fm_sheet_get_print_area(wb.handle, sheet, area, 1, &count), 0) << fm_get_last_error_message();
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiSheetKindGuard, ChartsheetPrintAreaNameAFileHoldsCanBeRestated) {
  WorkbookGuard wb;
  const std::size_t sheet = Build(Subject::kChartsheetXlsx, &wb);
  ASSERT_NE(wb.handle, nullptr);
  // A reader keeps the name as the file stores it, outside the setter's checks.
  std::vector<formulon::DefinedName> names = wb.handle->workbook().defined_names();
  names.emplace_back("_xlnm.Print_Area", "$A$1:$B$2", static_cast<std::int32_t>(sheet), false, "");
  wb.handle->workbook().set_defined_names(std::move(names));
  fm_defined_name_t input = FM_STRUCT_INIT(fm_defined_name_t);
  input.name = "_xlnm.Print_Area";
  input.formula = "$A$1:$B$2";
  input.local_sheet_id = static_cast<int32_t>(sheet);
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, &input), 0) << fm_get_last_error_message();
}

namespace {

struct ChartGuard {
  fm_chart_t* handle = nullptr;
  ~ChartGuard() { fm_chart_destroy(handle); }
};

struct ChartDataGuard {
  fm_chart_data_t* handle = nullptr;
  ~ChartDataGuard() { fm_chart_data_destroy(handle); }
};

fm_chart_literal_info_t NumberPoint(std::uint32_t idx, double number) {
  fm_chart_literal_info_t p = FM_STRUCT_INIT(fm_chart_literal_info_t);
  p.idx = idx;
  p.is_text = 0;
  p.number = number;
  return p;
}

}  // namespace

// The chart and drawing calls work on the chartsheet's own chart.
TEST(FormulonCApiSheetKindGuard, ChartsheetServesItsOwnChartAndDrawing) {
  for (const Subject subject : {Subject::kChartsheetXlsx, Subject::kChartsheetXlsb}) {
    SCOPED_TRACE(Name(subject));
    WorkbookGuard wb;
    const std::size_t sheet = Build(subject, &wb);
    ASSERT_NE(wb.handle, nullptr);

    int32_t present = 0;
    uint32_t id = 0;
    ASSERT_EQ(fm_sheet_get_chartsheet_chart(wb.handle, sheet, &present, &id), 0);
    ASSERT_EQ(present, 1);

    ChartGuard chart;
    EXPECT_EQ(fm_sheet_get_chart(wb.handle, sheet, id, &present, &chart.handle), 0) << fm_get_last_error_message();
    EXPECT_EQ(present, 1);
    ChartDataGuard data;
    EXPECT_EQ(fm_sheet_get_chart_data(wb.handle, sheet, id, &present, &data.handle), 0) << fm_get_last_error_message();
    EXPECT_EQ(present, 1);
    ChartDataGuard preview;
    const fm_chart_preview_overrides_spec_t overrides = FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
    EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, sheet, id, &overrides, &present, &preview.handle), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(present, 1);

    size_t objects = 0;
    EXPECT_EQ(fm_sheet_get_drawing_object_count(wb.handle, sheet, &objects), 0) << fm_get_last_error_message();
    EXPECT_EQ(objects, 1U);
    fm_drawing_object_info_t object = FM_STRUCT_INIT(fm_drawing_object_info_t);
    EXPECT_EQ(fm_sheet_get_drawing_object_at(wb.handle, sheet, 0, &object), 0) << fm_get_last_error_message();
    EXPECT_EQ(fm_sheet_get_drawing_object(wb.handle, sheet, id, &present, &object), 0) << fm_get_last_error_message();
    EXPECT_EQ(present, 1);
    const uint8_t* bytes = nullptr;
    size_t len = 0;
    fm_image_info_t image = FM_STRUCT_INIT(fm_image_info_t);
    EXPECT_EQ(fm_sheet_get_picture(wb.handle, sheet, id, &present, &bytes, &len, &image), 0)
        << fm_get_last_error_message();
    fm_picture_info_t picture = FM_STRUCT_INIT(fm_picture_info_t);
    EXPECT_EQ(fm_sheet_get_picture_info(wb.handle, sheet, id, &present, &picture), 0) << fm_get_last_error_message();
    fm_shape_info_t shape = FM_STRUCT_INIT(fm_shape_info_t);
    EXPECT_EQ(fm_sheet_get_shape(wb.handle, sheet, id, &present, &shape), 0) << fm_get_last_error_message();
    fm_slicer_info_t slicer = FM_STRUCT_INIT(fm_slicer_info_t);
    EXPECT_EQ(fm_sheet_get_slicer(wb.handle, sheet, id, &present, &slicer), 0) << fm_get_last_error_message();
  }
}

TEST(FormulonCApiSheetKindGuard, ChartsheetChartTakesEdits) {
  for (const Subject subject : {Subject::kChartsheetXlsx, Subject::kChartsheetXlsb}) {
    SCOPED_TRACE(Name(subject));
    WorkbookGuard wb;
    const std::size_t sheet = Build(subject, &wb);
    ASSERT_NE(wb.handle, nullptr);
    int32_t present = 0;
    uint32_t id = 0;
    ASSERT_EQ(fm_sheet_get_chartsheet_chart(wb.handle, sheet, &present, &id), 0);
    ASSERT_EQ(present, 1);
    ChartGuard chart;
    ASSERT_EQ(fm_sheet_get_chart(wb.handle, sheet, id, &present, &chart.handle), 0);
    fm_chart_info_t info = FM_STRUCT_INIT(fm_chart_info_t);
    ASSERT_EQ(fm_chart_get_info(chart.handle, &info), 0);
    fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
    ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
    fm_chart_axis_info_t axis = FM_STRUCT_INIT(fm_chart_axis_info_t);
    ASSERT_EQ(fm_chart_get_axis_at(chart.handle, 0, &axis), 0);

    fm_chart_title_t title = fm_chart_title_t{};
    title.kind = FM_CHART_TITLE_KIND_TEXT;
    title.text = "Edited";
    EXPECT_EQ(fm_sheet_set_chart_title(wb.handle, sheet, id, &title), 0)
        << "set_chart_title " << fm_get_last_error_message();
    EXPECT_EQ(fm_sheet_set_chart_legend(wb.handle, sheet, id, nullptr), 0)
        << "set_chart_legend " << fm_get_last_error_message();
    fm_chart_legend_patch_t legend = FM_STRUCT_INIT(fm_chart_legend_patch_t);
    legend.position_op = FM_PATCH_OP_SET;
    legend.position = FM_CHART_LEGEND_POSITION_TOP;
    EXPECT_EQ(fm_sheet_update_chart_legend(wb.handle, sheet, id, &legend), 0)
        << "update_chart_legend " << fm_get_last_error_message();
    fm_chart_display_patch_t display = FM_STRUCT_INIT(fm_chart_display_patch_t);
    display.plot_visible_only_op = FM_PATCH_OP_SET;
    display.plot_visible_only = 1;
    EXPECT_EQ(fm_sheet_update_chart_display(wb.handle, sheet, id, &display), 0)
        << "update_chart_display " << fm_get_last_error_message();
    fm_chart_plot_patch_t plot_patch = FM_STRUCT_INIT(fm_chart_plot_patch_t);
    plot_patch.vary_colors_op = FM_PATCH_OP_SET;
    plot_patch.vary_colors = 0;
    EXPECT_EQ(fm_sheet_update_chart_plot(wb.handle, sheet, id, 0, &plot_patch), 0)
        << "update_chart_plot " << fm_get_last_error_message();
    fm_chart_axis_patch_t axis_patch = FM_STRUCT_INIT(fm_chart_axis_patch_t);
    axis_patch.major_gridlines_op = FM_PATCH_OP_SET;
    axis_patch.major_gridlines = 1;
    EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, sheet, id, axis.id, &axis_patch), 0)
        << "update_chart_axis " << fm_get_last_error_message();
    EXPECT_EQ(fm_sheet_set_chart_type(wb.handle, sheet, id, 0, plot.type,
                                      plot.bar_direction == FM_CHART_BAR_DIRECTION_BAR ? 1 : 0, plot.grouping),
              0)
        << "set_chart_type " << fm_get_last_error_message();
    EXPECT_EQ(fm_sheet_set_chart_filter(wb.handle, sheet, id, nullptr, 0, nullptr, 0), 0)
        << "set_chart_filter " << fm_get_last_error_message();

    const std::vector<fm_chart_literal_info_t> values = {NumberPoint(0, 1.0), NumberPoint(1, 2.0)};
    fm_chart_series_patch_t series = FM_STRUCT_INIT(fm_chart_series_patch_t);
    series.values_op = FM_PATCH_OP_SET;
    series.values = fm_chart_data_source_spec_t{FM_CHART_SOURCE_KIND_LITERAL, nullptr, values.data(),
                                                static_cast<uint32_t>(values.size())};
    uint32_t added = 0;
    ASSERT_EQ(fm_sheet_add_chart_series(wb.handle, sheet, id, 0, &series, &added), 0)
        << "add_chart_series " << fm_get_last_error_message();
    EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, sheet, id, 0, added, &series), 0)
        << "update_chart_series " << fm_get_last_error_message();
    fm_chart_point_patch_t point = FM_STRUCT_INIT(fm_chart_point_patch_t);
    EXPECT_EQ(fm_sheet_update_chart_point(wb.handle, sheet, id, 0, added, 0, &point), 0)
        << "update_chart_point " << fm_get_last_error_message();
    EXPECT_EQ(fm_sheet_remove_chart_series(wb.handle, sheet, id, 0, added), 0)
        << "remove_chart_series " << fm_get_last_error_message();
    EXPECT_EQ(fm_sheet_set_chart_source(wb.handle, sheet, id, "Sheet1!$A$1:$B$2", FM_CHART_SERIES_DIRECTION_COLUMNS,
                                        FM_TRISTATE_TRUE, FM_TRISTATE_TRUE),
              0)
        << "set_chart_source " << fm_get_last_error_message();

    // The edits did not leave the chart behind.
    EXPECT_EQ(fm_sheet_get_chartsheet_chart(wb.handle, sheet, &present, &id), 0);
    EXPECT_EQ(present, 1);
  }
}

// A chartsheet reports the properties its part states. The tab colour and the margins change; a change to a setting
// its part holds no model for is refused, not dropped.
TEST(FormulonCApiSheetKindGuard, ChartsheetPropertiesRestateOrRefuse) {
  for (const Subject subject : {Subject::kChartsheetXlsx, Subject::kChartsheetXlsb}) {
    SCOPED_TRACE(Name(subject));
    WorkbookGuard wb;
    const std::size_t sheet = Build(subject, &wb);
    ASSERT_NE(wb.handle, nullptr);

    fm_color_t color = fm_color_t{};
    ASSERT_EQ(fm_sheet_get_tab_color(wb.handle, sheet, &color), 0) << fm_get_last_error_message();
    EXPECT_EQ(color.kind, FM_COLOR_KIND_RGB);
    EXPECT_EQ(color.argb, 0xFFFF0000U);
    EXPECT_EQ(fm_sheet_set_tab_color(wb.handle, sheet, &color), 0) << fm_get_last_error_message();
    fm_color_t other = color;
    other.argb = 0xFF00FF00U;
    EXPECT_EQ(fm_sheet_set_tab_color(wb.handle, sheet, &other), 0) << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_get_tab_color(wb.handle, sheet, &color), 0);
    EXPECT_EQ(color.argb, 0xFF00FF00U);
    EXPECT_EQ(fm_sheet_set_tab_color(wb.handle, sheet, nullptr), 0) << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_get_tab_color(wb.handle, sheet, &color), 0);
    EXPECT_EQ(color.kind, FM_COLOR_KIND_NONE);

    // The raw sheetPr restates by its own text and takes a code name; a worksheet-only child is refused.
    const char* sheet_pr = nullptr;
    ASSERT_EQ(fm_sheet_get_sheet_pr_xml(wb.handle, sheet, &sheet_pr), 0) << fm_get_last_error_message();
    const std::string reported = sheet_pr != nullptr ? sheet_pr : "";
    EXPECT_EQ(fm_sheet_set_sheet_pr_xml(wb.handle, sheet, reported.empty() ? nullptr : reported.c_str()), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(fm_sheet_set_sheet_pr_xml(wb.handle, sheet, "<sheetPr codeName=\"Other\"/>"), 0)
        << fm_get_last_error_message();
    const std::string changed = "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>";
    EXPECT_EQ(fm_sheet_set_sheet_pr_xml(wb.handle, sheet, changed.c_str()), kInvalidArgument);
    EXPECT_EQ(std::string(fm_get_last_error_message()),
              "xml states the pageSetUpPr, which a chartsheet does not store");

    // The margins are stored by the part: a patch moves them, and a value the patch cannot take is refused first.
    fm_page_margins_t read = FM_STRUCT_INIT(fm_page_margins_t);
    ASSERT_EQ(fm_sheet_get_page_margins(wb.handle, sheet, &read), 0) << fm_get_last_error_message();
    fm_page_margins_patch_t patch = FM_STRUCT_INIT(fm_page_margins_patch_t);
    EXPECT_EQ(fm_sheet_update_page_margins(wb.handle, sheet, &patch), 0) << fm_get_last_error_message();
    patch.left_in_op = FM_PATCH_OP_SET;
    patch.left_in = read.left_in + 1.0;
    EXPECT_EQ(fm_sheet_update_page_margins(wb.handle, sheet, &patch), 0) << fm_get_last_error_message();
    fm_page_margins_t after = FM_STRUCT_INIT(fm_page_margins_t);
    ASSERT_EQ(fm_sheet_get_page_margins(wb.handle, sheet, &after), 0);
    EXPECT_EQ(after.left_in, read.left_in + 1.0);
    patch.left_in = -1.0;
    EXPECT_EQ(fm_sheet_update_page_margins(wb.handle, sheet, &patch), kInvalidArgument);
  }
}

TEST(FormulonCApiSheetKindGuard, DialogsheetRefusesThePropertiesAChartsheetServes) {
  WorkbookGuard wb;
  const std::size_t sheet = Build(Subject::kDialogsheet, &wb);
  ASSERT_NE(wb.handle, nullptr);
  fm_color_t color = fm_color_t{};
  EXPECT_TRUE(IsGuardError(fm_sheet_get_tab_color(wb.handle, sheet, &color)));
  EXPECT_TRUE(IsGuardError(fm_sheet_set_tab_color(wb.handle, sheet, nullptr)));
  const char* xml = nullptr;
  EXPECT_TRUE(IsGuardError(fm_sheet_get_sheet_pr_xml(wb.handle, sheet, &xml)));
  EXPECT_TRUE(IsGuardError(fm_sheet_set_sheet_pr_xml(wb.handle, sheet, nullptr)));
}
