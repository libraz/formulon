//
// Stable C ABI tests for the print-authoring surface: raw print-settings
// XML, `fitToPage`, print area / titles, manual page breaks and the typed
// patch setters with their read-back getters.
//
// The driver is C++ for gtest convenience; every mutation and observation
// stays on the pure-C surface. Two properties get repeated attention
// because they are what the surface is for: a fragment the engine stores
// must be one Excel can open, and a setting written through any of the
// three surfaces must be visible to `fm_sheet_paginate` immediately.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <initializer_list>
#include <string>
#include <utility>
#include <vector>

#include "builtin_defined_names.h"
#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "defined_name.h"
#include "formulon_c_count_of.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "workbook.h"

namespace {

// `fm_row_break_t` and the three patch structs cross the ABI by pointer,
// and the bindings marshal them field by field from a hand-written layout
// table. Pin the layout so a field insertion breaks the build here rather
// than silently mis-decoding in the Python ctypes projection.
static_assert(sizeof(fm_row_break_t) == 20U, "fm_row_break_t ABI layout changed");
static_assert(offsetof(fm_row_break_t, manual) == 16U, "fm_row_break_t.manual offset changed");
static_assert(sizeof(fm_page_setup_t) == 52U, "fm_page_setup_t ABI layout changed");
static_assert(offsetof(fm_page_setup_t, fit_to_page) == 48U, "fm_page_setup_t.fit_to_page offset changed");
static_assert(sizeof(fm_page_margins_t) == 96U, "fm_page_margins_t ABI layout changed");
static_assert(offsetof(fm_page_margins_t, left_in) == 8U, "fm_page_margins_t.left_in offset changed");
static_assert(sizeof(fm_print_options_t) == 36U, "fm_print_options_t ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kPrintInvalidArea = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPrintInvalidArea);

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

// A break covering the whole axis (`first == last == 0`), or the given span.
fm_status_t SetRowBreak(fm_workbook_t* wb, uint32_t row, int32_t manual, uint32_t first = 0, uint32_t last = 0) {
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  brk.first = first;
  brk.last = last;
  brk.manual = manual;
  return fm_sheet_set_row_break(wb, 0, row, &brk);
}

fm_status_t SetColumnBreak(fm_workbook_t* wb, uint32_t col, int32_t manual, uint32_t first = 0, uint32_t last = 0) {
  fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
  brk.first = first;
  brk.last = last;
  brk.manual = manual;
  return fm_sheet_set_column_break(wb, 0, col, &brk);
}

struct PaginationGuard {
  fm_pagination_t* handle = nullptr;
  ~PaginationGuard() { fm_pagination_destroy(handle); }
  PaginationGuard() = default;
  PaginationGuard(const PaginationGuard&) = delete;
  PaginationGuard& operator=(const PaginationGuard&) = delete;
};

std::string GetPageSetupXml(fm_workbook_t* wb) {
  const char* xml = nullptr;
  EXPECT_EQ(fm_sheet_get_page_setup_xml(wb, 0, &xml), 0);
  return xml == nullptr ? std::string() : std::string(xml);
}

std::string GetSheetPrXml(fm_workbook_t* wb) {
  const char* xml = nullptr;
  EXPECT_EQ(fm_sheet_get_sheet_pr_xml(wb, 0, &xml), 0);
  return xml == nullptr ? std::string() : std::string(xml);
}

// `fitToPage` alone, through the page-setup patch.
fm_status_t SetFitToPage(fm_workbook_t* wb, std::size_t sheet_index, int32_t enabled) {
  fm_page_setup_patch_t patch = FM_STRUCT_INIT(fm_page_setup_patch_t);
  patch.fit_to_page_op = FM_PATCH_OP_SET;
  patch.fit_to_page = enabled;
  return fm_sheet_update_page_setup(wb, sheet_index, &patch);
}

}  // namespace

/* -------------------------------------------------------------------------- */
/* Raw XML - validation                                                       */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, FreshSheetReportsEmptyFragments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (auto* getter : {fm_sheet_get_page_setup_xml, fm_sheet_get_page_margins_xml, fm_sheet_get_print_options_xml,
                       fm_sheet_get_header_footer_xml, fm_sheet_get_sheet_pr_xml}) {
    const char* xml = nullptr;
    ASSERT_EQ(getter(wb.handle, 0, &xml), 0);
    // An absent element reads back as the empty string, never NULL: a
    // caller must be able to `strlen` the result unconditionally.
    ASSERT_NE(xml, nullptr);
    EXPECT_STREQ(xml, "");
  }
}

TEST(FormulonCApiPrintSettings, SetRoundTripsThroughGetter) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* fragment = "<pageSetup paperSize=\"9\" orientation=\"portrait\" scale=\"85\"/>";
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, fragment), 0);
  EXPECT_EQ(GetPageSetupXml(wb.handle), fragment);
}

TEST(FormulonCApiPrintSettings, NullRemovesElementAndResetsStructuredView) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup orientation=\"landscape\" scale=\"50\"/>"), 0);
  fm_page_setup_t before = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &before), 0);
  ASSERT_EQ(before.orientation, FM_PAGE_ORIENTATION_LANDSCAPE);
  ASSERT_EQ(before.scale_pct, 50U);

  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, nullptr), 0);
  EXPECT_EQ(GetPageSetupXml(wb.handle), "");
  fm_page_setup_t after = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &after), 0);
  EXPECT_EQ(after.orientation, FM_PAGE_ORIENTATION_PORTRAIT);
  EXPECT_EQ(after.scale_pct, 100U);
  EXPECT_EQ(after.has_orientation, 0);
  EXPECT_EQ(after.has_scale_pct, 0);
}

TEST(FormulonCApiPrintSettings, RejectsMalformedFragments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  struct Case {
    const char* xml;
    const char* why;
  };
  const Case cases[] = {
      {"<pageSetup orientation=\"portrait\"", "truncated"},
      {"<pageSetup/><pageSetup/>", "two top-level elements"},
      {"<pageMargins left=\"1\"/>", "wrong root name"},
      {"<x:pageSetup/>", "prefixed root lifted out of its namespace context"},
      {"<!-- comment only -->", "no element"},
      {"<?xml version=\"1.0\"?><pageSetup/>", "declaration as a top-level sibling"},
      {"not xml at all", "not markup"},
  };
  for (const Case& c : cases) {
    EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, c.xml), kInvalidArgument) << c.why;
    // A rejected set must leave the stored fragment untouched.
    EXPECT_EQ(GetPageSetupXml(wb.handle), "") << c.why;
  }
}

TEST(FormulonCApiPrintSettings, RejectsOversizedFragment) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::string padding(5000, 'x');
  const std::string oversized = "<pageSetup firstPageNumber=\"" + padding + "\"/>";
  EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, oversized.c_str()), FM_STATUS_SECURITY_RESOURCE_LIMIT);
  EXPECT_EQ(std::string(fm_get_last_error_context()),
            "part=page_setup_xml requested=" + std::to_string(oversized.size()) + " limit=4096");

  // `<headerFooter>` gets a wider ceiling because six sections of
  // formatting codes legitimately live there.
  const std::string big_header = "<headerFooter><oddHeader>" + padding + "</oddHeader></headerFooter>";
  EXPECT_EQ(fm_sheet_set_header_footer_xml(wb.handle, 0, big_header.c_str()), 0);
}

TEST(FormulonCApiPrintSettings, RejectsRelationshipIdWithoutPrinterSettingsPart) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // Storing this fragment would leave a relationship reference the writer
  // cannot resolve, which is exactly what makes Excel offer to repair a
  // file. The rejection is what turns a silently broken save into a
  // diagnosable caller error.
  EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup orientation=\"portrait\" r:id=\"rId1\"/>"),
            kInvalidArgument);
  EXPECT_EQ(GetPageSetupXml(wb.handle), "");
}

TEST(FormulonCApiPrintSettings, AcceptsRelationshipIdWhenPrinterSettingsPartExists) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().sheet(0).mutable_print_settings().printer_settings_path =
      "xl/printerSettings/printerSettings1.bin";
  const char* fragment = "<pageSetup orientation=\"portrait\" r:id=\"rId1\"/>";
  EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, fragment), 0);
  EXPECT_EQ(GetPageSetupXml(wb.handle), fragment);
}

/* -------------------------------------------------------------------------- */
/* fitToPage                                                                  */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, FitToPagePreservesOtherSheetPrContent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(
      fm_sheet_set_sheet_pr_xml(wb.handle, 0, "<sheetPr codeName=\"Sheet1\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"),
      0);
  ASSERT_EQ(SetFitToPage(wb.handle, 0, 1), 0);

  const std::string xml = GetSheetPrXml(wb.handle);
  // A VBA binding and a tab colour are not print settings, and toggling
  // fit-to-page must not be how a caller loses them.
  EXPECT_NE(xml.find("codeName=\"Sheet1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("<tabColor rgb=\"FFFF0000\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("fitToPage=\"1\""), std::string::npos) << xml;

  fm_page_setup_t setup = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &setup), 0);
  EXPECT_EQ(setup.fit_to_page, 1);
  EXPECT_EQ(setup.has_fit_to_page, 1);
}

TEST(FormulonCApiPrintSettings, ClearingFitToPageKeepsTheElements) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(
      fm_sheet_set_sheet_pr_xml(wb.handle, 0, "<sheetPr><pageSetUpPr fitToPage=\"1\" autoPageBreaks=\"0\"/></sheetPr>"),
      0);
  ASSERT_EQ(SetFitToPage(wb.handle, 0, 0), 0);

  const std::string xml = GetSheetPrXml(wb.handle);
  EXPECT_EQ(xml.find("fitToPage"), std::string::npos) << xml;
  EXPECT_NE(xml.find("autoPageBreaks=\"0\""), std::string::npos) << xml;

  fm_page_setup_t setup = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &setup), 0);
  EXPECT_EQ(setup.fit_to_page, 0);
  EXPECT_EQ(setup.has_fit_to_page, 0);
}

TEST(FormulonCApiPrintSettings, FitToPageCreatesSheetPrWhenAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetFitToPage(wb.handle, 0, 1), 0);
  EXPECT_EQ(GetSheetPrXml(wb.handle), "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>");
}

/* -------------------------------------------------------------------------- */
/* Print area and titles                                                      */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, SetPrintAreaWritesQualifiedAbsoluteDefinedName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t area{0, 0, 7, 5};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &area, 1), 0);

  const auto& names = wb.handle->workbook().defined_names();
  ASSERT_EQ(names.size(), 1U);
  EXPECT_EQ(names[0].name, "_xlnm.Print_Area");
  EXPECT_EQ(names[0].formula, "Sheet1!$A$1:$F$8");
  EXPECT_EQ(names[0].local_sheet_id, 0);

  fm_cell_rect_t ranges[1] = {};
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_print_area(wb.handle, 0, ranges, 1, &count), 0);
  ASSERT_EQ(count, 1U);
  EXPECT_EQ(ranges[0].first_row, 0U);
  EXPECT_EQ(ranges[0].first_col, 0U);
  EXPECT_EQ(ranges[0].last_row, 7U);
  EXPECT_EQ(ranges[0].last_col, 5U);
}

TEST(FormulonCApiPrintSettings, SetPrintAreaQuotesSheetNamesThatNeedIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_rename_sheet(wb.handle, 0, "It's 2026"), 0);
  const fm_cell_rect_t area{0, 0, 1, 1};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &area, 1), 0);
  // A space forces quoting and the apostrophe has to be doubled, or Excel
  // reads the formula as ending at the apostrophe.
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "'It''s 2026'!$A$1:$B$2");
}

TEST(FormulonCApiPrintSettings, SetPrintAreaKeepsMultiAreaOrderAndWholeAxisShape) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t two[] = {{4, 3, 19, 4}, {0, 0, 9, 1}};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, two, 2), 0);
  // Excel prints the areas in the stated order, so the engine must not
  // sort them.
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$D$5:$E$20,Sheet1!$A$1:$B$10");

  const fm_cell_rect_t columns{0, 0, formulon::Sheet::kMaxRows - 1U, 3};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &columns, 1), 0);
  // A rectangle spanning every row is stored as whole columns, as Excel
  // writes it; the getter reports the resolver's expanded rectangle.
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$A:$D");
  fm_cell_rect_t ranges[1] = {};
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_print_area(wb.handle, 0, ranges, 1, &count), 0);
  ASSERT_EQ(count, 1U);
  EXPECT_EQ(ranges[0].last_row, formulon::Sheet::kMaxRows - 1U);
  EXPECT_EQ(ranges[0].last_col, 3U);

  const fm_cell_rect_t rows{0, 0, 49, formulon::Sheet::kMaxCols - 1U};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &rows, 1), 0);
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$1:$50");
}

TEST(FormulonCApiPrintSettings, SetPrintAreaRejectsABadRectangleAndLeavesNameUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t area{0, 0, 7, 5};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &area, 1), 0);
  const fm_cell_rect_t inverted[] = {{0, 0, 7, 5}, {5, 0, 4, 0}};
  EXPECT_EQ(fm_sheet_set_print_area(wb.handle, 0, inverted, 2), kInvalidArgument);
  const fm_cell_rect_t off_grid{0, 0, formulon::Sheet::kMaxRows, 0};
  EXPECT_EQ(fm_sheet_set_print_area(wb.handle, 0, &off_grid, 1), kInvalidArgument);
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$A$1:$F$8");
}

TEST(FormulonCApiPrintSettings, SetPrintAreaOverTheRangeCapReportsResourceLimit) {
  // The per-call range cap is shared with the conditional-formatting and
  // data-validation entry points, and one failure condition carries one
  // status across all of them.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t area{0, 0, 7, 5};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &area, 1), 0);

  const std::vector<fm_cell_rect_t> areas(formulon::c_api::parts::kMaxRangesPerCApiCall + 1U, fm_cell_rect_t{});
  EXPECT_EQ(fm_sheet_set_print_area(wb.handle, 0, areas.data(), areas.size()), FM_STATUS_SECURITY_RESOURCE_LIMIT);
  EXPECT_STREQ(fm_get_last_error_context(), "part=ranges requested=16385 limit=16384");
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$A$1:$F$8");
}

// Several print areas and both title axes are stored as a bare union, which
// every save format writes and reads back unchanged.
TEST(FormulonCApiPrintSettings, MultiAreaPrintAreaAndTitlesSurviveEverySaveFormat) {
  for (const std::int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const fm_cell_rect_t two[] = {{0, 0, 1, 1}, {0, 3, 1, 4}};
    ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, two, 2), 0);
    fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
    titles.has_rows = 1;
    titles.first_row = 0U;
    titles.last_row = 1U;
    titles.has_columns = 1;
    titles.first_col = 0U;
    titles.last_col = 0U;
    ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, &titles), 0);

    std::uint8_t* bytes = nullptr;
    std::size_t len = 0;
    ASSERT_EQ(fm_workbook_save(wb.handle, format, &bytes, &len), 0) << format << ": " << fm_get_last_error_message();
    WorkbookGuard reloaded;
    const fm_status_t loaded = fm_workbook_load(bytes, len, nullptr, &reloaded.handle);
    fm_buffer_destroy(bytes);
    ASSERT_EQ(loaded, 0) << format << ": " << fm_get_last_error_message();

    fm_cell_rect_t areas[2] = {};
    std::size_t count = 0;
    ASSERT_EQ(fm_sheet_get_print_area(reloaded.handle, 0, areas, 2, &count), 0) << format;
    ASSERT_EQ(count, 2U) << format;
    for (std::size_t i = 0; i < 2U; ++i) {
      EXPECT_EQ(areas[i].first_row, two[i].first_row) << format;
      EXPECT_EQ(areas[i].first_col, two[i].first_col) << format;
      EXPECT_EQ(areas[i].last_row, two[i].last_row) << format;
      EXPECT_EQ(areas[i].last_col, two[i].last_col) << format;
    }
    fm_print_titles_t back = FM_STRUCT_INIT(fm_print_titles_t);
    ASSERT_EQ(fm_sheet_get_print_titles(reloaded.handle, 0, &back), 0) << format;
    EXPECT_EQ(back.has_rows, 1) << format;
    EXPECT_EQ(back.last_row, 1U) << format;
    EXPECT_EQ(back.has_columns, 1) << format;
    EXPECT_EQ(back.last_col, 0U) << format;
    for (const formulon::DefinedName& name : reloaded.handle->workbook().defined_names()) {
      EXPECT_EQ(name.formula.find('('), std::string::npos) << format << ": " << name.formula;
    }
  }
}

// A typed built-in name without its `_xlnm.` prefix reaches the print area
// visible from the entry sheet, or the qualifier's, and is stored prefixed.
TEST(FormulonCApiPrintSettings, TypedPrintAreaNameReachesTheVisibleBuiltIn) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t second = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", 1, &second), 0);
  const fm_cell_rect_t two[] = {{0, 0, 1, 1}, {0, 3, 1, 4}};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, two, 2), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 9, 0, "=AREAS(print_area)"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 10, 0, "=AREAS(_xlnm.Print_Area)"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, second, 0, 0, "=AREAS(Sheet1!PRINT_AREA)"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, second, 1, 0, "=AREAS(Print_Area)"), 0);

  struct Case {
    std::size_t sheet;
    uint32_t row;
    const char* stored;
    const char* shown;
  };
  const Case kCases[] = {{0, 9, "=AREAS(_xlnm.Print_Area)", "=AREAS(Print_Area)"},
                         {0, 10, "=AREAS(_xlnm.Print_Area)", "=AREAS(Print_Area)"},
                         {second, 0, "=AREAS(Sheet1!_xlnm.Print_Area)", "=AREAS(Sheet1!Print_Area)"},
                         {second, 1, "=AREAS(Print_Area)", "=AREAS(Print_Area)"}};
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  const fm_status_t loaded = fm_workbook_load(bytes, len, nullptr, &reloaded.handle);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(loaded, 0) << fm_get_last_error_message();
  for (fm_workbook_t* book : {wb.handle, reloaded.handle}) {
    ASSERT_EQ(fm_workbook_recalc(book), 0);
    for (const Case& c : kCases) {
      SCOPED_TRACE(c.stored);
      const formulon::Cell* cell = book->workbook().sheet(c.sheet).cell_at(c.row, 0);
      ASSERT_NE(cell, nullptr);
      EXPECT_EQ(cell->formula_text, c.stored);
      int32_t present = 0;
      const char* shown = nullptr;
      ASSERT_EQ(fm_sheet_get_formula(book, c.sheet, c.row, 0, &present, &shown), 0);
      EXPECT_STREQ(shown, c.shown);
      fm_value_t value{};
      ASSERT_EQ(fm_sheet_get_value(book, c.sheet, c.row, 0, &value), 0);
      if (c.sheet == 0U || c.row == 0U) {
        ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
        EXPECT_EQ(value.u.number, 2.0);
      } else {
        ASSERT_EQ(value.kind, FM_VALUE_KIND_ERROR);
        EXPECT_EQ(value.u.error_code, 4);  // #NAME?
      }
    }
  }
}

// A sheet without its own print area sees a workbook-scope one, typed bare
// or qualified with that sheet (stored `[0]!`); the sheet that has one keeps
// its own.
TEST(FormulonCApiPrintSettings, TypedPrintAreaNameFallsBackToTheWorkbookBuiltIn) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t second = 0;
  std::size_t third = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", 1, &second), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet3", 2, &third), 0);
  const fm_cell_rect_t area{0, 0, 1, 1};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &area, 1), 0);
  // A workbook-scope print area comes only from a file; the setter refuses one.
  std::vector<formulon::DefinedName> names = wb.handle->workbook().defined_names();
  names.emplace_back("_xlnm.Print_Area", "Sheet3!$A$1", -1, false, "");
  wb.handle->workbook().set_defined_names(std::move(names));
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 9, 0, "=ROWS(Print_Area)"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, second, 0, 0, "=ROWS(Print_Area)"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, second, 1, 0, "=ROWS(Sheet3!Print_Area)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  struct Case {
    std::size_t sheet;
    uint32_t row;
    const char* stored;
    double rows;
  };
  for (const Case& c : {Case{0, 9, "=ROWS(_xlnm.Print_Area)", 2.0}, Case{second, 0, "=ROWS(_xlnm.Print_Area)", 1.0},
                        Case{second, 1, "=ROWS([0]!_xlnm.Print_Area)", 1.0}}) {
    SCOPED_TRACE(c.stored);
    const formulon::Cell* cell = wb.handle->workbook().sheet(c.sheet).cell_at(c.row, 0);
    ASSERT_NE(cell, nullptr);
    EXPECT_EQ(cell->formula_text, c.stored);
    fm_value_t value{};
    ASSERT_EQ(fm_sheet_get_value(wb.handle, c.sheet, c.row, 0, &value), 0);
    ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
    EXPECT_EQ(value.u.number, c.rows);
  }
}

// A built-in name nothing defines keeps its canonical spelling and is #NAME?.
TEST(FormulonCApiPrintSettings, UnresolvedBuiltInNameTakesItsCanonicalSpelling) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=ROWS(print_titles)+ROWS(consolidate_area)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  const formulon::Cell* cell = wb.handle->workbook().sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=ROWS(Print_Titles)+ROWS(Consolidate_Area)");
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 0, &value), 0);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(value.u.error_code, 4);  // #NAME?
}

// Excel leaves =ZEILEN(Suchkriterien) #NAME? beside a user name Criteria but
// stores it as ROWS(Criteria), which binds that user name here.
TEST(FormulonCApiPrintSettings, LocalBuiltInSpellingBindsASameNamedUserName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-de_DE", &wb.handle), 0);
  const std::string body = "=" + wb.handle->workbook().sheet(0).name() + "!$A$1:$B$3";
  fm_defined_name_t name = FM_STRUCT_INIT(fm_defined_name_t);
  name.name = "Criteria";
  name.formula = body.c_str();
  name.local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK;
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &name), 0) << fm_get_last_error_message();
  const char* canonical = nullptr;
  ASSERT_EQ(fm_canonicalize_formula("=ZEILEN(Suchkriterien)", "mac-365-de_DE", nullptr, &canonical), 0);
  ASSERT_STREQ(canonical, "=ROWS(_xlnm.Criteria)");
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 9, 0, canonical), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  const formulon::Cell* cell = wb.handle->workbook().sheet(0).cell_at(9, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=ROWS(Criteria)");
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 9, 0, &value), 0);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(value.u.number, 3.0);
}

// A bare built-in name in a file binds by the opening locale's spelling:
// ja spells Criteria in English, de as Suchkriterien.
TEST(FormulonCApiPrintSettings, LoadedBareBuiltInNameBindsWhereTheLocaleSpellsItInEnglish) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 9, 0, "=ROWS(Criteria)"), 0);
  fm_defined_name_t name = FM_STRUCT_INIT(fm_defined_name_t);
  name.name = "_xlnm.Criteria";
  name.formula = "=Sheet1!$A$1:$B$3";
  name.local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK;
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &name), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 10, 0, "=ROWS(Criteria)"), 0);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0) << fm_get_last_error_message();
  struct Case {
    const char* profile;
    bool bound;
  };
  for (const Case& c : {Case{"mac-365-ja_JP", true}, Case{"mac-365-de_DE", false}}) {
    fm_load_options_spec_t spec = FM_STRUCT_INIT(fm_load_options_spec_t);
    spec.profile_id = c.profile;
    WorkbookGuard reloaded;
    ASSERT_EQ(fm_workbook_load(bytes, len, &spec, &reloaded.handle), 0) << fm_get_last_error_message();
    ASSERT_EQ(fm_workbook_recalc_full(reloaded.handle), 0);
    const formulon::Cell* bare = reloaded.handle->workbook().sheet(0).cell_at(9, 0);
    ASSERT_NE(bare, nullptr);
    EXPECT_EQ(bare->formula_text, c.bound ? "=ROWS(_xlnm.Criteria)" : "=ROWS(Criteria)") << c.profile;
    fm_value_t value{};
    ASSERT_EQ(fm_sheet_get_value(reloaded.handle, 0, 9, 0, &value), 0);
    if (c.bound) {
      ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER) << c.profile;
      EXPECT_EQ(value.u.number, 3.0);
    } else {
      ASSERT_EQ(value.kind, FM_VALUE_KIND_ERROR) << c.profile;
      EXPECT_EQ(value.u.error_code, 4);  // #NAME?
    }
    // The stored `_xlnm.` form binds under every locale.
    ASSERT_EQ(fm_sheet_get_value(reloaded.handle, 0, 10, 0, &value), 0);
    ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER) << c.profile;
    EXPECT_EQ(value.u.number, 3.0);
  }
  fm_buffer_destroy(bytes);
}

// Excel prefixes resolved built-in names in a book loaded with a bare
// _FilterDatabase that no local name resolves; the formula reads unprefixed here.
TEST(FormulonCApiPrintSettings, BareFilterDatabaseLeavesBuiltInNamesUnprefixed) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t second = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", 1, &second), 0);
  const fm_cell_rect_t area{0, 0, 1, 1};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &area, 1), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 9, 0, "=AREAS(Print_Area)"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, second, 0, 0, "=ROWS(_FilterDatabase)"), 0);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  const fm_status_t loaded = fm_workbook_load(bytes, len, nullptr, &reloaded.handle);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(loaded, 0) << fm_get_last_error_message();
  int32_t present = 0;
  const char* text = nullptr;
  ASSERT_EQ(fm_sheet_get_formula(reloaded.handle, second, 0, 0, &present, &text), 0);
  EXPECT_STREQ(text, "=ROWS(_FilterDatabase)");
  ASSERT_EQ(fm_sheet_get_formula(reloaded.handle, 0, 9, 0, &present, &text), 0);
  EXPECT_STREQ(text, "=AREAS(Print_Area)");
}

namespace {

fm_status_t DefineName(fm_workbook_t* wb, const char* name, const char* formula, int32_t scope,
                       const char* comment = nullptr, int32_t hidden = 0) {
  fm_defined_name_t input = FM_STRUCT_INIT(fm_defined_name_t);
  input.name = name;
  input.formula = formula;
  input.local_sheet_id = scope;
  input.comment = comment;
  input.hidden = hidden;
  return fm_workbook_set_defined_name(wb, &input);
}

// The stored name and scope of every defined name, in list order.
std::vector<std::pair<std::string, int32_t>> StoredNames(const fm_workbook_t* wb) {
  std::vector<std::pair<std::string, int32_t>> out;
  for (const formulon::DefinedName& name : wb->workbook().defined_names()) {
    out.emplace_back(name.name, name.local_sheet_id);
  }
  return out;
}

using NameList = std::vector<std::pair<std::string, int32_t>>;

}  // namespace

// Excel's Names.Add stores the locale's spelling of Print_Area asked for in
// workbook scope as the print area of the active worksheet.
TEST(FormulonCApiPrintSettings, LocaleSpellingOfPrintAreaDefinesTheSheetPrintArea) {
  for (const char* typed : {"Print_Area", "print_area", "PRINT_AREA"}) {
    SCOPED_TRACE(typed);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
    std::size_t second = 0;
    ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", 1, &second), 0);
    wb.handle->workbook().set_active_sheet(1U);
    ASSERT_EQ(DefineName(wb.handle, typed, "=Sheet2!$A$1:$B$2", FM_DEFINED_NAME_SCOPE_WORKBOOK, "note", 1), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(StoredNames(wb.handle), (NameList{{"_xlnm.Print_Area", 1}}));
    const formulon::DefinedName& stored = wb.handle->workbook().defined_names().front();
    EXPECT_EQ(stored.comment, "note");
    EXPECT_TRUE(stored.hidden);
    fm_cell_rect_t out[2]{};
    std::size_t count = 0;
    ASSERT_EQ(fm_sheet_get_print_area(wb.handle, second, out, 2, &count), 0);
    ASSERT_EQ(count, 1U);
    EXPECT_EQ(out[0].last_row, 1U);
    EXPECT_EQ(out[0].last_col, 1U);
    ASSERT_EQ(fm_sheet_get_print_area(wb.handle, 0, out, 2, &count), 0);
    EXPECT_EQ(count, 0U);

    // The typed spelling finds it in its sheet's scope only.
    int32_t present = 0;
    fm_defined_name_t read = FM_STRUCT_INIT(fm_defined_name_t);
    ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, typed, 1, &present, &read), 0);
    ASSERT_EQ(present, 1);
    EXPECT_STREQ(read.name, "_xlnm.Print_Area");
    EXPECT_STREQ(read.formula, "=Sheet2!$A$1:$B$2");
    ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, typed, FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &read), 0);
    EXPECT_EQ(present, 0);

    // The print API and the setter edit the same entry.
    const fm_cell_rect_t area{0, 0, 4, 3};
    ASSERT_EQ(fm_sheet_set_print_area(wb.handle, second, &area, 1), 0);
    ASSERT_EQ(DefineName(wb.handle, typed, "=Sheet2!$A$1:$C$3", 1), 0) << fm_get_last_error_message();
    EXPECT_EQ(StoredNames(wb.handle), (NameList{{"_xlnm.Print_Area", 1}}));
    ASSERT_EQ(fm_sheet_get_print_area(wb.handle, second, out, 2, &count), 0);
    ASSERT_EQ(count, 1U);
    EXPECT_EQ(out[0].last_row, 2U);
    ASSERT_EQ(fm_workbook_remove_defined_name(wb.handle, typed, 1), 0);
    EXPECT_TRUE(wb.handle->workbook().defined_names().empty());
  }
}

// A name is a built-in only in the locale's own spelling: de promotes
// Druckbereich and Drucktitel, and keeps the English spellings as user names.
TEST(FormulonCApiPrintSettings, OnlyTheLocaleSpellingDefinesABuiltInName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-de_DE", &wb.handle), 0);
  const std::string sheet = wb.handle->workbook().sheet(0).name();
  const std::string body = "=" + sheet + "!$A$1:$B$2";
  const std::string rows = "=" + sheet + "!$1:$2";
  ASSERT_EQ(DefineName(wb.handle, "Print_Area", body.c_str(), FM_DEFINED_NAME_SCOPE_WORKBOOK), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(DefineName(wb.handle, "Criteria", body.c_str(), FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  ASSERT_EQ(DefineName(wb.handle, "druckbereich", body.c_str(), FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  ASSERT_EQ(DefineName(wb.handle, "Drucktitel", rows.c_str(), FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  ASSERT_EQ(DefineName(wb.handle, "Datenbank", body.c_str(), FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  EXPECT_EQ(StoredNames(wb.handle), (NameList{{"Print_Area", -1},
                                              {"Criteria", -1},
                                              {"_xlnm.Print_Area", 0},
                                              {"_xlnm.Print_Titles", 0},
                                              {"_xlnm.Database", -1}}));
  fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
  ASSERT_EQ(fm_sheet_get_print_titles(wb.handle, 0, &titles), 0);
  EXPECT_EQ(titles.has_rows, 1);
  EXPECT_EQ(titles.last_row, 1U);
}

// en-US stores a built-in name in the letter case typed; ja stores the canonical one.
TEST(FormulonCApiPrintSettings, EnglishLocaleKeepsTheTypedCaseOfABuiltInName) {
  struct Case {
    const char* profile;
    const char* stored;
  };
  for (const Case& c : {Case{"mac-365-en_US", "_xlnm.PRINT_AREA"}, Case{"mac-365-ja_JP", "_xlnm.Print_Area"}}) {
    SCOPED_TRACE(c.profile);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_for_profile(c.profile, &wb.handle), 0);
    const std::string body = "=" + wb.handle->workbook().sheet(0).name() + "!$A$1:$B$2";
    ASSERT_EQ(DefineName(wb.handle, "PRINT_AREA", body.c_str(), FM_DEFINED_NAME_SCOPE_WORKBOOK), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(StoredNames(wb.handle), (NameList{{c.stored, 0}}));
    fm_cell_rect_t out[1]{};
    std::size_t count = 0;
    ASSERT_EQ(fm_sheet_get_print_area(wb.handle, 0, out, 1, &count), 0);
    EXPECT_EQ(count, 1U);
  }
}

// Other built-in names keep the scope asked for; the print titles go to the active sheet.
TEST(FormulonCApiPrintSettings, BuiltInNamesOtherThanThePrintAreaKeepTheirScope) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  ASSERT_EQ(DefineName(wb.handle, "Criteria", "=Sheet1!$A$1:$B$2", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(DefineName(wb.handle, "criteria", "=Sheet1!$A$1:$B$2", 0), 0);
  ASSERT_EQ(DefineName(wb.handle, "Auto_Open", "=Sheet1!$A$1", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  ASSERT_EQ(DefineName(wb.handle, "Print_Titles", "=$1:$1", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  EXPECT_EQ(
      StoredNames(wb.handle),
      (NameList{{"_xlnm.Criteria", -1}, {"_xlnm.Criteria", 0}, {"_xlnm.Auto_Open", -1}, {"_xlnm.Print_Titles", 0}}));
}

constexpr const char* kRelChartsheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
constexpr const char* kRelDialogsheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
constexpr const char* kRelMacrosheet = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";

// A new print area or print titles added at workbook scope goes to the active
// worksheet whatever its formula names, in either built-in spelling; added at a
// sheet it stays there.
TEST(FormulonCApiPrintSettings, WorkbookScopePrintAreaGoesToTheActiveWorksheet) {
  struct Case {
    const char* name;
    const char* formula;
  };
  for (const Case& c :
       {Case{"Druckbereich", "=$A$1:$A$3"}, Case{"Druckbereich", "=Sheet2!$A$1:$A$3"}, Case{"Druckbereich", "=5"},
        Case{"Druckbereich", "=Sheet1!A1+1"}, Case{"_xlnm.Print_Area", "=Sheet2!$A$1:$A$3"},
        Case{"_xlnm.Print_Area", "=5"}, Case{"Drucktitel", "=Sheet2!$1:$1"}, Case{"_xlnm.Print_Titles", "=$1:$1"}}) {
    SCOPED_TRACE(std::string(c.name) + " " + c.formula);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_for_profile("mac-365-de_DE", &wb.handle), 0);
    std::size_t second = 0;
    ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", 1, &second), 0);
    ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet3", 2, &second), 0);
    wb.handle->workbook().set_active_sheet(2U);
    const std::string stored =
        std::string("_xlnm.") + (std::strstr(c.name, "itel") != nullptr || std::strstr(c.name, "Titles") != nullptr
                                     ? "Print_Titles"
                                     : "Print_Area");
    ASSERT_EQ(DefineName(wb.handle, c.name, c.formula, FM_DEFINED_NAME_SCOPE_WORKBOOK), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(StoredNames(wb.handle), (NameList{{stored, 2}}));
    ASSERT_EQ(DefineName(wb.handle, c.name, c.formula, 1), 0) << fm_get_last_error_message();
    EXPECT_EQ(StoredNames(wb.handle), (NameList{{stored, 2}, {stored, 1}}));
  }
}

// The English spelling under another locale is an ordinary name in the scope asked for, whatever its formula.
TEST(FormulonCApiPrintSettings, EnglishPrintAreaSpellingUnderGermanIsAnOrdinaryName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-de_DE", &wb.handle), 0);
  ASSERT_EQ(DefineName(wb.handle, "Print_Area", "=5", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(DefineName(wb.handle, "Print_Area", "=Sheet1!$A$1:$A$3", 0), 0) << fm_get_last_error_message();
  EXPECT_EQ(StoredNames(wb.handle), (NameList{{"Print_Area", -1}, {"Print_Area", 0}}));
}

// With a chartsheet active, a new print area added at workbook scope stays there.
TEST(FormulonCApiPrintSettings, WorkbookScopePrintAreaStaysThereWithAChartsheetActive) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-de_DE", &wb.handle), 0);
  formulon::Workbook& book = wb.handle->workbook();
  book.add_sheet("Diagramm1");
  book.sheet(1).set_opaque_ooxml_sheet("xl/chartsheets/sheet1.xml", kRelChartsheet);
  book.set_active_sheet(1U);
  for (const char* formula : {"=Sheet1!$A$1:$A$3", "=5"}) {
    SCOPED_TRACE(formula);
    ASSERT_EQ(DefineName(wb.handle, "Druckbereich", formula, FM_DEFINED_NAME_SCOPE_WORKBOOK), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(StoredNames(wb.handle), (NameList{{"_xlnm.Print_Area", -1}}));
  }
  // The workbook-scope entry is read, updated in place and removed at workbook scope.
  int32_t present = 0;
  fm_defined_name_t read = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "Druckbereich", FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &read),
            0);
  EXPECT_EQ(present, 1);
  book.set_active_sheet(0U);
  ASSERT_EQ(DefineName(wb.handle, "_xlnm.Print_Area", "=Sheet1!$B$1", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  EXPECT_EQ(StoredNames(wb.handle), (NameList{{"_xlnm.Print_Area", -1}}));
  ASSERT_EQ(fm_workbook_remove_defined_name(wb.handle, "Druckbereich", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  EXPECT_TRUE(book.defined_names().empty());
}

// With a dialog or macro sheet active the add is refused, and nothing is stored.
TEST(FormulonCApiPrintSettings, WorkbookScopePrintAreaIsRefusedWithAnotherSheetKindActive) {
  for (const char* type : {kRelDialogsheet, kRelMacrosheet}) {
    SCOPED_TRACE(type);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_for_profile("mac-365-de_DE", &wb.handle), 0);
    formulon::Workbook& book = wb.handle->workbook();
    book.add_sheet("Other");
    book.sheet(1).set_opaque_ooxml_sheet("xl/other/sheet1.xml", type);
    book.set_active_sheet(1U);
    EXPECT_EQ(DefineName(wb.handle, "Druckbereich", "=Sheet1!$A$1:$A$3", FM_DEFINED_NAME_SCOPE_WORKBOOK),
              kInvalidArgument);
    EXPECT_EQ(DefineName(wb.handle, "_xlnm.Print_Area", "=5", FM_DEFINED_NAME_SCOPE_WORKBOOK), kInvalidArgument);
    EXPECT_TRUE(book.defined_names().empty());
    ASSERT_EQ(DefineName(wb.handle, "Druckbereich", "=Sheet1!$A$1:$A$3", 0), 0) << fm_get_last_error_message();
  }
}

// An ordinary name spelled like a built-in (a book loaded under another
// profile) is read, set and removed by its own spelling; the built-in by its
// `_xlnm.` spelling.
TEST(FormulonCApiPrintSettings, OrdinaryNameSpelledLikeABuiltInIsAddressedByItsSpelling) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-de_DE", &wb.handle), 0);
  formulon::Workbook& book = wb.handle->workbook();
  const std::string sheet = book.sheet(0).name();
  book.set_defined_names({formulon::DefinedName("Druckbereich", sheet + "!$A$1:$A$3", -1, false, ""),
                          formulon::DefinedName("_xlnm.Print_Area", sheet + "!$A$1:$A$5", -1, false, "")});
  int32_t present = 0;
  fm_defined_name_t read = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "DRUCKBEREICH", FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &read),
            0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(read.name, "Druckbereich");
  const std::string body = "=" + sheet + "!$B$1:$B$2";
  ASSERT_EQ(DefineName(wb.handle, "Druckbereich", body.c_str(), FM_DEFINED_NAME_SCOPE_WORKBOOK), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(StoredNames(wb.handle), (NameList{{"Druckbereich", -1}, {"_xlnm.Print_Area", -1}}));
  EXPECT_EQ(book.defined_names()[0].formula, sheet + "!$B$1:$B$2");
  ASSERT_EQ(fm_workbook_remove_defined_name(wb.handle, "Druckbereich", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  EXPECT_EQ(StoredNames(wb.handle), (NameList{{"_xlnm.Print_Area", -1}}));
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "Druckbereich", FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &read),
            0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(read.name, "_xlnm.Print_Area");
}

namespace {

// The English names of the built-ins a locale can spell for itself: every row of the spelling table.
std::vector<std::string> BuiltInNames() {
  std::vector<std::string> out;
  for (const formulon::BuiltinNameSpelling& spelling : formulon::kBuiltinNameSpellings) {
    out.emplace_back(spelling.english);
  }
  return out;
}

bool IsSheetBuiltIn(const std::string& english) {
  return english == "Print_Area" || english == "Print_Titles";
}

// Every profile id of the Mac host, one per locale.
std::vector<std::string> MacProfiles() {
  std::vector<std::string> out;
  std::size_t count = 0;
  EXPECT_EQ(fm_get_profile_id_count(&count), 0);
  for (std::size_t i = 0; i < count; ++i) {
    const char* id = nullptr;
    EXPECT_EQ(fm_get_profile_id_at(i, &id), 0);
    if (std::string(id).rfind("mac-365-", 0) == 0) {
      out.emplace_back(id);
    }
  }
  return out;
}

std::string Localized(const char* name, const std::string& profile) {
  const char* out = nullptr;
  EXPECT_EQ(fm_localize_defined_name(name, profile.c_str(), &out), 0);
  return out == nullptr ? std::string() : std::string(out);
}

std::string Canonical(const char* name, const std::string& profile) {
  const char* out = nullptr;
  EXPECT_EQ(fm_canonicalize_defined_name(name, profile.c_str(), &out), 0);
  return out == nullptr ? std::string() : std::string(out);
}

fm_status_t ResolveKey(const fm_workbook_t* wb, const char* name, const char* formula, int32_t scope,
                       std::string* key_name = nullptr, int32_t* key_scope = nullptr) {
  fm_defined_name_t input = FM_STRUCT_INIT(fm_defined_name_t);
  input.name = name;
  input.formula = formula;
  input.local_sheet_id = scope;
  const char* out_name = nullptr;
  int32_t out_scope = 99;
  const fm_status_t rc = fm_workbook_resolve_defined_name_key(wb, &input, &out_name, &out_scope);
  if (key_name != nullptr && out_name != nullptr) {
    *key_name = out_name;
  }
  if (key_scope != nullptr) {
    *key_scope = out_scope;
  }
  return rc;
}

}  // namespace

// The resolver names exactly the entry `fm_workbook_set_defined_name` writes, for every promoted
// built-in spelled the way every locale spells it.
TEST(FormulonCApiPrintSettings, ResolvedKeyIsWhereSetStoresEveryBuiltInInEveryLocale) {
  const std::vector<std::string> profiles = MacProfiles();
  ASSERT_EQ(profiles.size(), 14U);
  ASSERT_GE(BuiltInNames().size(), 10U);
  for (const std::string& profile : profiles) {
    for (const std::string& english : BuiltInNames()) {
      const std::string typed = Localized((std::string("_xlnm.") + english).c_str(), profile);
      SCOPED_TRACE(profile + " " + typed);
      WorkbookGuard wb;
      ASSERT_EQ(fm_workbook_create_for_profile(profile.c_str(), &wb.handle), 0);
      std::size_t second = 0;
      ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", 1, &second), 0);
      // A sheet built-in goes to the active sheet, not the one its formula names.
      const int32_t expected_scope = IsSheetBuiltIn(english) ? 0 : FM_DEFINED_NAME_SCOPE_WORKBOOK;
      std::string key_name;
      int32_t key_scope = 99;
      ASSERT_EQ(ResolveKey(wb.handle, typed.c_str(), "=Sheet2!$A$1:$B$2", FM_DEFINED_NAME_SCOPE_WORKBOOK, &key_name,
                           &key_scope),
                0)
          << fm_get_last_error_message();
      EXPECT_TRUE(wb.handle->workbook().defined_names().empty());
      EXPECT_EQ(key_scope, expected_scope);
      EXPECT_EQ(key_name, Canonical(typed.c_str(), profile));
      ASSERT_EQ(DefineName(wb.handle, typed.c_str(), "=Sheet2!$A$1:$B$2", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0)
          << fm_get_last_error_message();
      EXPECT_EQ(StoredNames(wb.handle), (NameList{{key_name, key_scope}}));

      // The key reads the entry back, and the localized display spelling reads the same entry.
      int32_t present = 0;
      fm_defined_name_t read = FM_STRUCT_INIT(fm_defined_name_t);
      ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, key_name.c_str(), key_scope, &present, &read), 0);
      EXPECT_EQ(present, 1);
      EXPECT_EQ(Localized(key_name.c_str(), profile), typed);
    }
  }
}

// The resolver and the setter agree on the value, the scope and the refusal.
TEST(FormulonCApiPrintSettings, ResolverRefusesWhatSetRefusesWithTheSameStatus) {
  for (const std::string& profile : MacProfiles()) {
    SCOPED_TRACE(profile);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_for_profile(profile.c_str(), &wb.handle), 0);
    std::size_t second = 0;
    ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", 1, &second), 0);
    for (int i = 0; i < 3; ++i) {
      ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, static_cast<uint32_t>(i), "h"), 0);
    }
    fm_table_spec_t spec = FM_STRUCT_INIT(fm_table_spec_t);
    spec.range = fm_cell_rect_t{0U, 0U, 2U, 2U};
    spec.name = "Sales";
    std::size_t table = 0;
    ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &table), 0) << fm_get_last_error_message();
    ASSERT_EQ(DefineName(wb.handle, "Taken", "=Sheet1!$A$1", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);

    struct Case {
      const char* name;
      const char* formula;
      int32_t scope;
    };
    const Case cases[] = {
        {"_xlnm.Print_Area", "=$A$1:$B$2", FM_DEFINED_NAME_SCOPE_WORKBOOK},  // new built-in: to the active sheet
        {"_xlnm.Print_Titles", "=$1:$1", FM_DEFINED_NAME_SCOPE_WORKBOOK},    // same for the titles
        {"Sales", "=Sheet1!$A$1", FM_DEFINED_NAME_SCOPE_WORKBOOK},           // a table's name
        {"sales", "=Sheet1!$A$1", 1},                                        // a table's name, sheet scope
        {"Fresh", "=Sheet1!$A$1", 7},                                        // a scope past the last sheet
        {"Fresh", "=Sheet1!$A$1", -2},                                       // a scope below workbook scope
        {"Fresh", "", FM_DEFINED_NAME_SCOPE_WORKBOOK},                       // empty formula
        {"Fresh", "=", FM_DEFINED_NAME_SCOPE_WORKBOOK},                      // formula that is only the `=`
        {"Fresh", "=SUM(", FM_DEFINED_NAME_SCOPE_WORKBOOK},                  // parse failure
        {"", "=Sheet1!$A$1", FM_DEFINED_NAME_SCOPE_WORKBOOK},                // empty name
        {"1Fresh", "=Sheet1!$A$1", FM_DEFINED_NAME_SCOPE_WORKBOOK},          // malformed name
        {"Taken", "=Sheet1!$B$1", FM_DEFINED_NAME_SCOPE_WORKBOOK},           // redefinition: accepted
        {"taken", "=Sheet1!$B$1", FM_DEFINED_NAME_SCOPE_WORKBOOK},           // redefinition in another case: accepted
        {"Fresh", "=Sheet1!$A$1", 1},                                        // a plain new name
    };
    int refused = 0;
    for (const Case& c : cases) {
      SCOPED_TRACE(std::string(c.name) + " " + c.formula + " " + std::to_string(c.scope));
      const std::size_t before = wb.handle->workbook().defined_names().size();
      std::string key_name;
      int32_t key_scope = 99;
      const fm_status_t resolved = ResolveKey(wb.handle, c.name, c.formula, c.scope, &key_name, &key_scope);
      EXPECT_EQ(wb.handle->workbook().defined_names().size(), before);
      const fm_status_t set = DefineName(wb.handle, c.name, c.formula, c.scope);
      EXPECT_EQ(resolved, set);
      refused += set != 0 ? 1 : 0;
      if (set == 0) {
        const auto& names = wb.handle->workbook().defined_names();
        const bool stored_at_key = std::any_of(names.begin(), names.end(), [&](const formulon::DefinedName& n) {
          return n.name == key_name && n.local_sheet_id == key_scope;
        });
        EXPECT_TRUE(stored_at_key);
      }
    }
    EXPECT_EQ(refused, 9);
  }
}

TEST(FormulonCApiPrintSettings, ResolverReportsTheStatusSetReports) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-en_US", &wb.handle), 0);
  EXPECT_EQ(ResolveKey(wb.handle, "Fresh", "=Sheet1!$A$1", 5), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(ResolveKey(wb.handle, "Fresh", "", FM_DEFINED_NAME_SCOPE_WORKBOOK), kInvalidArgument);
  const fm_status_t parse = ResolveKey(wb.handle, "Fresh", "=SUM(", FM_DEFINED_NAME_SCOPE_WORKBOOK);
  EXPECT_GE(parse, 1000);
  EXPECT_LE(parse, 1014);

  fm_defined_name_t input = FM_STRUCT_INIT(fm_defined_name_t);
  input.name = "Fresh";
  input.formula = "=Sheet1!$A$1";
  input.local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK;
  const char* out_name = nullptr;
  int32_t out_scope = 0;
  EXPECT_EQ(fm_workbook_resolve_defined_name_key(nullptr, &input, &out_name, &out_scope),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_resolve_defined_name_key(wb.handle, nullptr, &out_name, &out_scope),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_resolve_defined_name_key(wb.handle, &input, nullptr, &out_scope),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_resolve_defined_name_key(wb.handle, &input, &out_name, nullptr),
            FM_STATUS_BINDING_NULL_POINTER);
  input.formula = nullptr;
  EXPECT_EQ(fm_workbook_resolve_defined_name_key(wb.handle, &input, &out_name, &out_scope),
            FM_STATUS_BINDING_NULL_POINTER);
  input.formula = "=Sheet1!$A$1";
  input.name = "\xff";
  EXPECT_EQ(fm_workbook_resolve_defined_name_key(wb.handle, &input, &out_name, &out_scope),
            FM_STATUS_BINDING_UTF8ENCODING_ERROR);
}

// Resolve, read the entry the set would replace, and set it back with its attributes.
TEST(FormulonCApiPrintSettings, ResolvedKeyLetsACallerKeepTheAttributesASetResets) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-de_DE", &wb.handle), 0);
  std::size_t second = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", 1, &second), 0);
  ASSERT_EQ(DefineName(wb.handle, "Druckbereich", "=Sheet2!$A$1:$B$2", FM_DEFINED_NAME_SCOPE_WORKBOOK, "note", 1), 0)
      << fm_get_last_error_message();

  std::string key_name;
  int32_t key_scope = 0;
  ASSERT_EQ(
      ResolveKey(wb.handle, "DRUCKBEREICH", "=Sheet2!$A$1:$C$3", FM_DEFINED_NAME_SCOPE_WORKBOOK, &key_name, &key_scope),
      0);
  EXPECT_EQ(key_name, "_xlnm.Print_Area");
  EXPECT_EQ(key_scope, 0);
  int32_t present = 0;
  fm_defined_name_t prev = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, key_name.c_str(), key_scope, &present, &prev), 0);
  ASSERT_EQ(present, 1);
  const std::string comment = prev.comment == nullptr ? std::string() : prev.comment;
  const int32_t hidden = prev.hidden;

  // A plain set resets the attributes; the recipe keeps them.
  ASSERT_EQ(DefineName(wb.handle, "DRUCKBEREICH", "=Sheet2!$A$1:$C$3", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  EXPECT_EQ(wb.handle->workbook().defined_names().front().comment, "");
  EXPECT_FALSE(wb.handle->workbook().defined_names().front().hidden);
  ASSERT_EQ(DefineName(wb.handle, "DRUCKBEREICH", "=Sheet2!$A$1:$C$3", FM_DEFINED_NAME_SCOPE_WORKBOOK, comment.c_str(),
                       hidden),
            0);
  EXPECT_EQ(StoredNames(wb.handle), (NameList{{"_xlnm.Print_Area", 0}}));
  EXPECT_EQ(wb.handle->workbook().defined_names().front().comment, "note");
  EXPECT_TRUE(wb.handle->workbook().defined_names().front().hidden);
}

TEST(FormulonCApiPrintSettings, LocalizeAndCanonicalizeDefinedNameShowTheLocaleSpelling) {
  struct Case {
    const char* profile;
    const char* stored;
    const char* shown;
  };
  const Case cases[] = {
      {"mac-365-de_DE", "_xlnm.Print_Area", "Druckbereich"}, {"mac-365-de_DE", "_xlnm.PRINT_AREA", "Druckbereich"},
      {"mac-365-de_DE", "_xlnm.Print_Titles", "Drucktitel"}, {"mac-365-fr_FR", "_xlnm.Print_Area", "Zone_d_impression"},
      {"mac-365-ja_JP", "_xlnm.Print_Area", "Print_Area"},   {"mac-365-en_US", "_xlnm.Print_Area", "Print_Area"},
      {"mac-365-en_US", "_xlnm.print_area", "print_area"},   {"mac-365-en_US", "_xlnm.Foo", "_xlnm.Foo"},
      {"mac-365-de_DE", "_xlnm.Foo", "_xlnm.Foo"},           {"mac-365-de_DE", "Sales", "Sales"},
      {"mac-365-en_US", "Print_Area", "Print_Area"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(std::string(c.profile) + " " + c.stored);
    EXPECT_EQ(Localized(c.stored, c.profile), c.shown);
  }
  // The inverse: the locale spelling in any letter case; en-US keeps the case typed.
  EXPECT_EQ(Canonical("Druckbereich", "mac-365-de_DE"), "_xlnm.Print_Area");
  EXPECT_EQ(Canonical("druckbereich", "mac-365-de_DE"), "_xlnm.Print_Area");
  EXPECT_EQ(Canonical("Zone_d_impression", "mac-365-fr_FR"), "_xlnm.Print_Area");
  EXPECT_EQ(Canonical("Print_Area", "mac-365-de_DE"), "Print_Area");
  EXPECT_EQ(Canonical("Print_Area", "mac-365-ja_JP"), "_xlnm.Print_Area");
  EXPECT_EQ(Canonical("print_area", "mac-365-en_US"), "_xlnm.print_area");
  EXPECT_EQ(Canonical("PRINT_AREA", "mac-365-en_US"), "_xlnm.PRINT_AREA");
  EXPECT_EQ(Canonical("_xlnm.Print_Area", "mac-365-de_DE"), "_xlnm.Print_Area");
  EXPECT_EQ(Canonical("_xlnm.Foo", "mac-365-en_US"), "_xlnm.Foo");
  EXPECT_EQ(Canonical("Sales", "mac-365-en_US"), "Sales");
}

TEST(FormulonCApiPrintSettings, CanonicalizeOfLocalizeRoundTripsEveryBuiltIn) {
  for (const std::string& profile : MacProfiles()) {
    for (const std::string& english : BuiltInNames()) {
      const std::string stored = std::string("_xlnm.") + english;
      SCOPED_TRACE(profile + " " + stored);
      EXPECT_EQ(Canonical(Localized(stored.c_str(), profile).c_str(), profile), stored);
    }
  }
}

TEST(FormulonCApiPrintSettings, DefinedNameSpellingRefusesBadArguments) {
  const char* out = nullptr;
  EXPECT_EQ(fm_localize_defined_name("", "mac-365-en_US", &out), kInvalidArgument);
  EXPECT_EQ(fm_canonicalize_defined_name("", "mac-365-en_US", &out), kInvalidArgument);
  EXPECT_EQ(fm_localize_defined_name("Print_Area", "no-such-profile", &out), kInvalidArgument);
  EXPECT_EQ(fm_canonicalize_defined_name("Print_Area", "no-such-profile", &out), kInvalidArgument);
  EXPECT_EQ(fm_localize_defined_name(nullptr, "mac-365-en_US", &out), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_localize_defined_name("Print_Area", nullptr, &out), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_localize_defined_name("Print_Area", "mac-365-en_US", nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_canonicalize_defined_name("\xff", "mac-365-en_US", &out), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
}

TEST(FormulonCApiPrintSettings, EmptyPrintAreaRemovesTheDefinedName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t area{0, 0, 7, 5};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &area, 1), 0);
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, nullptr, 0), 0);
  EXPECT_TRUE(wb.handle->workbook().defined_names().empty());

  std::size_t count = 9;
  ASSERT_EQ(fm_sheet_get_print_area(wb.handle, 0, nullptr, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPrintSettings, GetPrintAreaSurfacesAMalformedDefinedName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // A file can arrive with a print area this engine cannot parse. Reporting
  // it as "no print area" would silently print the whole used range.
  ASSERT_TRUE(
      static_cast<bool>(wb.handle->workbook().set_defined_name_scoped("_xlnm.Print_Area", "Sheet1!$A$1:$$$", 0)));
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_print_area(wb.handle, 0, nullptr, 0, &count), kPrintInvalidArea);
}

TEST(FormulonCApiPrintSettings, PrintTitlesRoundTripRowsBeforeColumns) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
  titles.has_rows = 1;
  titles.first_row = 0;
  titles.last_row = 1;
  titles.has_columns = 1;
  titles.first_col = 0;
  titles.last_col = 0;
  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, &titles), 0);
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$1:$2,Sheet1!$A:$A");

  fm_print_titles_t read = FM_STRUCT_INIT(fm_print_titles_t);
  ASSERT_EQ(fm_sheet_get_print_titles(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.has_rows, 1);
  EXPECT_EQ(read.first_row, 0U);
  EXPECT_EQ(read.last_row, 1U);
  EXPECT_EQ(read.has_columns, 1);
  EXPECT_EQ(read.first_col, 0U);
  EXPECT_EQ(read.last_col, 0U);
}

TEST(FormulonCApiPrintSettings, PrintTitlesAcceptOneAxisAndClearOnNull) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
  titles.has_rows = 1;
  // The span of an axis whose flag is zero is ignored.
  titles.first_col = 7;
  titles.last_col = 2;
  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, &titles), 0);
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$1:$1");

  fm_print_titles_t read = FM_STRUCT_INIT(fm_print_titles_t);
  ASSERT_EQ(fm_sheet_get_print_titles(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.has_rows, 1);
  EXPECT_EQ(read.first_row, 0U);
  EXPECT_EQ(read.last_row, 0U);
  EXPECT_EQ(read.has_columns, 0);
  EXPECT_EQ(read.first_col, 0U);
  EXPECT_EQ(read.last_col, 0U);

  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, nullptr), 0);
  EXPECT_TRUE(wb.handle->workbook().defined_names().empty());
  ASSERT_EQ(fm_sheet_get_print_titles(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.has_rows, 0);
  EXPECT_EQ(read.has_columns, 0);

  // Both flags zero clears as NULL does.
  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, &titles), 0);
  const fm_print_titles_t none = FM_STRUCT_INIT(fm_print_titles_t);
  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, &none), 0);
  EXPECT_TRUE(wb.handle->workbook().defined_names().empty());
}

TEST(FormulonCApiPrintSettings, PrintTitlesRejectInvertedOrOutOfGridSpans) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_print_titles_t inverted = FM_STRUCT_INIT(fm_print_titles_t);
  inverted.has_rows = 1;
  inverted.first_row = 2;
  inverted.last_row = 1;
  EXPECT_EQ(fm_sheet_set_print_titles(wb.handle, 0, &inverted), kInvalidArgument);
  fm_print_titles_t past_rows = FM_STRUCT_INIT(fm_print_titles_t);
  past_rows.has_rows = 1;
  past_rows.last_row = 1048576U;
  EXPECT_EQ(fm_sheet_set_print_titles(wb.handle, 0, &past_rows), kInvalidArgument);
  fm_print_titles_t past_cols = FM_STRUCT_INIT(fm_print_titles_t);
  past_cols.has_columns = 1;
  past_cols.last_col = 16384U;
  EXPECT_EQ(fm_sheet_set_print_titles(wb.handle, 0, &past_cols), kInvalidArgument);
  EXPECT_TRUE(wb.handle->workbook().defined_names().empty());
  EXPECT_EQ(fm_sheet_get_print_titles(wb.handle, 0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
}

// The enum has two members; CLEAR removes the attribute instead.
TEST(FormulonCApiPrintSettings, PageSetupRejectsOrientationOutsideTheEnum) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::string before = GetPageSetupXml(wb.handle);
  fm_page_setup_patch_t setup = FM_STRUCT_INIT(fm_page_setup_patch_t);
  setup.orientation_op = FM_PATCH_OP_SET;
  for (const int32_t bad : {-1, 2}) {
    setup.orientation = bad;
    EXPECT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &setup), kInvalidArgument) << bad;
  }
  EXPECT_EQ(GetPageSetupXml(wb.handle), before);
}

/* -------------------------------------------------------------------------- */
/* Manual page breaks                                                         */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, BreaksUpsertStaySortedAndReplaceInPlace) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 40, FM_TRISTATE_TRUE), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 10, FM_TRISTATE_TRUE), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 25, FM_TRISTATE_FALSE), 0);
  ASSERT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 3U);

  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.row, 10U);
  EXPECT_EQ(brk.first, 0U);
  EXPECT_EQ(brk.last, 16383U);
  EXPECT_EQ(brk.manual, FM_TRISTATE_TRUE);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.row, 25U);
  EXPECT_EQ(brk.manual, FM_TRISTATE_FALSE);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 2, &brk), 0);
  EXPECT_EQ(brk.row, 40U);

  // Re-adding an existing index replaces it rather than duplicating.
  ASSERT_EQ(SetRowBreak(wb.handle, 25, FM_TRISTATE_TRUE), 0);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 3U);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.row, 25U);
  EXPECT_EQ(brk.manual, FM_TRISTATE_TRUE);
}

TEST(FormulonCApiPrintSettings, LoadedBreaksAreNormalisedBeforeTheAbiObservesThem) {
  // A third-party writer is free to spell `<brk>` in authoring order and
  // to repeat an index. The mutation API binary-searches these vectors
  // and the enumeration API documents ascending order, so a document
  // that arrives unsorted would make `fm_sheet_set_column_break` match
  // nothing and an upsert append a duplicate index.
  formulon::Workbook source = formulon::Workbook::create();
  formulon::SheetPrintSettings& print = source.sheet(0).mutable_print_settings();
  for (const std::uint32_t id : {5U, 30U, 12U, 30U}) {
    formulon::ManualBreak entry;
    entry.id = id;
    entry.max = formulon::Sheet::kMaxCols - 1U;
    entry.manual = true;
    print.manual_col_breaks.push_back(entry);
  }
  const auto bytes = source.save();
  ASSERT_TRUE(bytes.has_value());

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(bytes.value().data(), bytes.value().size(), nullptr, &wb.handle), 0);

  // The duplicate is gone and what remains is strictly increasing.
  ASSERT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 3U);
  fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
  std::uint32_t previous = 0;
  for (std::size_t i = 0; i < 3U; ++i) {
    ASSERT_EQ(fm_sheet_get_column_break_at(wb.handle, 0, i, &brk), 0);
    if (i > 0) {
      EXPECT_GT(brk.col, previous);
    }
    previous = brk.col;
  }
  ASSERT_EQ(fm_sheet_get_column_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.col, 12U);

  // The mutators now find what the enumeration reports.
  EXPECT_EQ(fm_sheet_set_column_break(wb.handle, 0, 12, nullptr), 0);
  EXPECT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 2U);
  ASSERT_EQ(SetColumnBreak(wb.handle, 12, FM_TRISTATE_TRUE), 0);
  EXPECT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 3U);
  ASSERT_EQ(fm_sheet_get_column_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.col, 12U);
}

TEST(FormulonCApiPrintSettings, LoadedBreakCountStaysWithinTheAuthoringCap) {
  // Reading a file must not leave a sheet holding more breaks than
  // `fm_sheet_set_*_break` accepts, or adding one break to a loaded
  // workbook would fail where the same edit on a fresh sheet succeeds.
  formulon::Workbook source = formulon::Workbook::create();
  formulon::SheetPrintSettings& print = source.sheet(0).mutable_print_settings();
  const std::size_t over_cap = formulon::kMaxManualBreaksPerAxis + 8U;
  for (std::size_t i = 0; i < over_cap; ++i) {
    formulon::ManualBreak entry;
    entry.id = static_cast<std::uint32_t>(over_cap - i);  // descending
    entry.max = formulon::Sheet::kMaxCols - 1U;
    entry.manual = true;
    print.manual_row_breaks.push_back(entry);
  }
  const auto bytes = source.save();
  ASSERT_TRUE(bytes.has_value());

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(bytes.value().data(), bytes.value().size(), nullptr, &wb.handle), 0);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), formulon::kMaxManualBreaksPerAxis);

  // An index already present is an in-place replacement, so it is
  // accepted at the cap; a new one is what the cap refuses.
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(SetRowBreak(wb.handle, brk.row, FM_TRISTATE_FALSE), 0);
  EXPECT_EQ(SetRowBreak(wb.handle, formulon::Sheet::kMaxRows - 1U, FM_TRISTATE_TRUE), kInvalidArgument);
}

TEST(FormulonCApiPrintSettings, BreakCountReportsARejectedSheetThroughItsStatus) {
  // An empty axis and a rejected sheet both leave nothing to enumerate; the
  // status is what separates them, with the diagnostic agreeing.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::size_t sheet_count = count_of(fm_workbook_get_sheet_count, wb.handle);
  constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  std::size_t count = 99U;

  // Empty axis: zero, and no diagnostic.
  EXPECT_EQ(fm_sheet_get_row_break_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  EXPECT_STREQ(fm_get_last_error_message(), "");
  EXPECT_EQ(fm_sheet_get_column_break_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  EXPECT_STREQ(fm_get_last_error_message(), "");

  // Sheet past the end: a rejection, and a diagnostic naming the entry point.
  EXPECT_EQ(fm_sheet_get_row_break_count(wb.handle, sheet_count, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_STRNE(fm_get_last_error_message(), "");
  EXPECT_EQ(fm_sheet_get_column_break_count(wb.handle, sheet_count, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_STRNE(fm_get_last_error_message(), "");

  // NULL handle and NULL out-param.
  EXPECT_EQ(fm_sheet_get_row_break_count(nullptr, 0, &count), kNullPointer);
  EXPECT_STRNE(fm_get_last_error_message(), "");
  EXPECT_EQ(fm_sheet_get_row_break_count(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_STRNE(fm_get_last_error_message(), "");

  // A populated axis clears the previous rejection rather than inheriting
  // it, so the empty message stays a trustworthy success signal.
  ASSERT_EQ(SetRowBreak(wb.handle, 10, FM_TRISTATE_TRUE), 0);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 1U);
  EXPECT_STREQ(fm_get_last_error_message(), "");

  // The indexed reader rejects the same sheet outright.
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  EXPECT_EQ(fm_sheet_get_row_break_at(wb.handle, sheet_count, 0, &brk), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  fm_column_break_t cbrk = FM_STRUCT_INIT(fm_column_break_t);
  EXPECT_EQ(fm_sheet_get_column_break_at(wb.handle, sheet_count, 0, &cbrk), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiPrintSettings, ColumnBreaksSpanEveryRow) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetColumnBreak(wb.handle, 5, FM_TRISTATE_TRUE), 0);
  fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
  ASSERT_EQ(fm_sheet_get_column_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.col, 5U);
  EXPECT_EQ(brk.last, 1048575U);
}

TEST(FormulonCApiPrintSettings, RemovingAnAbsentBreakIsANoOp) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 0, 7, nullptr), 0);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 0U);
  ASSERT_EQ(SetRowBreak(wb.handle, 7, FM_TRISTATE_TRUE), 0);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 0, 7, nullptr), 0);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 0U);
}

TEST(FormulonCApiPrintSettings, ClearBreaksEmptiesBothAxes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 3, FM_TRISTATE_TRUE), 0);
  ASSERT_EQ(SetColumnBreak(wb.handle, 4, FM_TRISTATE_TRUE), 0);
  ASSERT_EQ(fm_sheet_clear_breaks(wb.handle, 0), 0);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 0U);
  EXPECT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 0U);
}

TEST(FormulonCApiPrintSettings, BreakRemovalChecksTheGridBeforeLookingUpTheBreak) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 0, formulon::Sheet::kMaxRows, nullptr), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_break(wb.handle, 0, formulon::Sheet::kMaxCols, nullptr), kInvalidArgument);

  // The last grid row / column is addressable: added, then removed.
  ASSERT_EQ(SetRowBreak(wb.handle, formulon::Sheet::kMaxRows - 1U, FM_TRISTATE_TRUE), 0);
  ASSERT_EQ(SetColumnBreak(wb.handle, formulon::Sheet::kMaxCols - 1U, FM_TRISTATE_TRUE), 0);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 0, formulon::Sheet::kMaxRows - 1U, nullptr), 0);
  EXPECT_EQ(fm_sheet_set_column_break(wb.handle, 0, formulon::Sheet::kMaxCols - 1U, nullptr), 0);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 0, formulon::Sheet::kMaxRows - 1U, nullptr), 0);
}

TEST(FormulonCApiPrintSettings, BreaksRejectOutOfGridAndOverflow) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(SetRowBreak(wb.handle, formulon::Sheet::kMaxRows, FM_TRISTATE_TRUE), kInvalidArgument);
  EXPECT_EQ(SetColumnBreak(wb.handle, formulon::Sheet::kMaxCols, FM_TRISTATE_TRUE), kInvalidArgument);

  for (std::uint32_t row = 0; row < 1026U; ++row) {
    ASSERT_EQ(SetRowBreak(wb.handle, row, FM_TRISTATE_TRUE), 0) << "row " << row;
  }
  // Excel's own ceiling; going past it produces a file Excel truncates.
  EXPECT_EQ(SetRowBreak(wb.handle, 2000, FM_TRISTATE_TRUE), kInvalidArgument);
  EXPECT_STREQ(fm_get_last_error_context(), "part=manual_breaks requested=1027 limit=1026");
  // Replacing an existing break is still allowed at the ceiling.
  EXPECT_EQ(SetRowBreak(wb.handle, 0, FM_TRISTATE_FALSE), 0);
}

TEST(FormulonCApiPrintSettings, BreakAtRejectsOutOfRangeIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  EXPECT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_column_break_t cbrk = FM_STRUCT_INIT(fm_column_break_t);
  EXPECT_EQ(fm_sheet_get_column_break_at(wb.handle, 0, 0, &cbrk), FM_STATUS_INDEX_OUT_OF_RANGE);
}

/* -------------------------------------------------------------------------- */
/* Typed patch setters                                                        */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, PageSetupPatchPreservesUnmodelledAttributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(
      fm_sheet_set_page_setup_xml(
          wb.handle, 0, "<pageSetup paperSize=\"9\" orientation=\"portrait\" horizontalDpi=\"600\" copies=\"3\"/>"),
      0);

  fm_page_setup_patch_t patch = FM_STRUCT_INIT(fm_page_setup_patch_t);
  patch.orientation_op = FM_PATCH_OP_SET;
  patch.orientation = FM_PAGE_ORIENTATION_LANDSCAPE;
  ASSERT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &patch), 0);

  const std::string xml = GetPageSetupXml(wb.handle);
  EXPECT_NE(xml.find("orientation=\"landscape\""), std::string::npos) << xml;
  // The engine models neither of these, which is precisely why it must not
  // be the reason they disappear from an Excel-authored file.
  EXPECT_NE(xml.find("horizontalDpi=\"600\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("copies=\"3\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("paperSize=\"9\""), std::string::npos) << xml;
}

TEST(FormulonCApiPrintSettings, PageSetupPatchRejectsScaleOutsideExcelsRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_page_setup_patch_t patch = FM_STRUCT_INIT(fm_page_setup_patch_t);
  patch.scale_pct_op = FM_PATCH_OP_SET;
  patch.scale_pct = 9;
  // A mis-stated print scale lands on paper, so it is rejected rather than
  // clamped.
  EXPECT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &patch), kInvalidArgument);
  patch.scale_pct = 401;
  EXPECT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &patch), kInvalidArgument);
  EXPECT_EQ(GetPageSetupXml(wb.handle), "");

  patch.scale_pct = 400;
  EXPECT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &patch), 0);
}

TEST(FormulonCApiPrintSettings, PageSetupPatchRoutesFitToPageToSheetPr) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_page_setup_patch_t patch = FM_STRUCT_INIT(fm_page_setup_patch_t);
  patch.fit_to_page_op = FM_PATCH_OP_SET;
  patch.fit_to_page = 1;
  patch.fit_to_width_op = FM_PATCH_OP_SET;
  patch.fit_to_width = 1;
  patch.fit_to_height_op = FM_PATCH_OP_SET;
  patch.fit_to_height = 0;
  ASSERT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &patch), 0);

  // `fitToPage` selects the mode and lives in `<sheetPr>`; the page counts
  // that state the target live in `<pageSetup>`. Both halves are needed.
  EXPECT_NE(GetSheetPrXml(wb.handle).find("fitToPage=\"1\""), std::string::npos);
  const std::string setup_xml = GetPageSetupXml(wb.handle);
  EXPECT_NE(setup_xml.find("fitToWidth=\"1\""), std::string::npos) << setup_xml;
  EXPECT_NE(setup_xml.find("fitToHeight=\"0\""), std::string::npos) << setup_xml;
}

TEST(FormulonCApiPrintSettings, PageMarginsPatchTouchesOnlySetFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_page_margins_patch_t patch = FM_STRUCT_INIT(fm_page_margins_patch_t);
  patch.left_in_op = FM_PATCH_OP_SET;
  patch.left_in = 0.25;
  patch.top_in_op = FM_PATCH_OP_SET;
  patch.top_in = 1.5;
  ASSERT_EQ(fm_sheet_update_page_margins(wb.handle, 0, &patch), 0);

  fm_page_margins_t read = FM_STRUCT_INIT(fm_page_margins_t);
  ASSERT_EQ(fm_sheet_get_page_margins(wb.handle, 0, &read), 0);
  EXPECT_DOUBLE_EQ(read.left_in, 0.25);
  EXPECT_DOUBLE_EQ(read.top_in, 1.5);
  EXPECT_EQ(read.has_left_in, 1);
  EXPECT_EQ(read.has_top_in, 1);
  // An unstated margin reports the OOXML default value with its presence
  // flag clear, so the caller can tell "0.7 because the file says so" from
  // "0.7 because nothing says otherwise".
  EXPECT_DOUBLE_EQ(read.right_in, 0.7);
  EXPECT_EQ(read.has_right_in, 0);
}

TEST(FormulonCApiPrintSettings, PageMarginsPatchRejectsNonFiniteAndNegative) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_page_margins_patch_t patch = FM_STRUCT_INIT(fm_page_margins_patch_t);
  patch.bottom_in_op = FM_PATCH_OP_SET;
  patch.bottom_in = -1.0;
  // The paginator subtracts margins from the paper; a negative one
  // inflates the printable body past the sheet.
  EXPECT_EQ(fm_sheet_update_page_margins(wb.handle, 0, &patch), kInvalidArgument);
  patch.bottom_in = 1.0 / 0.0;
  EXPECT_EQ(fm_sheet_update_page_margins(wb.handle, 0, &patch), kInvalidArgument);
}

TEST(FormulonCApiPrintSettings, PrintOptionsAndHeaderFooterPatch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_print_options_patch_t options = FM_STRUCT_INIT(fm_print_options_patch_t);
  options.grid_lines_op = FM_PATCH_OP_SET;
  options.grid_lines = 1;
  options.horizontal_centered_op = FM_PATCH_OP_SET;
  options.horizontal_centered = 1;
  ASSERT_EQ(fm_sheet_update_print_options(wb.handle, 0, &options), 0);
  const char* options_xml = nullptr;
  ASSERT_EQ(fm_sheet_get_print_options_xml(wb.handle, 0, &options_xml), 0);
  EXPECT_STREQ(options_xml, "<printOptions gridLines=\"1\" horizontalCentered=\"1\"/>");

  fm_header_footer_patch_t hf = FM_STRUCT_INIT(fm_header_footer_patch_t);
  hf.odd_header_op = FM_PATCH_OP_SET;
  hf.odd_header = "&C\xE5\xB8\xB3\xE7\xA5\xA8";  // "&C帳票"
  hf.odd_footer_op = FM_PATCH_OP_SET;
  hf.odd_footer = "&R&P / &N";
  ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &hf), 0);
  const char* hf_xml = nullptr;
  ASSERT_EQ(fm_sheet_get_header_footer_xml(wb.handle, 0, &hf_xml), 0);
  // `&` introduces Excel's formatting codes, and it reaches the file
  // escaped: what a parser hands back is `&C...`, which is the code.
  EXPECT_STREQ(hf_xml,
               "<headerFooter><oddHeader>&amp;C\xE5\xB8\xB3\xE7\xA5\xA8</oddHeader>"
               "<oddFooter>&amp;R&amp;P / &amp;N</oddFooter></headerFooter>");
}

TEST(FormulonCApiPrintSettings, HeaderFooterSectionsLandInSchemaOrder) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // Written last-to-first on purpose: ECMA-376 fixes the child order, and
  // appending in call order would put `<oddHeader>` after `<firstFooter>`.
  fm_header_footer_patch_t late = FM_STRUCT_INIT(fm_header_footer_patch_t);
  late.first_footer_op = FM_PATCH_OP_SET;
  late.first_footer = "F";
  ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &late), 0);
  fm_header_footer_patch_t early = FM_STRUCT_INIT(fm_header_footer_patch_t);
  early.odd_header_op = FM_PATCH_OP_SET;
  early.odd_header = "H";
  early.even_footer_op = FM_PATCH_OP_SET;
  early.even_footer = "E";
  ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &early), 0);

  const char* xml = nullptr;
  ASSERT_EQ(fm_sheet_get_header_footer_xml(wb.handle, 0, &xml), 0);
  EXPECT_STREQ(xml,
               "<headerFooter><oddHeader>H</oddHeader><evenFooter>E</evenFooter>"
               "<firstFooter>F</firstFooter></headerFooter>");
}

TEST(FormulonCApiPrintSettings, HeaderFooterKeepLeavesSectionAndEmptySetClearsIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_header_footer_patch_t first = FM_STRUCT_INIT(fm_header_footer_patch_t);
  first.odd_header_op = FM_PATCH_OP_SET;
  first.odd_header = "Keep";
  first.odd_footer_op = FM_PATCH_OP_SET;
  first.odd_footer = "Drop";
  ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &first), 0);

  // KEEP leaves the header; SET with an empty string removes the footer.
  fm_header_footer_patch_t second = FM_STRUCT_INIT(fm_header_footer_patch_t);
  second.odd_footer_op = FM_PATCH_OP_SET;
  second.odd_footer = "";
  ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &second), 0);

  const char* xml = nullptr;
  ASSERT_EQ(fm_sheet_get_header_footer_xml(wb.handle, 0, &xml), 0);
  EXPECT_STREQ(xml, "<headerFooter><oddHeader>Keep</oddHeader></headerFooter>");
}

TEST(FormulonCApiPrintSettings, ClearingFitToPageThroughPageSetupDropsTheStatedAttribute) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetFitToPage(wb.handle, 0, 1), 0);
  ASSERT_EQ(SetFitToPage(wb.handle, 0, 0), 0);
  fm_page_setup_t setup = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &setup), 0);
  EXPECT_EQ(setup.fit_to_page, 0);
  EXPECT_EQ(setup.has_fit_to_page, 0);
  EXPECT_EQ(GetSheetPrXml(wb.handle), "<sheetPr><pageSetUpPr/></sheetPr>");
}

TEST(FormulonCApiPrintSettings, PrintOptionsGetterReportsStatedAndAbsentAttributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_print_options_t fresh = FM_STRUCT_INIT(fm_print_options_t);
  fresh.grid_lines = 7;
  ASSERT_EQ(fm_sheet_get_print_options(wb.handle, 0, &fresh), 0);
  EXPECT_EQ(fresh.grid_lines, 0);
  EXPECT_EQ(fresh.has_grid_lines, 0);
  EXPECT_EQ(fresh.has_headings, 0);
  EXPECT_EQ(fresh.has_horizontal_centered, 0);
  EXPECT_EQ(fresh.has_vertical_centered, 0);

  fm_print_options_patch_t patch = FM_STRUCT_INIT(fm_print_options_patch_t);
  patch.grid_lines_op = FM_PATCH_OP_SET;
  patch.grid_lines = 1;
  patch.headings_op = FM_PATCH_OP_SET;
  patch.headings = 0;
  ASSERT_EQ(fm_sheet_update_print_options(wb.handle, 0, &patch), 0);

  fm_print_options_t read = FM_STRUCT_INIT(fm_print_options_t);
  ASSERT_EQ(fm_sheet_get_print_options(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.grid_lines, 1);
  EXPECT_EQ(read.has_grid_lines, 1);
  // A stated `false` is distinct from an absent attribute.
  EXPECT_EQ(read.headings, 0);
  EXPECT_EQ(read.has_headings, 1);
  EXPECT_EQ(read.horizontal_centered, 0);
  EXPECT_EQ(read.has_horizontal_centered, 0);
  EXPECT_EQ(read.has_vertical_centered, 0);

  ASSERT_EQ(fm_sheet_set_print_options_xml(wb.handle, 0, "<printOptions verticalCentered=\"1\"/>"), 0);
  ASSERT_EQ(fm_sheet_get_print_options(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.vertical_centered, 1);
  EXPECT_EQ(read.has_vertical_centered, 1);
  EXPECT_EQ(read.has_grid_lines, 0);
}

TEST(FormulonCApiPrintSettings, HeaderFooterGetterReadsDecodedSectionsAndNullForAbsentOnes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_header_footer_t fresh = FM_STRUCT_INIT(fm_header_footer_t);
  ASSERT_EQ(fm_sheet_get_header_footer(wb.handle, 0, &fresh), 0);
  EXPECT_EQ(fresh.odd_header, nullptr);
  EXPECT_EQ(fresh.first_footer, nullptr);
  EXPECT_EQ(fresh.different_odd_even, 0);
  EXPECT_EQ(fresh.has_different_odd_even, 0);
  // CT_HeaderFooter defaults these two to true.
  EXPECT_EQ(fresh.scale_with_doc, 1);
  EXPECT_EQ(fresh.has_scale_with_doc, 0);
  EXPECT_EQ(fresh.align_with_margins, 1);
  EXPECT_EQ(fresh.has_align_with_margins, 0);

  fm_header_footer_patch_t patch = FM_STRUCT_INIT(fm_header_footer_patch_t);
  patch.odd_header_op = FM_PATCH_OP_SET;
  patch.odd_header = "&C\xE5\xB8\xB3\xE7\xA5\xA8";  // "&C帳票"
  patch.even_footer_op = FM_PATCH_OP_SET;
  patch.even_footer = "&R&P / &N";
  patch.different_odd_even_op = FM_PATCH_OP_SET;
  patch.different_odd_even = 1;
  patch.scale_with_doc_op = FM_PATCH_OP_SET;
  patch.scale_with_doc = 0;
  ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &patch), 0);

  fm_header_footer_t read = FM_STRUCT_INIT(fm_header_footer_t);
  ASSERT_EQ(fm_sheet_get_header_footer(wb.handle, 0, &read), 0);
  ASSERT_NE(read.odd_header, nullptr);
  // The text comes back decoded, in the form the setter takes.
  EXPECT_STREQ(read.odd_header, "&C\xE5\xB8\xB3\xE7\xA5\xA8");
  ASSERT_NE(read.even_footer, nullptr);
  EXPECT_STREQ(read.even_footer, "&R&P / &N");
  EXPECT_EQ(read.odd_footer, nullptr);
  EXPECT_EQ(read.even_header, nullptr);
  EXPECT_EQ(read.first_header, nullptr);
  EXPECT_EQ(read.first_footer, nullptr);
  EXPECT_EQ(read.different_odd_even, 1);
  EXPECT_EQ(read.has_different_odd_even, 1);
  EXPECT_EQ(read.scale_with_doc, 0);
  EXPECT_EQ(read.has_scale_with_doc, 1);
  EXPECT_EQ(read.has_different_first, 0);
  EXPECT_EQ(read.align_with_margins, 1);
  EXPECT_EQ(read.has_align_with_margins, 0);
}

/* -------------------------------------------------------------------------- */
/* Pagination linkage                                                         */
/* -------------------------------------------------------------------------- */

namespace {

/// Paginates sheet 0 and returns its page count.
std::uint32_t PageCount(fm_workbook_t* wb) {
  PaginationGuard pagination;
  EXPECT_EQ(fm_sheet_paginate(wb, 0, &pagination.handle), 0);
  return static_cast<std::uint32_t>(count_of(fm_pagination_get_page_count, pagination.handle));
}

/// Fills `A1:T200` so the sheet needs more than one page either way.
void FillGrid(fm_workbook_t* wb) {
  for (std::uint32_t row = 0; row < 200U; ++row) {
    for (std::uint32_t col = 0; col < 20U; ++col) {
      ASSERT_EQ(fm_sheet_set_number(wb, 0, row, col, 1.0), 0);
    }
  }
}

}  // namespace

TEST(FormulonCApiPrintSettings, OrientationChangeIsVisibleToPaginateImmediately) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);

  fm_page_setup_patch_t portrait = FM_STRUCT_INIT(fm_page_setup_patch_t);
  portrait.orientation_op = FM_PATCH_OP_SET;
  portrait.orientation = FM_PAGE_ORIENTATION_PORTRAIT;
  ASSERT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &portrait), 0);
  const std::uint32_t portrait_pages = PageCount(wb.handle);

  fm_page_setup_patch_t landscape = FM_STRUCT_INIT(fm_page_setup_patch_t);
  landscape.orientation_op = FM_PATCH_OP_SET;
  landscape.orientation = FM_PAGE_ORIENTATION_LANDSCAPE;
  ASSERT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &landscape), 0);
  const std::uint32_t landscape_pages = PageCount(wb.handle);

  // No save/load in between: the setter re-derives the structured view the
  // paginator reads, which is the whole point of routing every mutation
  // through the reader's parser.
  EXPECT_NE(portrait_pages, landscape_pages);
}

TEST(FormulonCApiPrintSettings, RawXmlSetterAlsoDrivesPagination) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const std::uint32_t before = PageCount(wb.handle);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup scale=\"25\"/>"), 0);
  EXPECT_LT(PageCount(wb.handle), before);
}

TEST(FormulonCApiPrintSettings, ManualRowBreakAddsAPage) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 9, 0, 1.0), 0);
  const std::uint32_t before = PageCount(wb.handle);
  ASSERT_EQ(SetRowBreak(wb.handle, 5, FM_TRISTATE_TRUE), 0);
  EXPECT_GT(PageCount(wb.handle), before);
}

TEST(FormulonCApiPrintSettings, PrintAreaNarrowsWhatPaginates) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const std::uint32_t full = PageCount(wb.handle);
  const fm_cell_rect_t area{0, 0, 1, 1};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &area, 1), 0);
  EXPECT_LT(PageCount(wb.handle), full);
  EXPECT_EQ(PageCount(wb.handle), 1U);
}

/* -------------------------------------------------------------------------- */
/* Argument handling                                                          */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, NullArgumentsReportBindingNullPointer) {
  constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_get_page_setup_xml(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_set_print_area(wb.handle, 0, nullptr, 1), kNullPointer);
  std::size_t area_count = 0;
  EXPECT_EQ(fm_sheet_get_print_area(wb.handle, 0, nullptr, 1, &area_count), kNullPointer);
  EXPECT_EQ(fm_sheet_get_print_area(wb.handle, 0, nullptr, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_get_page_setup(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_update_page_setup(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_update_page_margins(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_update_print_options(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_update_header_footer(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_get_print_options(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_get_header_footer(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_get_row_break_count(wb.handle, 0, nullptr), kNullPointer);
}

TEST(FormulonCApiPrintSettings, OutOfRangeSheetIndexIsRejectedEverywhere) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* xml = nullptr;
  EXPECT_EQ(fm_sheet_get_page_setup_xml(wb.handle, 9, &xml), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 9, "<pageSetup/>"), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetFitToPage(wb.handle, 9, 1), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  fm_print_options_t options = FM_STRUCT_INIT(fm_print_options_t);
  EXPECT_EQ(fm_sheet_get_print_options(wb.handle, 9, &options), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  fm_header_footer_t hf = FM_STRUCT_INIT(fm_header_footer_t);
  EXPECT_EQ(fm_sheet_get_header_footer(wb.handle, 9, &hf), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  const fm_cell_rect_t area{0, 0, 1, 1};
  EXPECT_EQ(fm_sheet_set_print_area(wb.handle, 9, &area, 1), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 9, 1, &brk), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 9, 1, nullptr), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_row_break_count(wb.handle, 9, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

/* -------------------------------------------------------------------------- */
/* Unit-suffixed fields, merging updates and record-based breaks              */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, PageSetupUpdatesMergeAcrossCalls) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_page_setup_patch_t first = FM_STRUCT_INIT(fm_page_setup_patch_t);
  first.scale_pct_op = FM_PATCH_OP_SET;
  first.scale_pct = 75U;
  ASSERT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &first), 0);
  fm_page_setup_patch_t second = FM_STRUCT_INIT(fm_page_setup_patch_t);
  second.orientation_op = FM_PATCH_OP_SET;
  second.orientation = FM_PAGE_ORIENTATION_LANDSCAPE;
  ASSERT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &second), 0);

  fm_page_setup_t read = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.scale_pct, 75U);
  EXPECT_EQ(read.has_scale_pct, 1);
  EXPECT_EQ(read.orientation, FM_PAGE_ORIENTATION_LANDSCAPE);
  EXPECT_EQ(read.has_orientation, 1);
  const std::string xml = GetPageSetupXml(wb.handle);
  EXPECT_NE(xml.find("scale=\"75\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("orientation=\"landscape\""), std::string::npos) << xml;
}

TEST(FormulonCApiPrintSettings, PageSetupScalePctDefaultsToOneHundredWhenAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_page_setup_t read = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.scale_pct, 100U);
  EXPECT_EQ(read.has_scale_pct, 0);
}

TEST(FormulonCApiPrintSettings, PageMarginsAreInchesInTheFileAndOnRead) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_page_margins_patch_t patch = FM_STRUCT_INIT(fm_page_margins_patch_t);
  patch.left_in_op = FM_PATCH_OP_SET;
  patch.left_in = 0.5;
  patch.header_in_op = FM_PATCH_OP_SET;
  patch.header_in = 0.25;
  ASSERT_EQ(fm_sheet_update_page_margins(wb.handle, 0, &patch), 0);

  const char* xml = nullptr;
  ASSERT_EQ(fm_sheet_get_page_margins_xml(wb.handle, 0, &xml), 0);
  ASSERT_NE(xml, nullptr);
  EXPECT_NE(std::string(xml).find("left=\"0.5\""), std::string::npos) << xml;
  EXPECT_NE(std::string(xml).find("header=\"0.25\""), std::string::npos) << xml;

  fm_page_margins_t read = FM_STRUCT_INIT(fm_page_margins_t);
  ASSERT_EQ(fm_sheet_get_page_margins(wb.handle, 0, &read), 0);
  EXPECT_DOUBLE_EQ(read.left_in, 0.5);
  EXPECT_DOUBLE_EQ(read.header_in, 0.25);
  EXPECT_EQ(read.has_header_in, 1);
}

TEST(FormulonCApiPrintSettings, PageMarginsUpdatesMergeAcrossCalls) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_page_margins_patch_t first = FM_STRUCT_INIT(fm_page_margins_patch_t);
  first.top_in_op = FM_PATCH_OP_SET;
  first.top_in = 1.25;
  ASSERT_EQ(fm_sheet_update_page_margins(wb.handle, 0, &first), 0);
  fm_page_margins_patch_t second = FM_STRUCT_INIT(fm_page_margins_patch_t);
  second.footer_in_op = FM_PATCH_OP_SET;
  second.footer_in = 0.4;
  ASSERT_EQ(fm_sheet_update_page_margins(wb.handle, 0, &second), 0);

  fm_page_margins_t read = FM_STRUCT_INIT(fm_page_margins_t);
  ASSERT_EQ(fm_sheet_get_page_margins(wb.handle, 0, &read), 0);
  EXPECT_DOUBLE_EQ(read.top_in, 1.25);
  EXPECT_EQ(read.has_top_in, 1);
  EXPECT_DOUBLE_EQ(read.footer_in, 0.4);
  EXPECT_EQ(read.has_footer_in, 1);
  EXPECT_EQ(read.has_left_in, 0);
}

TEST(FormulonCApiPrintSettings, PrintOptionsUpdatesMergeAcrossCalls) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_print_options_patch_t first = FM_STRUCT_INIT(fm_print_options_patch_t);
  first.grid_lines_op = FM_PATCH_OP_SET;
  first.grid_lines = 1;
  ASSERT_EQ(fm_sheet_update_print_options(wb.handle, 0, &first), 0);
  fm_print_options_patch_t second = FM_STRUCT_INIT(fm_print_options_patch_t);
  second.headings_op = FM_PATCH_OP_SET;
  second.headings = 1;
  ASSERT_EQ(fm_sheet_update_print_options(wb.handle, 0, &second), 0);

  fm_print_options_t read = FM_STRUCT_INIT(fm_print_options_t);
  ASSERT_EQ(fm_sheet_get_print_options(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.grid_lines, 1);
  EXPECT_EQ(read.headings, 1);
  EXPECT_EQ(read.has_horizontal_centered, 0);
}

TEST(FormulonCApiPrintSettings, HeaderFooterUpdatesMergeAcrossCalls) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_header_footer_patch_t first = FM_STRUCT_INIT(fm_header_footer_patch_t);
  first.odd_header_op = FM_PATCH_OP_SET;
  first.odd_header = "&CTitle";
  ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &first), 0);
  fm_header_footer_patch_t second = FM_STRUCT_INIT(fm_header_footer_patch_t);
  second.odd_footer_op = FM_PATCH_OP_SET;
  second.odd_footer = "&RPage &P";
  second.different_first_op = FM_PATCH_OP_SET;
  second.different_first = 1;
  ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &second), 0);

  fm_header_footer_t read = FM_STRUCT_INIT(fm_header_footer_t);
  ASSERT_EQ(fm_sheet_get_header_footer(wb.handle, 0, &read), 0);
  ASSERT_NE(read.odd_header, nullptr);
  EXPECT_STREQ(read.odd_header, "&CTitle");
  ASSERT_NE(read.odd_footer, nullptr);
  EXPECT_STREQ(read.odd_footer, "&RPage &P");
  EXPECT_EQ(read.different_first, 1);
}

TEST(FormulonCApiPrintSettings, RowBreakRecordStoresItsSpanAndManualFlag) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 12, FM_TRISTATE_TRUE, 2, 9), 0);
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.row, 12U);
  EXPECT_EQ(brk.first, 2U);
  EXPECT_EQ(brk.last, 9U);
  EXPECT_EQ(brk.manual, FM_TRISTATE_TRUE);

  // Upsert replaces the span and the flag in place.
  ASSERT_EQ(SetRowBreak(wb.handle, 12, FM_TRISTATE_FALSE, 4, 5), 0);
  ASSERT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 1U);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.first, 4U);
  EXPECT_EQ(brk.last, 5U);
  EXPECT_EQ(brk.manual, FM_TRISTATE_FALSE);
}

TEST(FormulonCApiPrintSettings, RowBreakWithZeroSpanCoversTheWholeSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 30, FM_TRISTATE_TRUE, 0, 0), 0);
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.first, 0U);
  EXPECT_EQ(brk.last, formulon::Sheet::kMaxCols - 1U);
}

TEST(FormulonCApiPrintSettings, RowBreakPositionComesFromTheRowArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  brk.row = 99U;  // ignored
  brk.manual = FM_TRISTATE_TRUE;
  ASSERT_EQ(fm_sheet_set_row_break(wb.handle, 0, 7, &brk), 0);
  fm_row_break_t read = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &read), 0);
  EXPECT_EQ(read.row, 7U);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 1U);
}

TEST(FormulonCApiPrintSettings, RowBreakNullRemovesTheBreakAndIsANoOpWhenAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 7, FM_TRISTATE_TRUE), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 8, FM_TRISTATE_TRUE), 0);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 0, 7, nullptr), 0);
  ASSERT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 1U);
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.row, 8U);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 0, 7, nullptr), 0);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 1U);
}

TEST(FormulonCApiPrintSettings, ColumnBreakRecordStoresItsSpanAndManualFlag) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetColumnBreak(wb.handle, 6, FM_TRISTATE_TRUE, 10, 20), 0);
  fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
  ASSERT_EQ(fm_sheet_get_column_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.col, 6U);
  EXPECT_EQ(brk.first, 10U);
  EXPECT_EQ(brk.last, 20U);
  EXPECT_EQ(brk.manual, FM_TRISTATE_TRUE);

  ASSERT_EQ(SetColumnBreak(wb.handle, 6, FM_TRISTATE_FALSE, 0, 0), 0);
  ASSERT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 1U);
  ASSERT_EQ(fm_sheet_get_column_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.first, 0U);
  EXPECT_EQ(brk.last, formulon::Sheet::kMaxRows - 1U);
  EXPECT_EQ(brk.manual, FM_TRISTATE_FALSE);
}

TEST(FormulonCApiPrintSettings, ColumnBreakNullRemovesTheBreakAndIsANoOpWhenAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_column_break(wb.handle, 0, 3, nullptr), 0);
  ASSERT_EQ(SetColumnBreak(wb.handle, 3, FM_TRISTATE_TRUE), 0);
  ASSERT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 1U);
  EXPECT_EQ(fm_sheet_set_column_break(wb.handle, 0, 3, nullptr), 0);
  EXPECT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 0U);
}

TEST(FormulonCApiPrintSettings, BreakRecordsRejectInvertedSpansAndBadFlags) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(SetRowBreak(wb.handle, 5, FM_TRISTATE_TRUE, 9, 2), kInvalidArgument);
  EXPECT_EQ(SetRowBreak(wb.handle, 5, 3), kInvalidArgument);
  EXPECT_EQ(SetRowBreak(wb.handle, 5, FM_TRISTATE_TRUE, 0, formulon::Sheet::kMaxCols), kInvalidArgument);
  EXPECT_EQ(SetColumnBreak(wb.handle, 5, FM_TRISTATE_TRUE, 9, 2), kInvalidArgument);
  EXPECT_EQ(SetColumnBreak(wb.handle, 5, -1), kInvalidArgument);
  EXPECT_EQ(SetColumnBreak(wb.handle, 5, FM_TRISTATE_TRUE, 0, formulon::Sheet::kMaxRows), kInvalidArgument);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 0U);
  EXPECT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 0U);
}

/* -------------------------------------------------------------------------- */
/* Patch op handling                                                          */
/* -------------------------------------------------------------------------- */

namespace {

constexpr fm_patch_op_t kBadOp = static_cast<fm_patch_op_t>(3);

const char* const kPageSetupXml = "<pageSetup paperSize=\"9\" orientation=\"landscape\" scale=\"80\"/>";
const char* const kMarginsXml =
    "<pageMargins left=\"1\" right=\"1\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>";
const char* const kOptionsXml = "<printOptions gridLines=\"1\"/>";
const char* const kHeaderFooterXml =
    "<headerFooter differentFirst=\"1\"><oddHeader>H</oddHeader><oddFooter>F</oddFooter></headerFooter>";

std::string Fragment(fm_status_t (*get)(const fm_workbook_t*, size_t, const char**), const fm_workbook_t* wb) {
  const char* xml = nullptr;
  EXPECT_EQ(get(wb, 0, &xml), 0);
  return xml == nullptr ? std::string() : std::string(xml);
}

}  // namespace

TEST(FormulonCApiPrintPatch, AllKeepPatchesAreNoOps) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, kPageSetupXml), 0);
  ASSERT_EQ(fm_sheet_set_page_margins_xml(wb.handle, 0, kMarginsXml), 0);
  ASSERT_EQ(fm_sheet_set_print_options_xml(wb.handle, 0, kOptionsXml), 0);
  ASSERT_EQ(fm_sheet_set_header_footer_xml(wb.handle, 0, kHeaderFooterXml), 0);
  const fm_page_setup_patch_t setup = FM_STRUCT_INIT(fm_page_setup_patch_t);
  const fm_page_margins_patch_t margins = FM_STRUCT_INIT(fm_page_margins_patch_t);
  const fm_print_options_patch_t options = FM_STRUCT_INIT(fm_print_options_patch_t);
  const fm_header_footer_patch_t hf = FM_STRUCT_INIT(fm_header_footer_patch_t);
  ASSERT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &setup), 0);
  ASSERT_EQ(fm_sheet_update_page_margins(wb.handle, 0, &margins), 0);
  ASSERT_EQ(fm_sheet_update_print_options(wb.handle, 0, &options), 0);
  ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &hf), 0);
  EXPECT_EQ(Fragment(fm_sheet_get_page_setup_xml, wb.handle), kPageSetupXml);
  EXPECT_EQ(Fragment(fm_sheet_get_page_margins_xml, wb.handle), kMarginsXml);
  EXPECT_EQ(Fragment(fm_sheet_get_print_options_xml, wb.handle), kOptionsXml);
  EXPECT_EQ(Fragment(fm_sheet_get_header_footer_xml, wb.handle), kHeaderFooterXml);
}

TEST(FormulonCApiPrintPatch, ClearingOrientationRemovesTheAttribute) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, kPageSetupXml), 0);
  fm_page_setup_patch_t clear = FM_STRUCT_INIT(fm_page_setup_patch_t);
  clear.orientation_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &clear), 0) << fm_get_last_error_message();
  fm_page_setup_t read = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.has_orientation, 0);
  EXPECT_EQ(read.orientation, FM_PAGE_ORIENTATION_PORTRAIT);
  EXPECT_EQ(read.has_paper_size, 1);
  EXPECT_EQ(read.paper_size, 9U);
  EXPECT_EQ(read.scale_pct, 80U);
}

TEST(FormulonCApiPrintPatch, PageSetupRejectsBadOpsAndWritesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, kPageSetupXml), 0);
  const auto with_valid_scale = [] {
    fm_page_setup_patch_t p = FM_STRUCT_INIT(fm_page_setup_patch_t);
    p.scale_pct_op = FM_PATCH_OP_SET;
    p.scale_pct = 50U;
    return p;
  };
  std::vector<fm_page_setup_patch_t> rejected;
  fm_page_setup_patch_t p = with_valid_scale();
  p.paper_size_op = FM_PATCH_OP_CLEAR;
  rejected.push_back(p);
  p = with_valid_scale();
  p.fit_to_page_op = FM_PATCH_OP_CLEAR;
  rejected.push_back(p);
  p = with_valid_scale();
  p.orientation_op = kBadOp;
  rejected.push_back(p);
  p = with_valid_scale();
  p.fit_to_width_op = kBadOp;
  rejected.push_back(p);
  for (const fm_page_setup_patch_t& patch : rejected) {
    EXPECT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &patch), kInvalidArgument);
    EXPECT_EQ(Fragment(fm_sheet_get_page_setup_xml, wb.handle), kPageSetupXml);
  }
  EXPECT_EQ(fm_sheet_update_page_setup(wb.handle, 9, &p), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiPrintPatch, PageMarginsAndPrintOptionsCannotBeCleared) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_page_margins_xml(wb.handle, 0, kMarginsXml), 0);
  ASSERT_EQ(fm_sheet_set_print_options_xml(wb.handle, 0, kOptionsXml), 0);
  for (const fm_patch_op_t op : {FM_PATCH_OP_CLEAR, kBadOp}) {
    fm_page_margins_patch_t margins = FM_STRUCT_INIT(fm_page_margins_patch_t);
    margins.left_in_op = FM_PATCH_OP_SET;
    margins.left_in = 2.0;
    margins.footer_in_op = op;
    EXPECT_EQ(fm_sheet_update_page_margins(wb.handle, 0, &margins), kInvalidArgument) << op;
    fm_print_options_patch_t options = FM_STRUCT_INIT(fm_print_options_patch_t);
    options.headings_op = FM_PATCH_OP_SET;
    options.headings = 1;
    options.grid_lines_op = op;
    EXPECT_EQ(fm_sheet_update_print_options(wb.handle, 0, &options), kInvalidArgument) << op;
  }
  // A valid SET beside a rejected value is not written either.
  fm_page_margins_patch_t mixed = FM_STRUCT_INIT(fm_page_margins_patch_t);
  mixed.left_in_op = FM_PATCH_OP_SET;
  mixed.left_in = 2.0;
  mixed.bottom_in_op = FM_PATCH_OP_SET;
  mixed.bottom_in = -1.0;
  EXPECT_EQ(fm_sheet_update_page_margins(wb.handle, 0, &mixed), kInvalidArgument);
  EXPECT_EQ(Fragment(fm_sheet_get_page_margins_xml, wb.handle), kMarginsXml);
  EXPECT_EQ(Fragment(fm_sheet_get_print_options_xml, wb.handle), kOptionsXml);
}

TEST(FormulonCApiPrintPatch, HeaderFooterClearRemovesASection) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_header_footer_xml(wb.handle, 0, kHeaderFooterXml), 0);
  fm_header_footer_patch_t clear = FM_STRUCT_INIT(fm_header_footer_patch_t);
  clear.odd_header_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &clear), 0) << fm_get_last_error_message();
  fm_header_footer_t read = FM_STRUCT_INIT(fm_header_footer_t);
  ASSERT_EQ(fm_sheet_get_header_footer(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.odd_header, nullptr);
  ASSERT_NE(read.odd_footer, nullptr);
  EXPECT_STREQ(read.odd_footer, "F");
  EXPECT_EQ(read.different_first, 1);
}

TEST(FormulonCApiPrintPatch, HeaderFooterRejectsBadPatchesAndWritesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_header_footer_xml(wb.handle, 0, kHeaderFooterXml), 0);
  struct Case {
    const char* what;
    fm_status_t expected;
    fm_header_footer_patch_t patch;
  };
  const auto base = [] {
    fm_header_footer_patch_t p = FM_STRUCT_INIT(fm_header_footer_patch_t);
    p.even_header_op = FM_PATCH_OP_SET;
    p.even_header = "E";
    return p;
  };
  std::vector<Case> cases;
  Case c{"clear on a flag", kInvalidArgument, base()};
  c.patch.different_first_op = FM_PATCH_OP_CLEAR;
  cases.push_back(c);
  c = Case{"op outside 0..2", kInvalidArgument, base()};
  c.patch.first_footer_op = kBadOp;
  cases.push_back(c);
  c = Case{"set with a NULL section", FM_STATUS_INVALID_ARGUMENT, base()};
  c.patch.odd_footer_op = FM_PATCH_OP_SET;
  c.patch.odd_footer = nullptr;
  cases.push_back(c);
  c = Case{"set malformed UTF-8", FM_STATUS_BINDING_UTF8ENCODING_ERROR, base()};
  c.patch.odd_footer_op = FM_PATCH_OP_SET;
  c.patch.odd_footer = "\xff";
  cases.push_back(c);
  for (const Case& entry : cases) {
    EXPECT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &entry.patch), entry.expected) << entry.what;
    EXPECT_EQ(Fragment(fm_sheet_get_header_footer_xml, wb.handle), kHeaderFooterXml) << entry.what;
  }
}

TEST(FormulonCApiPrintSettings, GetPrintAreaFillsAtMostCapacityAndReportsTheFullCount) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t areas[] = {{0, 0, 1, 1}, {4, 4, 5, 5}, {9, 0, 9, 3}};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, areas, 3), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_print_area(wb.handle, 0, nullptr, 0, &count), 0);
  EXPECT_EQ(count, 3U);
  fm_cell_rect_t out[2] = {};
  const fm_cell_rect_t sentinel{77, 77, 77, 77};
  out[1] = sentinel;
  count = 0;
  ASSERT_EQ(fm_sheet_get_print_area(wb.handle, 0, out, 1, &count), 0);
  EXPECT_EQ(count, 3U);
  EXPECT_EQ(out[0].first_row, 0U);
  EXPECT_EQ(out[0].last_col, 1U);
  // Nothing past the capacity is written.
  EXPECT_EQ(out[1].first_row, 77U);
}

TEST(FormulonCApiPageBreakManualDefault, ZeroInitBreakIsManual) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_row_break_t row = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_set_row_break(wb.handle, 0, 7, &row), 0);
  fm_column_break_t col = FM_STRUCT_INIT(fm_column_break_t);
  ASSERT_EQ(fm_sheet_set_column_break(wb.handle, 0, 3, &col), 0);
  fm_row_break_t row_back = FM_STRUCT_INIT(fm_row_break_t);
  fm_column_break_t col_back = FM_STRUCT_INIT(fm_column_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &row_back), 0);
  ASSERT_EQ(fm_sheet_get_column_break_at(wb.handle, 0, 0, &col_back), 0);
  EXPECT_EQ(row_back.manual, FM_TRISTATE_TRUE);
  EXPECT_EQ(col_back.manual, FM_TRISTATE_TRUE);
}

TEST(FormulonCApiPageBreakManualDefault, ReadReturnsTrueOrFalse) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 4, FM_TRISTATE_FALSE), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 9, FM_TRISTATE_DEFAULT), 0);
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.manual, FM_TRISTATE_FALSE);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.manual, FM_TRISTATE_TRUE);
}

/* -------------------------------------------------------------------------- */
/* A workbook-scope built-in hidden by a sheet-local one                      */
/* -------------------------------------------------------------------------- */

namespace {

std::string ReadFormula(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  int32_t present = 0;
  const char* text = nullptr;
  EXPECT_EQ(fm_sheet_get_formula(wb, 0, row, col, &present, &text), 0);
  return text != nullptr ? text : "";
}

double ReadNumber(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  fm_value_t value{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, row, col, &value), 0);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  return value.u.number;
}

}  // namespace

TEST(FormulonCApiPrintSettings, SelfBookQualifierReachesAWorkbookPrintAreaHiddenByTheSheetsOwn) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().set_defined_names(
      {formulon::DefinedName("_xlnm.Print_Area", "Sheet1!$A$1:$B$3", -1, false, ""),
       formulon::DefinedName("_xlnm.Print_Area", "Sheet1!$A$1:$B$5", 0, false, "")});
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=ROWS([0]!Print_Area)"), 0) << fm_get_last_error_message();
  EXPECT_EQ(wb.handle->workbook().sheet(0).cell_at(0, 0)->formula_text, "=ROWS([0]!_xlnm.Print_Area)");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(ReadNumber(wb.handle, 0, 0), 3.0);
  const std::string shown = ReadFormula(wb.handle, 0, 0);
  EXPECT_EQ(shown, "=ROWS([0]!Print_Area)");

  // The shown text enters again as the same formula.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, shown.c_str()), 0) << fm_get_last_error_message();
  EXPECT_EQ(wb.handle->workbook().sheet(0).cell_at(0, 1)->formula_text, "=ROWS([0]!_xlnm.Print_Area)");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(ReadNumber(wb.handle, 0, 1), 3.0);
  EXPECT_EQ(ReadFormula(wb.handle, 0, 1), shown);

  // The bare name still reads the sheet's own.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 2, "=ROWS(Print_Area)"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(ReadNumber(wb.handle, 0, 2), 5.0);
}

TEST(FormulonCApiPrintSettings, SelfBookQualifierFallsToTheLowestSheetsLocalPrintArea) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().set_defined_names(
      {formulon::DefinedName("_xlnm.Print_Area", "Sheet1!$A$1:$B$5", 0, false, "")});
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=ROWS([0]!print_area)"), 0) << fm_get_last_error_message();
  EXPECT_EQ(wb.handle->workbook().sheet(0).cell_at(0, 0)->formula_text, "=ROWS([0]!_xlnm.Print_Area)");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(ReadNumber(wb.handle, 0, 0), 5.0);
}
