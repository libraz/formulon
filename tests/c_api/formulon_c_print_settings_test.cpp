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

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
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
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);
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

TEST(FormulonCApiPrintSettings, EmptyStringRemovesElementAndResetsStructuredView) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup orientation=\"landscape\" scale=\"50\"/>"), 0);
  fm_page_setup_t before = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &before), 0);
  ASSERT_EQ(before.orientation, FM_PAGE_ORIENTATION_LANDSCAPE);
  ASSERT_EQ(before.scale_pct, 50U);

  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, ""), 0);
  EXPECT_EQ(GetPageSetupXml(wb.handle), "");
  fm_page_setup_t after = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &after), 0);
  EXPECT_EQ(after.orientation, FM_PAGE_ORIENTATION_PORTRAIT);
  EXPECT_EQ(after.scale_pct, 100U);
  EXPECT_EQ(after.orientation_engaged, 0);
  EXPECT_EQ(after.scale_pct_engaged, 0);
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
  EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, oversized.c_str()), kPreconditionFailed);

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
  EXPECT_NE(xml.find("fitToPage=\"true\""), std::string::npos) << xml;

  fm_page_setup_t setup = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &setup), 0);
  EXPECT_EQ(setup.fit_to_page, 1);
  EXPECT_EQ(setup.fit_to_page_engaged, 1);
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
  EXPECT_EQ(setup.fit_to_page_engaged, 0);
}

TEST(FormulonCApiPrintSettings, FitToPageCreatesSheetPrWhenAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetFitToPage(wb.handle, 0, 1), 0);
  EXPECT_EQ(GetSheetPrXml(wb.handle), "<sheetPr><pageSetUpPr fitToPage=\"true\"/></sheetPr>");
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

TEST(FormulonCApiPrintSettings, SetPrintAreaOverTheRangeCapReportsInvalidArgument) {
  // The per-call range cap is shared with the conditional-formatting and
  // data-validation entry points, and one failure condition carries one
  // status across all of them.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t area{0, 0, 7, 5};
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, &area, 1), 0);

  const std::vector<fm_cell_rect_t> areas(formulon::c_api::parts::kMaxRangesPerCApiCall + 1U, fm_cell_rect_t{});
  EXPECT_EQ(fm_sheet_set_print_area(wb.handle, 0, areas.data(), areas.size()), kInvalidArgument);
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$A$1:$F$8");
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
  ASSERT_EQ(SetRowBreak(wb.handle, 40, 1), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 10, 1), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 25, 0), 0);
  ASSERT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 3U);

  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.row, 10U);
  EXPECT_EQ(brk.first, 0U);
  EXPECT_EQ(brk.last, 16383U);
  EXPECT_EQ(brk.manual, 1);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.row, 25U);
  EXPECT_EQ(brk.manual, 0);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 2, &brk), 0);
  EXPECT_EQ(brk.row, 40U);

  // Re-adding an existing index replaces it rather than duplicating.
  ASSERT_EQ(SetRowBreak(wb.handle, 25, 1), 0);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 3U);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.row, 25U);
  EXPECT_EQ(brk.manual, 1);
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
  ASSERT_EQ(fm_workbook_load(bytes.value().data(), bytes.value().size(), &wb.handle), 0);

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
  ASSERT_EQ(SetColumnBreak(wb.handle, 12, 1), 0);
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
  ASSERT_EQ(fm_workbook_load(bytes.value().data(), bytes.value().size(), &wb.handle), 0);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), formulon::kMaxManualBreaksPerAxis);

  // An index already present is an in-place replacement, so it is
  // accepted at the cap; a new one is what the cap refuses.
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(SetRowBreak(wb.handle, brk.row, 0), 0);
  EXPECT_EQ(SetRowBreak(wb.handle, formulon::Sheet::kMaxRows - 1U, 1), kPreconditionFailed);
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
  ASSERT_EQ(SetRowBreak(wb.handle, 10, 1), 0);
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
  ASSERT_EQ(SetColumnBreak(wb.handle, 5, 1), 0);
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
  ASSERT_EQ(SetRowBreak(wb.handle, 7, 1), 0);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 0, 7, nullptr), 0);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 0U);
}

TEST(FormulonCApiPrintSettings, ClearBreaksEmptiesBothAxes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 3, 1), 0);
  ASSERT_EQ(SetColumnBreak(wb.handle, 4, 1), 0);
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
  ASSERT_EQ(SetRowBreak(wb.handle, formulon::Sheet::kMaxRows - 1U, 1), 0);
  ASSERT_EQ(SetColumnBreak(wb.handle, formulon::Sheet::kMaxCols - 1U, 1), 0);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 0, formulon::Sheet::kMaxRows - 1U, nullptr), 0);
  EXPECT_EQ(fm_sheet_set_column_break(wb.handle, 0, formulon::Sheet::kMaxCols - 1U, nullptr), 0);
  EXPECT_EQ(fm_sheet_set_row_break(wb.handle, 0, formulon::Sheet::kMaxRows - 1U, nullptr), 0);
}

TEST(FormulonCApiPrintSettings, BreaksRejectOutOfGridAndOverflow) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(SetRowBreak(wb.handle, formulon::Sheet::kMaxRows, 1), kInvalidArgument);
  EXPECT_EQ(SetColumnBreak(wb.handle, formulon::Sheet::kMaxCols, 1), kInvalidArgument);

  for (std::uint32_t row = 0; row < 1026U; ++row) {
    ASSERT_EQ(SetRowBreak(wb.handle, row, 1), 0) << "row " << row;
  }
  // Excel's own ceiling; going past it produces a file Excel truncates.
  EXPECT_EQ(SetRowBreak(wb.handle, 2000, 1), kPreconditionFailed);
  // Replacing an existing break is still allowed at the ceiling.
  EXPECT_EQ(SetRowBreak(wb.handle, 0, 0), 0);
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
  EXPECT_NE(GetSheetPrXml(wb.handle).find("fitToPage=\"true\""), std::string::npos);
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
  EXPECT_EQ(read.left_in_engaged, 1);
  EXPECT_EQ(read.top_in_engaged, 1);
  // An unstated margin reports the OOXML default value with its presence
  // flag clear, so the caller can tell "0.7 because the file says so" from
  // "0.7 because nothing says otherwise".
  EXPECT_DOUBLE_EQ(read.right_in, 0.7);
  EXPECT_EQ(read.right_in_engaged, 0);
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
  EXPECT_STREQ(options_xml, "<printOptions gridLines=\"true\" horizontalCentered=\"true\"/>");

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
  EXPECT_EQ(setup.fit_to_page_engaged, 0);
  EXPECT_EQ(GetSheetPrXml(wb.handle), "<sheetPr><pageSetUpPr/></sheetPr>");
}

TEST(FormulonCApiPrintSettings, PrintOptionsGetterReportsStatedAndAbsentAttributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_print_options_t fresh = FM_STRUCT_INIT(fm_print_options_t);
  fresh.grid_lines = 7;
  ASSERT_EQ(fm_sheet_get_print_options(wb.handle, 0, &fresh), 0);
  EXPECT_EQ(fresh.grid_lines, 0);
  EXPECT_EQ(fresh.grid_lines_engaged, 0);
  EXPECT_EQ(fresh.headings_engaged, 0);
  EXPECT_EQ(fresh.horizontal_centered_engaged, 0);
  EXPECT_EQ(fresh.vertical_centered_engaged, 0);

  fm_print_options_patch_t patch = FM_STRUCT_INIT(fm_print_options_patch_t);
  patch.grid_lines_op = FM_PATCH_OP_SET;
  patch.grid_lines = 1;
  patch.headings_op = FM_PATCH_OP_SET;
  patch.headings = 0;
  ASSERT_EQ(fm_sheet_update_print_options(wb.handle, 0, &patch), 0);

  fm_print_options_t read = FM_STRUCT_INIT(fm_print_options_t);
  ASSERT_EQ(fm_sheet_get_print_options(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.grid_lines, 1);
  EXPECT_EQ(read.grid_lines_engaged, 1);
  // A stated `false` is distinct from an absent attribute.
  EXPECT_EQ(read.headings, 0);
  EXPECT_EQ(read.headings_engaged, 1);
  EXPECT_EQ(read.horizontal_centered, 0);
  EXPECT_EQ(read.horizontal_centered_engaged, 0);
  EXPECT_EQ(read.vertical_centered_engaged, 0);

  ASSERT_EQ(fm_sheet_set_print_options_xml(wb.handle, 0, "<printOptions verticalCentered=\"1\"/>"), 0);
  ASSERT_EQ(fm_sheet_get_print_options(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.vertical_centered, 1);
  EXPECT_EQ(read.vertical_centered_engaged, 1);
  EXPECT_EQ(read.grid_lines_engaged, 0);
}

TEST(FormulonCApiPrintSettings, HeaderFooterGetterReadsDecodedSectionsAndNullForAbsentOnes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_header_footer_t fresh = FM_STRUCT_INIT(fm_header_footer_t);
  ASSERT_EQ(fm_sheet_get_header_footer(wb.handle, 0, &fresh), 0);
  EXPECT_EQ(fresh.odd_header, nullptr);
  EXPECT_EQ(fresh.first_footer, nullptr);
  EXPECT_EQ(fresh.different_odd_even, 0);
  EXPECT_EQ(fresh.different_odd_even_engaged, 0);
  // CT_HeaderFooter defaults these two to true.
  EXPECT_EQ(fresh.scale_with_doc, 1);
  EXPECT_EQ(fresh.scale_with_doc_engaged, 0);
  EXPECT_EQ(fresh.align_with_margins, 1);
  EXPECT_EQ(fresh.align_with_margins_engaged, 0);

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
  EXPECT_EQ(read.different_odd_even_engaged, 1);
  EXPECT_EQ(read.scale_with_doc, 0);
  EXPECT_EQ(read.scale_with_doc_engaged, 1);
  EXPECT_EQ(read.different_first_engaged, 0);
  EXPECT_EQ(read.align_with_margins, 1);
  EXPECT_EQ(read.align_with_margins_engaged, 0);
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
  ASSERT_EQ(SetRowBreak(wb.handle, 5, 1), 0);
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
  EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, nullptr), kNullPointer);
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
  EXPECT_EQ(read.scale_pct_engaged, 1);
  EXPECT_EQ(read.orientation, FM_PAGE_ORIENTATION_LANDSCAPE);
  EXPECT_EQ(read.orientation_engaged, 1);
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
  EXPECT_EQ(read.scale_pct_engaged, 0);
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
  EXPECT_EQ(read.header_in_engaged, 1);
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
  EXPECT_EQ(read.top_in_engaged, 1);
  EXPECT_DOUBLE_EQ(read.footer_in, 0.4);
  EXPECT_EQ(read.footer_in_engaged, 1);
  EXPECT_EQ(read.left_in_engaged, 0);
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
  EXPECT_EQ(read.horizontal_centered_engaged, 0);
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
  ASSERT_EQ(SetRowBreak(wb.handle, 12, 1, 2, 9), 0);
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.row, 12U);
  EXPECT_EQ(brk.first, 2U);
  EXPECT_EQ(brk.last, 9U);
  EXPECT_EQ(brk.manual, 1);

  // Upsert replaces the span and the flag in place.
  ASSERT_EQ(SetRowBreak(wb.handle, 12, 0, 4, 5), 0);
  ASSERT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 1U);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.first, 4U);
  EXPECT_EQ(brk.last, 5U);
  EXPECT_EQ(brk.manual, 0);
}

TEST(FormulonCApiPrintSettings, RowBreakWithZeroSpanCoversTheWholeSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 30, 1, 0, 0), 0);
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
  brk.manual = 1;
  ASSERT_EQ(fm_sheet_set_row_break(wb.handle, 0, 7, &brk), 0);
  fm_row_break_t read = FM_STRUCT_INIT(fm_row_break_t);
  ASSERT_EQ(fm_sheet_get_row_break_at(wb.handle, 0, 0, &read), 0);
  EXPECT_EQ(read.row, 7U);
  EXPECT_EQ(count_of(fm_sheet_get_row_break_count, wb.handle, 0), 1U);
}

TEST(FormulonCApiPrintSettings, RowBreakNullRemovesTheBreakAndIsANoOpWhenAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 7, 1), 0);
  ASSERT_EQ(SetRowBreak(wb.handle, 8, 1), 0);
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
  ASSERT_EQ(SetColumnBreak(wb.handle, 6, 1, 10, 20), 0);
  fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
  ASSERT_EQ(fm_sheet_get_column_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.col, 6U);
  EXPECT_EQ(brk.first, 10U);
  EXPECT_EQ(brk.last, 20U);
  EXPECT_EQ(brk.manual, 1);

  ASSERT_EQ(SetColumnBreak(wb.handle, 6, 0, 0, 0), 0);
  ASSERT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 1U);
  ASSERT_EQ(fm_sheet_get_column_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.first, 0U);
  EXPECT_EQ(brk.last, formulon::Sheet::kMaxRows - 1U);
  EXPECT_EQ(brk.manual, 0);
}

TEST(FormulonCApiPrintSettings, ColumnBreakNullRemovesTheBreakAndIsANoOpWhenAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_column_break(wb.handle, 0, 3, nullptr), 0);
  ASSERT_EQ(SetColumnBreak(wb.handle, 3, 1), 0);
  ASSERT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 1U);
  EXPECT_EQ(fm_sheet_set_column_break(wb.handle, 0, 3, nullptr), 0);
  EXPECT_EQ(count_of(fm_sheet_get_column_break_count, wb.handle, 0), 0U);
}

TEST(FormulonCApiPrintSettings, BreakRecordsRejectInvertedSpansAndBadFlags) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(SetRowBreak(wb.handle, 5, 1, 9, 2), kInvalidArgument);
  EXPECT_EQ(SetRowBreak(wb.handle, 5, 2), kInvalidArgument);
  EXPECT_EQ(SetRowBreak(wb.handle, 5, 1, 0, formulon::Sheet::kMaxCols), kInvalidArgument);
  EXPECT_EQ(SetColumnBreak(wb.handle, 5, 1, 9, 2), kInvalidArgument);
  EXPECT_EQ(SetColumnBreak(wb.handle, 5, -1), kInvalidArgument);
  EXPECT_EQ(SetColumnBreak(wb.handle, 5, 1, 0, formulon::Sheet::kMaxRows), kInvalidArgument);
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
const char* const kOptionsXml = "<printOptions gridLines=\"true\"/>";
const char* const kHeaderFooterXml =
    "<headerFooter differentFirst=\"true\"><oddHeader>H</oddHeader><oddFooter>F</oddFooter></headerFooter>";

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
  EXPECT_EQ(read.orientation_engaged, 0);
  EXPECT_EQ(read.orientation, FM_PAGE_ORIENTATION_PORTRAIT);
  EXPECT_EQ(read.paper_size_engaged, 1);
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
  c = Case{"set with a NULL section", kInvalidArgument, base()};
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
