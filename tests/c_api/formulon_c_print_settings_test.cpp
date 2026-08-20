//
// Stable C ABI tests for the print-authoring surface: raw print-settings
// XML, the `fitToPage` helper, print area / titles, manual page breaks and
// the typed patch setters.
//
// The driver is C++ for gtest convenience; every mutation and observation
// stays on the pure-C surface. Two properties get repeated attention
// because they are what the surface is for: a fragment the engine stores
// must be one Excel can open, and a setting written through any of the
// three surfaces must be visible to `fm_workbook_paginate` immediately.

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "workbook.h"

namespace {

// `fm_page_break_t` and the three patch structs cross the ABI by pointer,
// and the bindings marshal them field by field from a hand-written layout
// table. Pin the layout so a field insertion breaks the build here rather
// than silently mis-decoding in the Python ctypes projection.
static_assert(sizeof(fm_page_break_t) == 16U, "fm_page_break_t ABI layout changed");
static_assert(offsetof(fm_page_break_t, manual) == 12U, "fm_page_break_t.manual offset changed");
static_assert(sizeof(fm_page_setup_t) == 48U, "fm_page_setup_t ABI layout changed");
static_assert(offsetof(fm_page_setup_t, fit_to_page) == 44U, "fm_page_setup_t.fit_to_page offset changed");
static_assert(sizeof(fm_page_margins_t) == 96U, "fm_page_margins_t ABI layout changed");
static_assert(offsetof(fm_page_margins_t, left) == 8U, "fm_page_margins_t.left offset changed");
static_assert(sizeof(fm_print_options_t) == 32U, "fm_print_options_t ABI layout changed");

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

}  // namespace

/* -------------------------------------------------------------------------- */
/* Raw XML - validation                                                       */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, FreshSheetReportsEmptyFragments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* fragment = "<pageSetup paperSize=\"9\" orientation=\"portrait\" scale=\"85\"/>";
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, fragment), 0);
  EXPECT_EQ(GetPageSetupXml(wb.handle), fragment);
}

TEST(FormulonCApiPrintSettings, EmptyStringRemovesElementAndResetsStructuredView) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup orientation=\"landscape\" scale=\"50\"/>"), 0);
  fm_page_setup_t before{};
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &before), 0);
  ASSERT_EQ(before.orientation, FM_ORIENTATION_LANDSCAPE);
  ASSERT_EQ(before.scale, 50U);

  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, ""), 0);
  EXPECT_EQ(GetPageSetupXml(wb.handle), "");
  fm_page_setup_t after{};
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &after), 0);
  EXPECT_EQ(after.orientation, FM_ORIENTATION_DEFAULT);
  EXPECT_EQ(after.scale, 100U);
  EXPECT_EQ(after.orientation_engaged, 0);
  EXPECT_EQ(after.scale_engaged, 0);
}

TEST(FormulonCApiPrintSettings, RejectsMalformedFragments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(
      fm_sheet_set_sheet_pr_xml(wb.handle, 0, "<sheetPr codeName=\"Sheet1\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"),
      0);
  ASSERT_EQ(fm_sheet_set_fit_to_page(wb.handle, 0, 1), 0);

  const std::string xml = GetSheetPrXml(wb.handle);
  // A VBA binding and a tab colour are not print settings, and toggling
  // fit-to-page must not be how a caller loses them.
  EXPECT_NE(xml.find("codeName=\"Sheet1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("<tabColor rgb=\"FFFF0000\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("fitToPage=\"true\""), std::string::npos) << xml;

  fm_page_setup_t setup{};
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &setup), 0);
  EXPECT_EQ(setup.fit_to_page, 1);
  EXPECT_EQ(setup.fit_to_page_engaged, 1);
}

TEST(FormulonCApiPrintSettings, ClearingFitToPageKeepsTheElements) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(
      fm_sheet_set_sheet_pr_xml(wb.handle, 0, "<sheetPr><pageSetUpPr fitToPage=\"1\" autoPageBreaks=\"0\"/></sheetPr>"),
      0);
  ASSERT_EQ(fm_sheet_set_fit_to_page(wb.handle, 0, 0), 0);

  const std::string xml = GetSheetPrXml(wb.handle);
  EXPECT_EQ(xml.find("fitToPage"), std::string::npos) << xml;
  EXPECT_NE(xml.find("autoPageBreaks=\"0\""), std::string::npos) << xml;

  fm_page_setup_t setup{};
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &setup), 0);
  EXPECT_EQ(setup.fit_to_page, 0);
  EXPECT_EQ(setup.fit_to_page_engaged, 0);
}

TEST(FormulonCApiPrintSettings, FitToPageCreatesSheetPrWhenAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_fit_to_page(wb.handle, 0, 1), 0);
  EXPECT_EQ(GetSheetPrXml(wb.handle), "<sheetPr><pageSetUpPr fitToPage=\"true\"/></sheetPr>");
}

/* -------------------------------------------------------------------------- */
/* Print area and titles                                                      */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, SetPrintAreaWritesQualifiedAbsoluteDefinedName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, "A1:F8"), 0);

  const auto& names = wb.handle->workbook().defined_names();
  ASSERT_EQ(names.size(), 1U);
  EXPECT_EQ(names[0].name, "_xlnm.Print_Area");
  EXPECT_EQ(names[0].formula, "Sheet1!$A$1:$F$8");
  EXPECT_EQ(names[0].local_sheet_id, 0);

  const char* ranges = nullptr;
  ASSERT_EQ(fm_sheet_get_print_area(wb.handle, 0, &ranges), 0);
  EXPECT_STREQ(ranges, "A1:F8");
}

TEST(FormulonCApiPrintSettings, SetPrintAreaQuotesSheetNamesThatNeedIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_rename_sheet(wb.handle, 0, "It's 2026"), 0);
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, "A1:B2"), 0);
  // A space forces quoting and the apostrophe has to be doubled, or Excel
  // reads the formula as ending at the apostrophe.
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "'It''s 2026'!$A$1:$B$2");
}

TEST(FormulonCApiPrintSettings, SetPrintAreaKeepsMultiAreaOrderAndWholeAxisShape) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, "D5:E20, A1:B10"), 0);
  // Excel prints the areas in the stated order, so the engine must not
  // sort them.
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$D$5:$E$20,Sheet1!$A$1:$B$10");

  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, "A:D"), 0);
  // A whole-column area keeps its authored shape in the file; only the
  // getter reports the resolver's expanded rectangle.
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$A:$D");
  const char* ranges = nullptr;
  ASSERT_EQ(fm_sheet_get_print_area(wb.handle, 0, &ranges), 0);
  EXPECT_STREQ(ranges, "A1:D1048576");
}

TEST(FormulonCApiPrintSettings, SetPrintAreaRejectsMalformedAreaAndLeavesNameUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, "A1:F8"), 0);
  EXPECT_EQ(fm_sheet_set_print_area(wb.handle, 0, "A1:F8,not-a-range"), kPrintInvalidArea);
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$A$1:$F$8");
}

TEST(FormulonCApiPrintSettings, EmptyPrintAreaRemovesTheDefinedName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, "A1:F8"), 0);
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, ""), 0);
  EXPECT_TRUE(wb.handle->workbook().defined_names().empty());

  const char* ranges = nullptr;
  ASSERT_EQ(fm_sheet_get_print_area(wb.handle, 0, &ranges), 0);
  EXPECT_STREQ(ranges, "");
}

TEST(FormulonCApiPrintSettings, GetPrintAreaSurfacesAMalformedDefinedName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // A file can arrive with a print area this engine cannot parse. Reporting
  // it as "no print area" would silently print the whole used range.
  ASSERT_TRUE(
      static_cast<bool>(wb.handle->workbook().set_defined_name_scoped("_xlnm.Print_Area", "Sheet1!$A$1:$$$", 0)));
  const char* ranges = nullptr;
  EXPECT_EQ(fm_sheet_get_print_area(wb.handle, 0, &ranges), kPrintInvalidArea);
}

TEST(FormulonCApiPrintSettings, PrintTitlesRoundTripRowsBeforeColumns) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, "1:2", "A:A"), 0);
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$1:$2,Sheet1!$A:$A");

  const char* rows = nullptr;
  const char* cols = nullptr;
  ASSERT_EQ(fm_sheet_get_print_titles(wb.handle, 0, &rows, &cols), 0);
  // Both pointers come from one scratch refresh and must stay valid
  // together; reading the first after the second is written is the case
  // that would dangle if the getter cleared between pushes.
  EXPECT_STREQ(rows, "1:2");
  EXPECT_STREQ(cols, "A:A");
}

TEST(FormulonCApiPrintSettings, PrintTitlesAcceptOneAxisAndClearOnBothEmpty) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, "1:1", nullptr), 0);
  EXPECT_EQ(wb.handle->workbook().defined_names()[0].formula, "Sheet1!$1:$1");

  const char* rows = nullptr;
  const char* cols = nullptr;
  ASSERT_EQ(fm_sheet_get_print_titles(wb.handle, 0, &rows, &cols), 0);
  EXPECT_STREQ(rows, "1:1");
  EXPECT_STREQ(cols, "");

  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, "", ""), 0);
  EXPECT_TRUE(wb.handle->workbook().defined_names().empty());
}

TEST(FormulonCApiPrintSettings, PrintTitlesRejectMalformedSpans) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // A cell range is not a title span: Excel only repeats whole rows or
  // whole columns.
  EXPECT_EQ(fm_sheet_set_print_titles(wb.handle, 0, "A1:B2", ""), kPrintInvalidArea);
  EXPECT_EQ(fm_sheet_set_print_titles(wb.handle, 0, "", "1:2"), kPrintInvalidArea);
  EXPECT_EQ(fm_sheet_set_print_titles(wb.handle, 0, "0:2", ""), kPrintInvalidArea);
  EXPECT_TRUE(wb.handle->workbook().defined_names().empty());
}

/* -------------------------------------------------------------------------- */
/* Manual page breaks                                                         */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, BreaksUpsertStaySortedAndReplaceInPlace) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, 40, 1), 0);
  ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, 10, 1), 0);
  ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, 25, 0), 0);
  ASSERT_EQ(fm_sheet_row_break_count(wb.handle, 0), 3U);

  fm_page_break_t brk{};
  ASSERT_EQ(fm_sheet_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.id, 10U);
  EXPECT_EQ(brk.min, 0U);
  EXPECT_EQ(brk.max, 16383U);
  EXPECT_EQ(brk.manual, 1);
  ASSERT_EQ(fm_sheet_row_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.id, 25U);
  EXPECT_EQ(brk.manual, 0);
  ASSERT_EQ(fm_sheet_row_break_at(wb.handle, 0, 2, &brk), 0);
  EXPECT_EQ(brk.id, 40U);

  // Re-adding an existing index replaces it rather than duplicating.
  ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, 25, 1), 0);
  EXPECT_EQ(fm_sheet_row_break_count(wb.handle, 0), 3U);
  ASSERT_EQ(fm_sheet_row_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.id, 25U);
  EXPECT_EQ(brk.manual, 1);
}

TEST(FormulonCApiPrintSettings, LoadedBreaksAreNormalisedBeforeTheAbiObservesThem) {
  // A third-party writer is free to spell `<brk>` in authoring order and
  // to repeat an index. The mutation API binary-searches these vectors
  // and the enumeration API documents ascending order, so a document
  // that arrives unsorted would make `fm_sheet_remove_col_break` match
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
  ASSERT_EQ(fm_sheet_col_break_count(wb.handle, 0), 3U);
  fm_page_break_t brk{};
  std::uint32_t previous = 0;
  for (std::size_t i = 0; i < 3U; ++i) {
    ASSERT_EQ(fm_sheet_col_break_at(wb.handle, 0, i, &brk), 0);
    if (i > 0) {
      EXPECT_GT(brk.id, previous);
    }
    previous = brk.id;
  }
  ASSERT_EQ(fm_sheet_col_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.id, 12U);

  // The mutators now find what the enumeration reports.
  EXPECT_EQ(fm_sheet_remove_col_break(wb.handle, 0, 12), 0);
  EXPECT_EQ(fm_sheet_col_break_count(wb.handle, 0), 2U);
  ASSERT_EQ(fm_sheet_add_col_break(wb.handle, 0, 12, 1), 0);
  EXPECT_EQ(fm_sheet_col_break_count(wb.handle, 0), 3U);
  ASSERT_EQ(fm_sheet_col_break_at(wb.handle, 0, 1, &brk), 0);
  EXPECT_EQ(brk.id, 12U);
}

TEST(FormulonCApiPrintSettings, LoadedBreakCountStaysWithinTheAuthoringCap) {
  // Reading a file must not leave a sheet holding more breaks than
  // `fm_sheet_add_*_break` accepts, or adding one break to a loaded
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
  EXPECT_EQ(fm_sheet_row_break_count(wb.handle, 0), formulon::kMaxManualBreaksPerAxis);

  // An index already present is an in-place replacement, so it is
  // accepted at the cap; a new one is what the cap refuses.
  fm_page_break_t brk{};
  ASSERT_EQ(fm_sheet_row_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(fm_sheet_add_row_break(wb.handle, 0, brk.id, 0), 0);
  EXPECT_EQ(fm_sheet_add_row_break(wb.handle, 0, formulon::Sheet::kMaxRows - 1U, 1), kPreconditionFailed);
}

TEST(FormulonCApiPrintSettings, BreakCountSignalsARejectedSheetThroughTheDiagnostic) {
  // `size_t` leaves no room for a status, so a rejected count and an empty
  // axis both read as zero. A caller enumerating breaks as
  // count-then-`_at` never reaches the `_at` call that would surface the
  // rejection, so the diagnostic is the only thing that separates the two
  // and it has to be reliable in both directions.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const std::size_t sheet_count = fm_workbook_sheet_count(wb.handle);

  // Empty axis: zero, and no diagnostic.
  EXPECT_EQ(fm_sheet_row_break_count(wb.handle, 0), 0U);
  EXPECT_STREQ(fm_last_error_message(), "");
  EXPECT_EQ(fm_sheet_col_break_count(wb.handle, 0), 0U);
  EXPECT_STREQ(fm_last_error_message(), "");

  // Sheet past the end: zero, and a diagnostic naming the entry point.
  EXPECT_EQ(fm_sheet_row_break_count(wb.handle, sheet_count), 0U);
  EXPECT_STRNE(fm_last_error_message(), "");
  EXPECT_EQ(fm_sheet_col_break_count(wb.handle, sheet_count), 0U);
  EXPECT_STRNE(fm_last_error_message(), "");

  // NULL handle takes the same path.
  EXPECT_EQ(fm_sheet_row_break_count(nullptr, 0), 0U);
  EXPECT_STRNE(fm_last_error_message(), "");

  // A populated axis clears the previous rejection rather than inheriting
  // it, so the empty message stays a trustworthy success signal.
  ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, 10, 1), 0);
  EXPECT_EQ(fm_sheet_row_break_count(wb.handle, 0), 1U);
  EXPECT_STREQ(fm_last_error_message(), "");

  // The indexed reader rejects the same sheet outright.
  fm_page_break_t brk{};
  EXPECT_EQ(fm_sheet_row_break_at(wb.handle, sheet_count, 0, &brk), kInvalidArgument);
  EXPECT_EQ(fm_sheet_col_break_at(wb.handle, sheet_count, 0, &brk), kInvalidArgument);
}

TEST(FormulonCApiPrintSettings, ColumnBreaksSpanEveryRow) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_add_col_break(wb.handle, 0, 5, 1), 0);
  fm_page_break_t brk{};
  ASSERT_EQ(fm_sheet_col_break_at(wb.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.id, 5U);
  EXPECT_EQ(brk.max, 1048575U);
}

TEST(FormulonCApiPrintSettings, RemovingAnAbsentBreakSucceeds) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // A caller clearing a break it is not sure exists is not making a
  // mistake, so absence is not an error.
  EXPECT_EQ(fm_sheet_remove_row_break(wb.handle, 0, 7), 0);
  ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, 7, 1), 0);
  EXPECT_EQ(fm_sheet_remove_row_break(wb.handle, 0, 7), 0);
  EXPECT_EQ(fm_sheet_row_break_count(wb.handle, 0), 0U);
}

TEST(FormulonCApiPrintSettings, ClearBreaksEmptiesBothAxes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, 3, 1), 0);
  ASSERT_EQ(fm_sheet_add_col_break(wb.handle, 0, 4, 1), 0);
  ASSERT_EQ(fm_sheet_clear_breaks(wb.handle, 0), 0);
  EXPECT_EQ(fm_sheet_row_break_count(wb.handle, 0), 0U);
  EXPECT_EQ(fm_sheet_col_break_count(wb.handle, 0), 0U);
}

TEST(FormulonCApiPrintSettings, BreaksRejectOutOfGridAndOverflow) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_add_row_break(wb.handle, 0, formulon::Sheet::kMaxRows, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_col_break(wb.handle, 0, formulon::Sheet::kMaxCols, 1), kInvalidArgument);

  for (std::uint32_t row = 0; row < 1026U; ++row) {
    ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, row, 1), 0) << "row " << row;
  }
  // Excel's own ceiling; going past it produces a file Excel truncates.
  EXPECT_EQ(fm_sheet_add_row_break(wb.handle, 0, 2000, 1), kPreconditionFailed);
  // Replacing an existing break is still allowed at the ceiling.
  EXPECT_EQ(fm_sheet_add_row_break(wb.handle, 0, 0, 0), 0);
}

TEST(FormulonCApiPrintSettings, BreakAtRejectsOutOfRangeIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_page_break_t brk{};
  EXPECT_EQ(fm_sheet_row_break_at(wb.handle, 0, 0, &brk), kInvalidArgument);
  EXPECT_EQ(fm_sheet_col_break_at(wb.handle, 0, 0, &brk), kInvalidArgument);
}

/* -------------------------------------------------------------------------- */
/* Typed patch setters                                                        */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, PageSetupPatchPreservesUnmodelledAttributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(
      fm_sheet_set_page_setup_xml(
          wb.handle, 0, "<pageSetup paperSize=\"9\" orientation=\"portrait\" horizontalDpi=\"600\" copies=\"3\"/>"),
      0);

  fm_page_setup_t patch{};
  patch.orientation_engaged = 1;
  patch.orientation = FM_ORIENTATION_LANDSCAPE;
  ASSERT_EQ(fm_sheet_set_page_setup(wb.handle, 0, &patch), 0);

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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_page_setup_t patch{};
  patch.scale_engaged = 1;
  patch.scale = 9;
  // A mis-stated print scale lands on paper, so it is rejected rather than
  // clamped the way the on-screen zoom is.
  EXPECT_EQ(fm_sheet_set_page_setup(wb.handle, 0, &patch), kInvalidArgument);
  patch.scale = 401;
  EXPECT_EQ(fm_sheet_set_page_setup(wb.handle, 0, &patch), kInvalidArgument);
  EXPECT_EQ(GetPageSetupXml(wb.handle), "");

  patch.scale = 400;
  EXPECT_EQ(fm_sheet_set_page_setup(wb.handle, 0, &patch), 0);
}

TEST(FormulonCApiPrintSettings, PageSetupPatchRoutesFitToPageToSheetPr) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_page_setup_t patch{};
  patch.fit_to_page_engaged = 1;
  patch.fit_to_page = 1;
  patch.fit_to_width_engaged = 1;
  patch.fit_to_width = 1;
  patch.fit_to_height_engaged = 1;
  patch.fit_to_height = 0;
  ASSERT_EQ(fm_sheet_set_page_setup(wb.handle, 0, &patch), 0);

  // `fitToPage` selects the mode and lives in `<sheetPr>`; the page counts
  // that state the target live in `<pageSetup>`. Both halves are needed.
  EXPECT_NE(GetSheetPrXml(wb.handle).find("fitToPage=\"true\""), std::string::npos);
  const std::string setup_xml = GetPageSetupXml(wb.handle);
  EXPECT_NE(setup_xml.find("fitToWidth=\"1\""), std::string::npos) << setup_xml;
  EXPECT_NE(setup_xml.find("fitToHeight=\"0\""), std::string::npos) << setup_xml;
}

TEST(FormulonCApiPrintSettings, PageMarginsPatchTouchesOnlyEngagedFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_page_margins_t patch{};
  patch.left_engaged = 1;
  patch.left = 0.25;
  patch.top_engaged = 1;
  patch.top = 1.5;
  ASSERT_EQ(fm_sheet_set_page_margins(wb.handle, 0, &patch), 0);

  fm_page_margins_t read{};
  ASSERT_EQ(fm_sheet_get_page_margins(wb.handle, 0, &read), 0);
  EXPECT_DOUBLE_EQ(read.left, 0.25);
  EXPECT_DOUBLE_EQ(read.top, 1.5);
  EXPECT_EQ(read.left_engaged, 1);
  EXPECT_EQ(read.top_engaged, 1);
  // An unstated margin reports the OOXML default value with its presence
  // flag clear, so the caller can tell "0.7 because the file says so" from
  // "0.7 because nothing says otherwise".
  EXPECT_DOUBLE_EQ(read.right, 0.7);
  EXPECT_EQ(read.right_engaged, 0);
}

TEST(FormulonCApiPrintSettings, PageMarginsPatchRejectsNonFiniteAndNegative) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_page_margins_t patch{};
  patch.bottom_engaged = 1;
  patch.bottom = -1.0;
  // The paginator subtracts margins from the paper; a negative one
  // inflates the printable body past the sheet.
  EXPECT_EQ(fm_sheet_set_page_margins(wb.handle, 0, &patch), kInvalidArgument);
  patch.bottom = 1.0 / 0.0;
  EXPECT_EQ(fm_sheet_set_page_margins(wb.handle, 0, &patch), kInvalidArgument);
}

TEST(FormulonCApiPrintSettings, PrintOptionsAndHeaderFooterPatch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_print_options_t options{};
  options.grid_lines_engaged = 1;
  options.grid_lines = 1;
  options.horizontal_centered_engaged = 1;
  options.horizontal_centered = 1;
  ASSERT_EQ(fm_sheet_set_print_options(wb.handle, 0, &options), 0);
  const char* options_xml = nullptr;
  ASSERT_EQ(fm_sheet_get_print_options_xml(wb.handle, 0, &options_xml), 0);
  EXPECT_STREQ(options_xml, "<printOptions gridLines=\"true\" horizontalCentered=\"true\"/>");

  fm_header_footer_t hf{};
  hf.odd_header = "&C\xE5\xB8\xB3\xE7\xA5\xA8";  // "&C帳票"
  hf.odd_footer = "&R&P / &N";
  ASSERT_EQ(fm_sheet_set_header_footer(wb.handle, 0, &hf), 0);
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // Written last-to-first on purpose: ECMA-376 fixes the child order, and
  // appending in call order would put `<oddHeader>` after `<firstFooter>`.
  fm_header_footer_t late{};
  late.first_footer = "F";
  ASSERT_EQ(fm_sheet_set_header_footer(wb.handle, 0, &late), 0);
  fm_header_footer_t early{};
  early.odd_header = "H";
  early.even_footer = "E";
  ASSERT_EQ(fm_sheet_set_header_footer(wb.handle, 0, &early), 0);

  const char* xml = nullptr;
  ASSERT_EQ(fm_sheet_get_header_footer_xml(wb.handle, 0, &xml), 0);
  EXPECT_STREQ(xml,
               "<headerFooter><oddHeader>H</oddHeader><evenFooter>E</evenFooter>"
               "<firstFooter>F</firstFooter></headerFooter>");
}

TEST(FormulonCApiPrintSettings, HeaderFooterNullLeavesSectionAndEmptyClearsIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_header_footer_t first{};
  first.odd_header = "Keep";
  first.odd_footer = "Drop";
  ASSERT_EQ(fm_sheet_set_header_footer(wb.handle, 0, &first), 0);

  fm_header_footer_t second{};
  second.odd_header = nullptr;  // untouched
  second.odd_footer = "";       // cleared
  ASSERT_EQ(fm_sheet_set_header_footer(wb.handle, 0, &second), 0);

  const char* xml = nullptr;
  ASSERT_EQ(fm_sheet_get_header_footer_xml(wb.handle, 0, &xml), 0);
  EXPECT_STREQ(xml, "<headerFooter><oddHeader>Keep</oddHeader></headerFooter>");
}

/* -------------------------------------------------------------------------- */
/* Pagination linkage                                                         */
/* -------------------------------------------------------------------------- */

namespace {

/// Paginates sheet 0 and returns its page count.
std::uint32_t PageCount(fm_workbook_t* wb) {
  PaginationGuard pagination;
  EXPECT_EQ(fm_workbook_paginate(wb, 0, &pagination.handle), 0);
  return fm_pagination_page_count(pagination.handle);
}

/// Fills `A1:T200` so the sheet needs more than one page either way.
void FillGrid(fm_workbook_t* wb) {
  for (std::uint32_t row = 0; row < 200U; ++row) {
    for (std::uint32_t col = 0; col < 20U; ++col) {
      ASSERT_EQ(fm_workbook_set_number(wb, 0, row, col, 1.0), 0);
    }
  }
}

}  // namespace

TEST(FormulonCApiPrintSettings, OrientationChangeIsVisibleToPaginateImmediately) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);

  fm_page_setup_t portrait{};
  portrait.orientation_engaged = 1;
  portrait.orientation = FM_ORIENTATION_PORTRAIT;
  ASSERT_EQ(fm_sheet_set_page_setup(wb.handle, 0, &portrait), 0);
  const std::uint32_t portrait_pages = PageCount(wb.handle);

  fm_page_setup_t landscape{};
  landscape.orientation_engaged = 1;
  landscape.orientation = FM_ORIENTATION_LANDSCAPE;
  ASSERT_EQ(fm_sheet_set_page_setup(wb.handle, 0, &landscape), 0);
  const std::uint32_t landscape_pages = PageCount(wb.handle);

  // No save/load in between: the setter re-derives the structured view the
  // paginator reads, which is the whole point of routing every mutation
  // through the reader's parser.
  EXPECT_NE(portrait_pages, landscape_pages);
}

TEST(FormulonCApiPrintSettings, RawXmlSetterAlsoDrivesPagination) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const std::uint32_t before = PageCount(wb.handle);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup scale=\"25\"/>"), 0);
  EXPECT_LT(PageCount(wb.handle), before);
}

TEST(FormulonCApiPrintSettings, ManualRowBreakAddsAPage) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 9, 0, 1.0), 0);
  const std::uint32_t before = PageCount(wb.handle);
  ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, 5, 1), 0);
  EXPECT_GT(PageCount(wb.handle), before);
}

TEST(FormulonCApiPrintSettings, PrintAreaNarrowsWhatPaginates) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const std::uint32_t full = PageCount(wb.handle);
  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, "A1:B2"), 0);
  EXPECT_LT(PageCount(wb.handle), full);
  EXPECT_EQ(PageCount(wb.handle), 1U);
}

/* -------------------------------------------------------------------------- */
/* Argument handling                                                          */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiPrintSettings, NullArgumentsReportBindingNullPointer) {
  constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_get_page_setup_xml(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_set_print_area(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_get_page_setup(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_set_page_setup(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_set_page_margins(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_set_print_options(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_set_header_footer(wb.handle, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_row_break_at(wb.handle, 0, 0, nullptr), kNullPointer);
}

TEST(FormulonCApiPrintSettings, OutOfRangeSheetIndexIsRejectedEverywhere) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* xml = nullptr;
  EXPECT_EQ(fm_sheet_get_page_setup_xml(wb.handle, 9, &xml), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 9, "<pageSetup/>"), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_fit_to_page(wb.handle, 9, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_print_area(wb.handle, 9, "A1:B2"), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_row_break(wb.handle, 9, 1, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_row_break_count(wb.handle, 9), 0U);
}
