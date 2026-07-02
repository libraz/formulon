// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Round-trip integration tests for the print-geometry surface added to
// the sheet model: `<sheetFormatPr>` defaults, manual `<rowBreaks>` /
// `<colBreaks>`, and the structured `PageSetup` / `PageMargins` views
// parsed alongside the raw XML passthrough strings.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

/// Highest 0-based column index in an Excel 365 worksheet (column XFD).
/// Used as the `max` span bound on a manual row break.
constexpr std::uint32_t kLastColumnIndex = 16383U;

/// Highest 0-based row index in an Excel 365 worksheet. Used as the
/// `max` span bound on a manual column break.
constexpr std::uint32_t kLastRowIndex = 1048575U;

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

std::vector<std::uint8_t> SaveOrDie(const Workbook& workbook) {
  auto save_or = workbook.save();
  EXPECT_TRUE(static_cast<bool>(save_or)) << "save() failed: " << save_or.error().message;
  return save_or.value();
}

struct PartFile {
  const char* path;
  std::string_view body;
};

/// Materialises `parts` into a heap-allocated zip archive byte vector via
/// miniz. Mirrors the helper used by the existing roundtrip suite.
std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& part : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, part.path, part.body.data(), part.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << "miniz add failed for " << part.path;
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(archive_size);
  if (archive_ptr != nullptr && archive_size > 0) {
    std::memcpy(out.data(), archive_ptr, archive_size);
  }
  mz_free(archive_ptr);
  return out;
}

/// Builds a minimal one-sheet xlsx package whose `sheet1.xml` body is the
/// caller-supplied `<worksheet>` fragment. Lets each test exercise a
/// hand-authored worksheet part without a printerSettings rel.
std::vector<std::uint8_t> BuildPackageWithSheetXml(std::string_view sheet_xml) {
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n"
      "  <Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
      "</Types>\n";
  const std::string_view package_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view workbook_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>\n"
      "</workbook>\n";
  const std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "</Relationships>\n";
  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
  });
}

// ---------------------------------------------------------------------------
// <sheetFormatPr> defaults
// ---------------------------------------------------------------------------

TEST(SheetPrintGeometry, SheetFormatPrDefaultsParsedWhenPresent) {
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheetFormatPr baseColWidth=\"10\" defaultColWidth=\"12.5\" defaultRowHeight=\"18.75\"/>\n"
      "  <sheetData/>\n"
      "</worksheet>\n";
  const std::vector<std::uint8_t> source = BuildPackageWithSheetXml(sheet_xml);
  auto result_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const SheetFormatDefaults& defaults = result_or.value().workbook.sheet(0).format_defaults();
  EXPECT_TRUE(defaults.has_default_col_width);
  EXPECT_DOUBLE_EQ(defaults.default_col_width, 12.5);
  EXPECT_TRUE(defaults.has_default_row_height);
  EXPECT_DOUBLE_EQ(defaults.default_row_height, 18.75);
  EXPECT_DOUBLE_EQ(defaults.base_col_width, 10.0);
}

TEST(SheetPrintGeometry, SheetFormatPrAbsentLeavesStructDefaults) {
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheetData/>\n"
      "</worksheet>\n";
  const std::vector<std::uint8_t> source = BuildPackageWithSheetXml(sheet_xml);
  auto result_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const SheetFormatDefaults& defaults = result_or.value().workbook.sheet(0).format_defaults();
  EXPECT_FALSE(defaults.has_default_col_width);
  EXPECT_FALSE(defaults.has_default_row_height);
  EXPECT_DOUBLE_EQ(defaults.default_col_width, 0.0);
  EXPECT_DOUBLE_EQ(defaults.default_row_height, 0.0);
  // baseColWidth carries the OOXML spec default of 8 when absent.
  EXPECT_DOUBLE_EQ(defaults.base_col_width, 8.0);
}

TEST(SheetPrintGeometry, SheetFormatPrPartialElementSetsOnlyPresentFlags) {
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheetFormatPr defaultRowHeight=\"15\"/>\n"
      "  <sheetData/>\n"
      "</worksheet>\n";
  const std::vector<std::uint8_t> source = BuildPackageWithSheetXml(sheet_xml);
  auto result_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const SheetFormatDefaults& defaults = result_or.value().workbook.sheet(0).format_defaults();
  EXPECT_FALSE(defaults.has_default_col_width);
  EXPECT_TRUE(defaults.has_default_row_height);
  EXPECT_DOUBLE_EQ(defaults.default_row_height, 15.0);
  EXPECT_DOUBLE_EQ(defaults.base_col_width, 8.0);
}

// ---------------------------------------------------------------------------
// Manual page breaks
// ---------------------------------------------------------------------------

TEST(SheetPrintGeometry, ManualBreaksReadFromWorksheetXml) {
  // OOXML stores break ids 1-based; the reader normalises to 0-based.
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheetData/>\n"
      "  <rowBreaks count=\"2\" manualBreakCount=\"2\">\n"
      "    <brk id=\"10\" min=\"0\" max=\"16383\" man=\"1\"/>\n"
      "    <brk id=\"20\" min=\"0\" max=\"16383\" man=\"1\"/>\n"
      "  </rowBreaks>\n"
      "  <colBreaks count=\"1\" manualBreakCount=\"1\">\n"
      "    <brk id=\"5\" min=\"0\" max=\"1048575\" man=\"1\"/>\n"
      "  </colBreaks>\n"
      "</worksheet>\n";
  const std::vector<std::uint8_t> source = BuildPackageWithSheetXml(sheet_xml);
  auto result_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const SheetPrintSettings& print = result_or.value().workbook.sheet(0).print_settings();
  ASSERT_EQ(print.manual_row_breaks.size(), 2U);
  EXPECT_EQ(print.manual_row_breaks[0].id, 9U);
  EXPECT_EQ(print.manual_row_breaks[0].max, kLastColumnIndex);
  EXPECT_TRUE(print.manual_row_breaks[0].manual);
  EXPECT_EQ(print.manual_row_breaks[1].id, 19U);
  ASSERT_EQ(print.manual_col_breaks.size(), 1U);
  EXPECT_EQ(print.manual_col_breaks[0].id, 4U);
  EXPECT_EQ(print.manual_col_breaks[0].max, kLastRowIndex);
}

TEST(SheetPrintGeometry, ManualBreaksSurviveWriteReadCycle) {
  // Arbitrary 0-based break positions chosen so the assertions read
  // unambiguously; the two column breaks differ so the sort below is
  // exercised.
  constexpr std::uint32_t kRowBreakIndex = 12U;
  constexpr std::uint32_t kFirstColBreakIndex = 3U;
  constexpr std::uint32_t kSecondColBreakIndex = 7U;

  Workbook src = Workbook::create();
  SheetPrintSettings& print = src.sheet(0).mutable_print_settings();
  // `man` defaults to false (ECMA-376); user-inserted breaks are manual,
  // so mark them explicitly so they re-emit with `man="1"` and read back
  // as manual.
  ManualBreak row_break;
  row_break.id = kRowBreakIndex;
  row_break.min = 0U;
  row_break.max = kLastColumnIndex;
  row_break.manual = true;
  print.manual_row_breaks.push_back(row_break);
  ManualBreak col_break_a;
  col_break_a.id = kFirstColBreakIndex;
  col_break_a.min = 0U;
  col_break_a.max = kLastRowIndex;
  col_break_a.manual = true;
  print.manual_col_breaks.push_back(col_break_a);
  ManualBreak col_break_b;
  col_break_b.id = kSecondColBreakIndex;
  col_break_b.min = 0U;
  col_break_b.max = kLastRowIndex;
  col_break_b.manual = true;
  print.manual_col_breaks.push_back(col_break_b);

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const SheetPrintSettings& loaded = result_or.value().workbook.sheet(0).print_settings();
  ASSERT_EQ(loaded.manual_row_breaks.size(), 1U);
  EXPECT_EQ(loaded.manual_row_breaks[0].id, kRowBreakIndex);
  EXPECT_EQ(loaded.manual_row_breaks[0].max, kLastColumnIndex);
  EXPECT_TRUE(loaded.manual_row_breaks[0].manual);
  ASSERT_EQ(loaded.manual_col_breaks.size(), 2U);
  std::vector<ManualBreak> sorted = loaded.manual_col_breaks;
  std::sort(sorted.begin(), sorted.end(), [](const ManualBreak& a, const ManualBreak& b) { return a.id < b.id; });
  EXPECT_EQ(sorted[0].id, kFirstColBreakIndex);
  EXPECT_EQ(sorted[1].id, kSecondColBreakIndex);
  EXPECT_EQ(sorted[1].max, kLastRowIndex);
}

TEST(SheetPrintGeometry, NoManualBreaksEmitsNoBreakElements) {
  Workbook src = Workbook::create();
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const SheetPrintSettings& loaded = result_or.value().workbook.sheet(0).print_settings();
  EXPECT_TRUE(loaded.manual_row_breaks.empty());
  EXPECT_TRUE(loaded.manual_col_breaks.empty());
}

// ---------------------------------------------------------------------------
// Structured PageSetup / PageMargins (parsed alongside the raw strings)
// ---------------------------------------------------------------------------

TEST(SheetPrintGeometry, StructuredPageSetupAndMarginsParsed) {
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>\n"
      "  <sheetData/>\n"
      "  <pageMargins left=\"1.1\" right=\"1.2\" top=\"0.9\" bottom=\"0.8\" header=\"0.5\" footer=\"0.45\"/>\n"
      "  <pageSetup paperSize=\"1\" orientation=\"landscape\" scale=\"85\" fitToWidth=\"2\" fitToHeight=\"3\"/>\n"
      "</worksheet>\n";
  const std::vector<std::uint8_t> source = BuildPackageWithSheetXml(sheet_xml);
  auto result_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const SheetPrintSettings& print = result_or.value().workbook.sheet(0).print_settings();

  // Raw passthrough strings remain populated and untouched.
  EXPECT_NE(print.page_setup_xml.find("orientation=\"landscape\""), std::string::npos) << print.page_setup_xml;
  EXPECT_NE(print.page_margins_xml.find("left=\"1.1\""), std::string::npos) << print.page_margins_xml;

  // Structured PageSetup view.
  EXPECT_EQ(print.page_setup.orientation, Orientation::kLandscape);
  EXPECT_EQ(print.page_setup.paper_size, 1U);
  EXPECT_EQ(print.page_setup.scale, 85U);
  EXPECT_EQ(print.page_setup.fit_to_width, 2U);
  EXPECT_EQ(print.page_setup.fit_to_height, 3U);
  EXPECT_TRUE(print.page_setup.fit_to_page);

  // Structured PageMargins view.
  EXPECT_DOUBLE_EQ(print.page_margins.left, 1.1);
  EXPECT_DOUBLE_EQ(print.page_margins.right, 1.2);
  EXPECT_DOUBLE_EQ(print.page_margins.top, 0.9);
  EXPECT_DOUBLE_EQ(print.page_margins.bottom, 0.8);
  EXPECT_DOUBLE_EQ(print.page_margins.header, 0.5);
  EXPECT_DOUBLE_EQ(print.page_margins.footer, 0.45);
}

TEST(SheetPrintGeometry, StructuredPageSetupKeepsDefaultsWhenAttributesAbsent) {
  // A bare `<pageSetup/>` with no attributes; the structured view keeps
  // its OOXML-default values, and `fit_to_page` stays false without a
  // `<pageSetUpPr>`.
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheetData/>\n"
      "  <pageSetup/>\n"
      "</worksheet>\n";
  const std::vector<std::uint8_t> source = BuildPackageWithSheetXml(sheet_xml);
  auto result_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const SheetPrintSettings& print = result_or.value().workbook.sheet(0).print_settings();
  EXPECT_EQ(print.page_setup.orientation, Orientation::kDefault);
  EXPECT_EQ(print.page_setup.paper_size, 9U);
  EXPECT_EQ(print.page_setup.scale, 100U);
  EXPECT_EQ(print.page_setup.fit_to_width, 1U);
  EXPECT_EQ(print.page_setup.fit_to_height, 1U);
  EXPECT_FALSE(print.page_setup.fit_to_page);
  // PageMargins with no `<pageMargins>` element retains the spec defaults.
  EXPECT_DOUBLE_EQ(print.page_margins.left, 0.7);
  EXPECT_DOUBLE_EQ(print.page_margins.top, 0.75);
  EXPECT_DOUBLE_EQ(print.page_margins.header, 0.3);
}

}  // namespace
}  // namespace formulon
