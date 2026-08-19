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
#include "print/pagination.h"
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

/// Wraps `<sheetFormatPr>` / `<cols>` attributes around a two-cell sheet
/// whose used range spans 50 rows, so a change in the fallback track size
/// is visible as a change in the page count.
std::string SheetXmlWithGeometry(std::string_view format_attrs, std::string_view cols, std::string_view trailing = {}) {
  return std::string(
             "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
             "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
             "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
             "  <sheetFormatPr ")
      .append(format_attrs)
      .append("/>\n  ")
      .append(cols)
      .append(
          "<sheetData>\n"
          "    <row r=\"1\"><c r=\"A1\"><v>1</v></c></row>\n"
          "    <row r=\"200\"><c r=\"D200\"><v>2</v></c></row>\n"
          "  </sheetData>\n  ")
      .append(trailing)
      .append("</worksheet>\n");
}

TEST(SheetPrintGeometry, SheetFormatPrMeasurementsOutsideTheLexicalSpaceAreTreatedAsAbsent) {
  // `strtod` reads these as +inf, a quiet NaN, 16, 0 and +inf. Any of them
  // reaching the model becomes the fallback size for every un-overridden
  // track on the sheet.
  for (const char* spelling : {"INF", "NaN", "0x10", "abc", "1e999", "-5"}) {
    const std::string format_attrs = std::string("baseColWidth=\"")
                                         .append(spelling)
                                         .append("\" defaultColWidth=\"")
                                         .append(spelling)
                                         .append("\" defaultRowHeight=\"")
                                         .append(spelling)
                                         .append("\"");
    const std::vector<std::uint8_t> source = BuildPackageWithSheetXml(SheetXmlWithGeometry(format_attrs, ""));
    auto result_or = io::read_ooxml(SpanOf(source));
    ASSERT_TRUE(static_cast<bool>(result_or)) << spelling << ": " << result_or.error().message;
    const SheetFormatDefaults& defaults = result_or.value().workbook.sheet(0).format_defaults();
    EXPECT_FALSE(defaults.has_default_col_width) << spelling;
    EXPECT_FALSE(defaults.has_default_row_height) << spelling;
    EXPECT_DOUBLE_EQ(defaults.base_col_width, 8.0) << spelling;
  }
}

TEST(SheetPrintGeometry, ColumnWidthOutsideTheLexicalSpaceIsTreatedAsAbsent) {
  for (const char* spelling : {"INF", "NaN", "0x10", "abc", "1e999", "-5"}) {
    const std::string cols = std::string("<cols><col min=\"1\" max=\"4\" width=\"")
                                 .append(spelling)
                                 .append("\" customWidth=\"1\"/></cols>\n  ");
    const std::vector<std::uint8_t> source =
        BuildPackageWithSheetXml(SheetXmlWithGeometry("defaultRowHeight=\"15\"", cols));
    auto result_or = io::read_ooxml(SpanOf(source));
    ASSERT_TRUE(static_cast<bool>(result_or)) << spelling << ": " << result_or.error().message;
    // The span carried nothing but the unusable width, so it contributes
    // no column layout rather than one sized by a non-measurement.
    EXPECT_TRUE(result_or.value().workbook.sheet(0).layout().columns.empty()) << spelling;
  }
}

TEST(SheetPrintGeometry, HostileTrackSizesPaginateLikeTheSchemaDefaults) {
  // The observable consequence of admitting one: an infinite default row
  // height breaks a page before every row, and a NaN one makes every
  // accumulation comparison false so the whole sheet reports a single
  // page. Both must now match the well-formed sheet exactly.
  // The baseline omits both measurements rather than spelling the schema
  // defaults: "treated as absent" is the claim, so the comparison has to be
  // against a genuinely absent attribute. Spelling a number here made the
  // test assert that the engine's fallback equals that literal, which is a
  // different (and geometry-model-dependent) claim.
  const std::vector<std::uint8_t> baseline_bytes = BuildPackageWithSheetXml(SheetXmlWithGeometry("", ""));
  auto baseline_or = io::read_ooxml(SpanOf(baseline_bytes));
  ASSERT_TRUE(static_cast<bool>(baseline_or)) << baseline_or.error().message;
  auto baseline_pages_or = print::paginate(baseline_or.value().workbook, 0U);
  ASSERT_TRUE(static_cast<bool>(baseline_pages_or)) << baseline_pages_or.error().message;
  const print::PaginationResult& baseline = baseline_pages_or.value();
  ASSERT_GT(baseline.page_count, 1U) << "the baseline must span pages for this comparison to bite";

  for (const char* spelling : {"INF", "NaN", "1e999", "abc"}) {
    const std::string format_attrs = std::string("defaultRowHeight=\"").append(spelling).append("\"");
    const std::vector<std::uint8_t> source = BuildPackageWithSheetXml(SheetXmlWithGeometry(format_attrs, ""));
    auto result_or = io::read_ooxml(SpanOf(source));
    ASSERT_TRUE(static_cast<bool>(result_or)) << spelling << ": " << result_or.error().message;
    auto pages_or = print::paginate(result_or.value().workbook, 0U);
    ASSERT_TRUE(static_cast<bool>(pages_or)) << spelling << ": " << pages_or.error().message;
    EXPECT_EQ(pages_or.value().page_count, baseline.page_count) << spelling;
    EXPECT_EQ(pages_or.value().h_breaks, baseline.h_breaks) << spelling;
    EXPECT_EQ(pages_or.value().v_breaks, baseline.v_breaks) << spelling;
  }
}

TEST(SheetPrintGeometry, PageMarginsOutsideTheLexicalSpaceKeepTheDefaults) {
  // The margins are subtracted from the paper to get the printable body,
  // so an infinity or a NaN collapses that body and breaks a page before
  // every track, while a negative one inflates it past the paper. The raw
  // `<pageMargins>` string is round-tripped verbatim either way, so the
  // file still says what it said.
  const std::vector<std::uint8_t> baseline_bytes = BuildPackageWithSheetXml(SheetXmlWithGeometry(
      "defaultRowHeight=\"15\" defaultColWidth=\"8.43\"", "",
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\n  "));
  auto baseline_or = io::read_ooxml(SpanOf(baseline_bytes));
  ASSERT_TRUE(static_cast<bool>(baseline_or)) << baseline_or.error().message;
  auto baseline_pages_or = print::paginate(baseline_or.value().workbook, 0U);
  ASSERT_TRUE(static_cast<bool>(baseline_pages_or)) << baseline_pages_or.error().message;
  const print::PaginationResult& baseline = baseline_pages_or.value();
  ASSERT_GT(baseline.page_count, 1U);

  for (const char* spelling : {"INF", "NaN", "1e999", "-99", "abc"}) {
    const std::string margins = std::string("<pageMargins left=\"")
                                    .append(spelling)
                                    .append("\" right=\"")
                                    .append(spelling)
                                    .append("\" top=\"")
                                    .append(spelling)
                                    .append("\" bottom=\"")
                                    .append(spelling)
                                    .append("\" header=\"0.3\" footer=\"0.3\"/>\n  ");
    const std::vector<std::uint8_t> source =
        BuildPackageWithSheetXml(SheetXmlWithGeometry("defaultRowHeight=\"15\" defaultColWidth=\"8.43\"", "", margins));
    auto result_or = io::read_ooxml(SpanOf(source));
    ASSERT_TRUE(static_cast<bool>(result_or)) << spelling << ": " << result_or.error().message;
    const PageMargins& margins_read = result_or.value().workbook.sheet(0).print_settings().page_margins;
    EXPECT_DOUBLE_EQ(margins_read.left, 0.7) << spelling;
    EXPECT_DOUBLE_EQ(margins_read.top, 0.75) << spelling;
    auto pages_or = print::paginate(result_or.value().workbook, 0U);
    ASSERT_TRUE(static_cast<bool>(pages_or)) << spelling << ": " << pages_or.error().message;
    EXPECT_EQ(pages_or.value().page_count, baseline.page_count) << spelling;
    EXPECT_EQ(pages_or.value().h_breaks, baseline.h_breaks) << spelling;
  }
}

TEST(SheetPrintGeometry, SheetFormatPrDefaultsSurviveWriteReadCycle) {
  Workbook src = Workbook::create();
  SheetFormatDefaults& defaults = src.sheet(0).mutable_format_defaults();
  defaults.base_col_width = 10.0;
  defaults.default_col_width = 12.5;
  defaults.default_row_height = 18.75;
  defaults.has_default_col_width = true;
  defaults.has_default_row_height = true;

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const SheetFormatDefaults& loaded = result_or.value().workbook.sheet(0).format_defaults();
  EXPECT_DOUBLE_EQ(loaded.base_col_width, 10.0);
  EXPECT_TRUE(loaded.has_default_col_width);
  EXPECT_DOUBLE_EQ(loaded.default_col_width, 12.5);
  EXPECT_TRUE(loaded.has_default_row_height);
  EXPECT_DOUBLE_EQ(loaded.default_row_height, 18.75);
}

// ---------------------------------------------------------------------------
// Manual page breaks
// ---------------------------------------------------------------------------

TEST(SheetPrintGeometry, ManualBreaksReadFromWorksheetXml) {
  // OOXML's `id` is the 0-based index the break precedes, which is what
  // the model stores -- the reader passes it through. Excel 365 writes
  // `id="20"` for a break placed before row 21, so an `id="10"` here means
  // the page ends after ten rows and the next starts at 0-based row 10.
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
  EXPECT_EQ(print.manual_row_breaks[0].id, 10U);
  EXPECT_EQ(print.manual_row_breaks[0].max, kLastColumnIndex);
  EXPECT_TRUE(print.manual_row_breaks[0].manual);
  EXPECT_EQ(print.manual_row_breaks[1].id, 20U);
  ASSERT_EQ(print.manual_col_breaks.size(), 1U);
  EXPECT_EQ(print.manual_col_breaks[0].id, 5U);
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
