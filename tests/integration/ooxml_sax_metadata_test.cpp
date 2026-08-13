//
// Integration test: sheets large enough to route through the streaming
// SAX read path (>= kSaxThresholdBytes) must still recover the non-cell
// worksheet metadata (view, merges, hyperlinks, protection, autoFilter,
// print settings) and resolve shared-formula groups. Before this, the
// SAX path emitted only cells and silently dropped everything else.
//
// The sheet is padded past 256 KiB so the reader picks the SAX branch on
// native builds; the assertions then confirm the recovered metadata.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_types.h"  // complete ConditionalFormat for vector::size
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/sheet_reader.h"  // kSaxThresholdBytes
#include "miniz.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

struct PartFile {
  const char* path;
  std::string_view body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.path, p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << "miniz add failed for " << p.path;
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

// Builds a worksheet whose <sheetData> is padded well past the SAX
// threshold, carrying a shared-formula group up front and a full set of
// non-cell metadata elements after the data.
std::string BuildLargeSheet() {
  std::string body;
  body.reserve(400U * 1024U);
  body.append(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "<sheetViews><sheetView showGridLines=\"0\" tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><v>10</v></c>"
      "<c r=\"B1\"><f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*2</f><v>20</v></c></row>\n"
      "<row r=\"2\"><c r=\"A2\"><v>11</v></c><c r=\"B2\"><f t=\"shared\" si=\"0\"/><v>22</v></c></row>\n"
      "<row r=\"3\"><c r=\"A3\"><v>12</v></c><c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>24</v></c></row>\n");
  // Pad with plain numeric rows until the payload comfortably exceeds the
  // SAX threshold so the reader chooses the streaming branch.
  std::uint32_t r = 4;
  while (body.size() < static_cast<std::size_t>(io::kSaxThresholdBytes) + 64U * 1024U) {
    body.append("<row r=\"");
    body.append(std::to_string(r));
    body.append("\"><c r=\"A");
    body.append(std::to_string(r));
    body.append("\"><v>");
    body.append(std::to_string(r));
    body.append("</v></c></row>\n");
    ++r;
  }
  body.append("</sheetData>\n");
  // Non-cell metadata siblings, in ECMA-376 order.
  body.append("<sheetProtection sheet=\"1\" formatCells=\"0\"/>\n");
  body.append("<autoFilter ref=\"A1:C1\"/>\n");
  body.append("<mergeCells count=\"1\"><mergeCell ref=\"D1:E1\"/></mergeCells>\n");
  body.append("<hyperlinks><hyperlink ref=\"A1:B2\" location=\"Sheet1!A1\"/></hyperlinks>\n");
  body.append("<printOptions horizontalCentered=\"1\" gridLines=\"1\"/>\n");
  body.append("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\n");
  body.append("<pageSetup orientation=\"landscape\"/>\n");
  // UTF-8 Japanese header string must survive verbatim.
  body.append(
      "<headerFooter><oddHeader>&amp;C\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E\xE3\x83\x98\xE3\x83\x83\xE3\x83\x80"
      "</oddHeader></headerFooter>\n");
  body.append("</worksheet>\n");
  return body;
}

std::vector<std::uint8_t> BuildPackage(const std::string& sheet_body) {
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
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
      {"xl/worksheets/sheet1.xml", std::string_view(sheet_body)},
  });
}

TEST(OoxmlSaxMetadata, LargeSheetRecoversMetadataAndSharedFormulas) {
  // SAX is compiled out on WASM (threshold == SIZE_MAX); this native test
  // binary always has it enabled, but guard defensively.
  if (io::kSaxThresholdBytes == static_cast<std::size_t>(-1)) {
    GTEST_SKIP() << "SAX path disabled in this build";
  }
  const std::string sheet_body = BuildLargeSheet();
  ASSERT_GE(sheet_body.size(), static_cast<std::size_t>(io::kSaxThresholdBytes))
      << "sheet must exceed the SAX threshold to exercise the streaming path";

  auto load_or = io::read_ooxml(SpanOf(BuildPackage(sheet_body)));
  ASSERT_TRUE(static_cast<bool>(load_or)) << load_or.error().message;
  const Workbook& wb = load_or.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Sheet& sheet = wb.sheet(0);

  // Cells streamed in (a padding cell is present).
  const Cell* a10 = sheet.cell_at(9U, 0U);
  ASSERT_NE(a10, nullptr);
  ASSERT_TRUE(a10->cached_value.is_number());
  EXPECT_DOUBLE_EQ(a10->cached_value.as_number(), 10.0);

  // Shared-formula follower resolved on the SAX path.
  const Cell* b2 = sheet.cell_at(1U, 1U);
  ASSERT_NE(b2, nullptr);
  EXPECT_EQ(b2->formula_text, "=A2*2");

  // sheetView display attributes recovered.
  EXPECT_FALSE(sheet.view().show_grid_lines);
  EXPECT_TRUE(sheet.view().tab_selected);

  // Merge recovered.
  ASSERT_EQ(sheet.merges().size(), 1U);

  // Protection recovered, including the explicit unlock (formatCells="0").
  EXPECT_TRUE(sheet.protection().enabled);
  EXPECT_FALSE(sheet.protection().format_cells);
  EXPECT_TRUE(sheet.protection().sort) << "omitted lock-by-default flag reads as locked";

  // autoFilter + print settings recovered.
  EXPECT_FALSE(sheet.auto_filter_xml().empty());
  EXPECT_FALSE(sheet.print_settings().print_options_xml.empty());
  EXPECT_FALSE(sheet.print_settings().page_setup_xml.empty());

  // Japanese header string preserved verbatim.
  EXPECT_NE(sheet.print_settings().header_footer_xml.find(
                "\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E\xE3\x83\x98\xE3\x83\x83\xE3\x83\x80"),
            std::string::npos);

  // Range-ref hyperlink loaded (would have aborted the whole read before).
  ASSERT_EQ(sheet.hyperlinks().size(), 1U);
  EXPECT_EQ(sheet.hyperlinks()[0].row, 0U);
  EXPECT_EQ(sheet.hyperlinks()[0].col, 0U);
  EXPECT_EQ(sheet.hyperlinks()[0].last_row, 1U);
  EXPECT_EQ(sheet.hyperlinks()[0].last_col, 1U);
}

// A compact sheet exercising every metadata surface plus shared / array
// formulas. Read once via each path (threshold injected) and compared,
// this guards against a reader being wired into only one path — the exact
// defect class IO-F fixed. The threshold seam forces SAX without a
// multi-hundred-KiB fixture.
constexpr std::string_view kParitySheet =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
    "<sheetViews><sheetView showGridLines=\"0\" tabSelected=\"1\" workbookViewId=\"0\">"
    "<pane xSplit=\"1\" ySplit=\"2\" state=\"frozen\"/></sheetView></sheetViews>\n"
    "<cols><col min=\"1\" max=\"1\" width=\"20\" customWidth=\"1\"/></cols>\n"
    "<sheetData>\n"
    "<row r=\"1\" ht=\"30\" customHeight=\"1\"><c r=\"A1\"><v>10</v></c>"
    "<c r=\"B1\"><f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*2</f><v>20</v></c>"
    "<c r=\"C1\"><f t=\"array\" ref=\"C1\">SUM(A1:A3)</f><v>33</v></c></row>\n"
    "<row r=\"2\" hidden=\"1\"><c r=\"A2\"><v>11</v></c><c r=\"B2\"><f t=\"shared\" si=\"0\"/><v>22</v></c></row>\n"
    "<row r=\"3\" outlineLevel=\"1\"><c r=\"A3\"><v>12</v></c><c r=\"B3\"><f t=\"shared\" "
    "si=\"0\"/><v>24</v></c></row>\n"
    "</sheetData>\n"
    "<sheetProtection sheet=\"1\" formatCells=\"0\"/>\n"
    "<autoFilter ref=\"A1:C1\"/>\n"
    "<mergeCells count=\"1\"><mergeCell ref=\"D1:E1\"/></mergeCells>\n"
    "<conditionalFormatting sqref=\"A1:A3\"><cfRule type=\"cellIs\" operator=\"greaterThan\" priority=\"1\">"
    "<formula>5</formula></cfRule></conditionalFormatting>\n"
    "<hyperlinks><hyperlink ref=\"A1:B2\" location=\"Sheet1!A1\"/></hyperlinks>\n"
    "<printOptions horizontalCentered=\"1\" gridLines=\"1\"/>\n"
    "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\n"
    "<pageSetup orientation=\"landscape\"/>\n"
    "<rowBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"1\" max=\"16383\" man=\"1\"/></rowBreaks>\n"
    "<headerFooter><oddHeader>&amp;C\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E</oddHeader></headerFooter>\n"
    "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\">"
    "<x14:conditionalFormattings "
    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
    "<x14:conditionalFormatting "
    "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
    "<x14:cfRule type=\"dataBar\" id=\"{DB000000-0000-0000-0000-000000000001}\"><x14:dataBar/></x14:cfRule>"
    "<xm:sqref>A1:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>\n"
    "</worksheet>\n";

// Asserts the two sheets agree on every metadata surface both read paths
// populate.
void ExpectSheetsEquivalent(const Sheet& dom, const Sheet& sax) {
  EXPECT_EQ(dom.view().freeze_rows, sax.view().freeze_rows);
  EXPECT_EQ(dom.view().freeze_cols, sax.view().freeze_cols);
  EXPECT_EQ(dom.view().show_grid_lines, sax.view().show_grid_lines);
  EXPECT_EQ(dom.view().tab_selected, sax.view().tab_selected);
  ASSERT_EQ(dom.layout().columns.size(), sax.layout().columns.size());
  for (std::size_t i = 0; i < dom.layout().columns.size(); ++i) {
    EXPECT_EQ(dom.layout().columns[i].first, sax.layout().columns[i].first);
    EXPECT_DOUBLE_EQ(dom.layout().columns[i].width, sax.layout().columns[i].width);
  }
  // Per-row overrides (height / hidden / outline) — the IO-F SAX residual
  // now recovered via the row-start callback.
  ASSERT_EQ(dom.layout().row_overrides.size(), sax.layout().row_overrides.size());
  for (std::size_t i = 0; i < dom.layout().row_overrides.size(); ++i) {
    EXPECT_EQ(dom.layout().row_overrides[i].row, sax.layout().row_overrides[i].row);
    EXPECT_DOUBLE_EQ(dom.layout().row_overrides[i].height, sax.layout().row_overrides[i].height);
    EXPECT_EQ(dom.layout().row_overrides[i].hidden, sax.layout().row_overrides[i].hidden);
    EXPECT_EQ(dom.layout().row_overrides[i].outline_level, sax.layout().row_overrides[i].outline_level);
  }
  // Worksheet-level extLst (x14 conditional-formatting data).
  EXPECT_EQ(dom.ext_lst_xml(), sax.ext_lst_xml());
  ASSERT_EQ(dom.merges().size(), sax.merges().size());
  EXPECT_EQ(dom.conditional_formats().size(), sax.conditional_formats().size());
  EXPECT_EQ(dom.protection().enabled, sax.protection().enabled);
  EXPECT_EQ(dom.protection().sheet, sax.protection().sheet);
  EXPECT_EQ(dom.protection().format_cells, sax.protection().format_cells);
  EXPECT_EQ(dom.protection().sort, sax.protection().sort);
  ASSERT_EQ(dom.hyperlinks().size(), sax.hyperlinks().size());
  for (std::size_t i = 0; i < dom.hyperlinks().size(); ++i) {
    EXPECT_EQ(dom.hyperlinks()[i].row, sax.hyperlinks()[i].row);
    EXPECT_EQ(dom.hyperlinks()[i].col, sax.hyperlinks()[i].col);
    EXPECT_EQ(dom.hyperlinks()[i].last_row, sax.hyperlinks()[i].last_row);
    EXPECT_EQ(dom.hyperlinks()[i].last_col, sax.hyperlinks()[i].last_col);
    EXPECT_EQ(dom.hyperlinks()[i].location, sax.hyperlinks()[i].location);
  }
  EXPECT_EQ(dom.print_settings().header_footer_xml, sax.print_settings().header_footer_xml);
  EXPECT_EQ(dom.print_settings().print_options_xml, sax.print_settings().print_options_xml);
  EXPECT_EQ(dom.print_settings().page_setup_xml, sax.print_settings().page_setup_xml);
  EXPECT_EQ(dom.auto_filter_xml(), sax.auto_filter_xml());
  ASSERT_EQ(dom.print_settings().manual_row_breaks.size(), sax.print_settings().manual_row_breaks.size());
  for (std::size_t i = 0; i < dom.print_settings().manual_row_breaks.size(); ++i) {
    EXPECT_EQ(dom.print_settings().manual_row_breaks[i].id, sax.print_settings().manual_row_breaks[i].id);
    EXPECT_EQ(dom.print_settings().manual_row_breaks[i].manual, sax.print_settings().manual_row_breaks[i].manual);
  }
  // Formulas: shared follower (B2) + array master (C1).
  for (const std::pair<std::uint32_t, std::uint32_t>& rc :
       {std::pair<std::uint32_t, std::uint32_t>{1U, 1U}, std::pair<std::uint32_t, std::uint32_t>{0U, 2U}}) {
    const Cell* d = dom.cell_at(rc.first, rc.second);
    const Cell* s = sax.cell_at(rc.first, rc.second);
    ASSERT_NE(d, nullptr);
    ASSERT_NE(s, nullptr);
    EXPECT_EQ(d->formula_text, s->formula_text) << "cell (" << rc.first << "," << rc.second << ")";
  }
}

TEST(OoxmlSaxMetadata, DomAndSaxPathsAgree) {
  if (io::kSaxThresholdBytes == static_cast<std::size_t>(-1)) {
    GTEST_SKIP() << "SAX path disabled in this build";
  }
  const std::vector<std::uint8_t> pkg = BuildPackage(std::string(kParitySheet));

  // DOM path: SIZE_MAX threshold forces the DOM branch.
  auto dom_or = io::internal::ReadOoxmlWithSaxThresholdForTesting(SpanOf(pkg), static_cast<std::size_t>(-1));
  ASSERT_TRUE(static_cast<bool>(dom_or)) << dom_or.error().message;
  // SAX path: threshold 0 forces the streaming branch on this small sheet.
  auto sax_or = io::internal::ReadOoxmlWithSaxThresholdForTesting(SpanOf(pkg), 0U);
  ASSERT_TRUE(static_cast<bool>(sax_or)) << sax_or.error().message;

  ASSERT_EQ(dom_or.value().workbook.sheet_count(), 1U);
  ASSERT_EQ(sax_or.value().workbook.sheet_count(), 1U);
  ExpectSheetsEquivalent(dom_or.value().workbook.sheet(0), sax_or.value().workbook.sheet(0));

  // Sanity: the shared follower actually resolved (not left blank) on both.
  EXPECT_EQ(dom_or.value().workbook.sheet(0).cell_at(1U, 1U)->formula_text, "=A2*2");
  EXPECT_EQ(sax_or.value().workbook.sheet(0).cell_at(1U, 1U)->formula_text, "=A2*2");
}

TEST(OoxmlSaxMetadata, EmptyNumericValueHasSameBlankMeaningOnBothPaths) {
  if (io::kSaxThresholdBytes == static_cast<std::size_t>(-1)) {
    GTEST_SKIP() << "SAX path disabled in this build";
  }
  const std::vector<std::uint8_t> pkg = BuildPackage(
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<sheetData><row r=\"1\"><c r=\"A1\"><v/></c></row></sheetData></worksheet>");

  auto dom_or = io::internal::ReadOoxmlWithSaxThresholdForTesting(SpanOf(pkg), static_cast<std::size_t>(-1));
  auto sax_or = io::internal::ReadOoxmlWithSaxThresholdForTesting(SpanOf(pkg), 0U);
  ASSERT_TRUE(static_cast<bool>(dom_or)) << dom_or.error().message;
  ASSERT_TRUE(static_cast<bool>(sax_or)) << sax_or.error().message;

  const Cell* dom = dom_or.value().workbook.sheet(0).cell_at(0U, 0U);
  const Cell* sax = sax_or.value().workbook.sheet(0).cell_at(0U, 0U);
  // Blank cells are omitted from sparse storage; whether a future storage
  // implementation materialises them or not, both readers must agree.
  EXPECT_EQ(dom != nullptr, sax != nullptr);
  if (dom != nullptr) {
    EXPECT_TRUE(dom->cached_value.is_blank());
  }
  if (sax != nullptr) {
    EXPECT_TRUE(sax->cached_value.is_blank());
  }
}

// Real Excel writes a spilling `=SEQUENCE(3)` at F6 as
// `<f t="array" ref="F6:F8">` with F7/F8 as bare cached `<v>` cells. The
// reader must register F6:F8 as a spill region so the cached targets do
// not read back as literals that collide with the anchor's re-spill on
// recalc (`#SPILL!`). Verified on both read paths.
TEST(OoxmlSaxMetadata, ArraySpillFromCacheDoesNotCollideOnRecalc) {
  if (io::kSaxThresholdBytes == static_cast<std::size_t>(-1)) {
    GTEST_SKIP() << "SAX path disabled in this build";
  }
  constexpr std::string_view kArraySheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"6\"><c r=\"F6\"><f t=\"array\" ref=\"F6:F8\">SEQUENCE(3)</f><v>1</v></c></row>\n"
      "<row r=\"7\"><c r=\"F7\"><v>2</v></c></row>\n"
      "<row r=\"8\"><c r=\"F8\"><v>3</v></c></row>\n"
      "</sheetData>\n</worksheet>\n";
  const std::vector<std::uint8_t> pkg = BuildPackage(std::string(kArraySheet));

  // threshold SIZE_MAX -> DOM; threshold 0 -> SAX. Both must behave the same.
  for (const std::size_t threshold : {static_cast<std::size_t>(-1), static_cast<std::size_t>(0)}) {
    auto load_or = io::internal::ReadOoxmlWithSaxThresholdForTesting(SpanOf(pkg), threshold);
    ASSERT_TRUE(static_cast<bool>(load_or)) << load_or.error().message;
    Workbook wb = std::move(load_or.value().workbook);
    const Sheet& sheet = wb.sheet(0);

    // Before recalc, F7/F8 are spill phantoms carrying the cached values —
    // not blocking literals.
    EXPECT_DOUBLE_EQ(sheet.resolve_cell_value(6U, 5U).as_number(), 2.0)
        << "F7 cached phantom (threshold=" << threshold << ")";

    auto recalc_or = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(recalc_or)) << (recalc_or ? "" : recalc_or.error().message);

    // After recalc the anchor spills 1/2/3 with no #SPILL! collision.
    const Value f6 = wb.sheet(0).resolve_cell_value(5U, 5U);
    ASSERT_TRUE(f6.is_number()) << "F6 must spill to 1, not #SPILL! (threshold=" << threshold << ")";
    EXPECT_DOUBLE_EQ(f6.as_number(), 1.0);
    EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(6U, 5U).as_number(), 2.0);
    EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(7U, 5U).as_number(), 3.0);
  }
}

}  // namespace
}  // namespace formulon
