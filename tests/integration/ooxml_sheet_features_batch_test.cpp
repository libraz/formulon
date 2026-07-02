// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Integration test: worksheet-level elements and attributes that the
// writer previously dropped must round-trip through a load->save cycle:
//   * `<sheetView>` display attributes (showGridLines, tabSelected, view)
//   * `<sheetProtection>` with an explicit unlock of a lock-by-default
//     action (formatCells="0") — the case a "emit-when-true" writer
//     silently re-locked
//   * `<autoFilter>`, `<printOptions>`, `<headerFooter>` raw capture
//   * range-ref hyperlinks (`ref="A1:B2"`)
//   * `<tableStyleInfo>` on a table part

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pugixml.hpp"
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

std::vector<std::uint8_t> ReadPart(io::ZipReader& zip, std::string_view name) {
  auto bytes_or = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << "missing part: " << name;
  if (!bytes_or) {
    return {};
  }
  return bytes_or.value();
}

// Single-sheet package whose worksheet body is supplied by the caller.
std::vector<std::uint8_t> BuildPackage(std::string_view sheet_body) {
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
      {"xl/worksheets/sheet1.xml", sheet_body},
  });
}

// Reloads the workbook produced by saving `source`, returning the parsed
// worksheet DOM through `out_doc` / `out_bytes` (kept alive by the caller).
pugi::xml_node RoundTripWorksheet(const std::vector<std::uint8_t>& source, pugi::xml_document& out_doc,
                                  std::vector<std::uint8_t>& out_bytes) {
  auto load_or = io::read_ooxml(SpanOf(source));
  EXPECT_TRUE(static_cast<bool>(load_or)) << (load_or ? "" : load_or.error().message);
  if (!load_or) {
    return {};
  }
  auto save_or = load_or.value().workbook.save();
  EXPECT_TRUE(static_cast<bool>(save_or)) << (save_or ? "" : save_or.error().message);
  if (!save_or) {
    return {};
  }
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  out_bytes = ReadPart(zip, "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(out_doc.load_buffer(out_bytes.data(), out_bytes.size()));
  return out_doc.child("worksheet");
}

constexpr std::string_view kSheetWithFeatures =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
    "  <sheetViews><sheetView showGridLines=\"0\" tabSelected=\"1\" view=\"pageBreakPreview\" "
    "workbookViewId=\"0\"/></sheetViews>\n"
    "  <sheetData/>\n"
    "  <sheetProtection sheet=\"1\" formatCells=\"0\"/>\n"
    "  <autoFilter ref=\"A1:C1\"/>\n"
    "  <hyperlinks><hyperlink ref=\"A1:B2\" location=\"Sheet1!A1\"/></hyperlinks>\n"
    "  <printOptions horizontalCentered=\"1\" gridLines=\"1\"/>\n"
    "  <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\n"
    "  <pageSetup orientation=\"landscape\"/>\n"
    "  <headerFooter><oddHeader>&amp;LLeft&amp;RRight</oddHeader></headerFooter>\n"
    "</worksheet>\n";

TEST(OoxmlSheetFeaturesBatch, WorksheetElementsRoundTrip) {
  pugi::xml_document doc;
  std::vector<std::uint8_t> bytes;
  pugi::xml_node ws = RoundTripWorksheet(BuildPackage(kSheetWithFeatures), doc, bytes);
  ASSERT_TRUE(ws);

  // sheetView display attributes preserved.
  pugi::xml_node sv = ws.child("sheetViews").child("sheetView");
  ASSERT_TRUE(sv);
  EXPECT_STREQ(sv.attribute("showGridLines").value(), "0");
  EXPECT_STREQ(sv.attribute("tabSelected").value(), "1");
  EXPECT_STREQ(sv.attribute("view").value(), "pageBreakPreview");

  // sheetProtection: the explicit unlock of a lock-by-default action must
  // survive as formatCells="0" (not dropped, which would re-lock it).
  pugi::xml_node sp = ws.child("sheetProtection");
  ASSERT_TRUE(sp);
  EXPECT_STREQ(sp.attribute("sheet").value(), "1");
  EXPECT_STREQ(sp.attribute("formatCells").value(), "0");
  // A lock-by-default action left at its default must NOT be emitted.
  EXPECT_TRUE(sp.attribute("sort").empty());

  // autoFilter / printOptions / headerFooter preserved.
  EXPECT_STREQ(ws.child("autoFilter").attribute("ref").value(), "A1:C1");
  ASSERT_TRUE(ws.child("printOptions"));
  EXPECT_STREQ(ws.child("printOptions").attribute("gridLines").value(), "1");
  ASSERT_TRUE(ws.child("headerFooter"));
  EXPECT_STREQ(ws.child("headerFooter").child("oddHeader").text().get(), "&LLeft&RRight");

  // Range hyperlink ref preserved verbatim.
  pugi::xml_node hl = ws.child("hyperlinks").child("hyperlink");
  ASSERT_TRUE(hl);
  EXPECT_STREQ(hl.attribute("ref").value(), "A1:B2");

  // ECMA-376 order: autoFilter before mergeCells/hyperlinks region,
  // printOptions before pageMargins, headerFooter after pageSetup.
  int idx = 0;
  int i_autofilter = -1;
  int i_hyperlinks = -1;
  int i_printoptions = -1;
  int i_pagemargins = -1;
  int i_pagesetup = -1;
  int i_headerfooter = -1;
  for (pugi::xml_node n = ws.first_child(); n; n = n.next_sibling(), ++idx) {
    const std::string_view name = n.name();
    if (name == "autoFilter") {
      i_autofilter = idx;
    } else if (name == "hyperlinks") {
      i_hyperlinks = idx;
    } else if (name == "printOptions") {
      i_printoptions = idx;
    } else if (name == "pageMargins") {
      i_pagemargins = idx;
    } else if (name == "pageSetup") {
      i_pagesetup = idx;
    } else if (name == "headerFooter") {
      i_headerfooter = idx;
    }
  }
  EXPECT_LT(i_autofilter, i_hyperlinks);
  EXPECT_LT(i_hyperlinks, i_printoptions);
  EXPECT_LT(i_printoptions, i_pagemargins);
  EXPECT_LT(i_pagemargins, i_pagesetup);
  EXPECT_LT(i_pagesetup, i_headerfooter);
}

TEST(OoxmlSheetFeaturesBatch, SheetPrRoundTripsWithoutPageSetUpPr) {
  // A `<sheetPr>` carrying only tabColor / codeName (no <pageSetUpPr>) must
  // survive a save cycle — the capture used to gate on <pageSetUpPr> and
  // dropped these sheets' <sheetPr> wholesale.
  struct Case {
    std::string_view sheet_pr;
    std::string_view must_contain;
  };
  const Case cases[] = {
      {"<sheetPr codeName=\"ThisSheet\"/>", "codeName=\"ThisSheet\""},
      {"<sheetPr><tabColor rgb=\"FFFF0000\"/></sheetPr>", "tabColor"},
      {"<sheetPr codeName=\"S\"><tabColor rgb=\"FF00FF00\"/><pageSetUpPr fitToPage=\"1\"/></sheetPr>", "pageSetUpPr"},
  };
  for (const Case& c : cases) {
    std::string body =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n  ";
    body.append(c.sheet_pr);
    body.append("\n  <sheetData/>\n</worksheet>\n");

    pugi::xml_document doc;
    std::vector<std::uint8_t> bytes;
    pugi::xml_node ws = RoundTripWorksheet(BuildPackage(body), doc, bytes);
    ASSERT_TRUE(ws) << c.sheet_pr;
    pugi::xml_node sp = ws.child("sheetPr");
    ASSERT_TRUE(sp) << "<sheetPr> dropped on save for: " << c.sheet_pr;
    const std::string_view out(reinterpret_cast<const char*>(bytes.data()), bytes.size());
    EXPECT_NE(out.find(c.must_contain), std::string_view::npos) << "lost content for: " << c.sheet_pr;
  }
}

TEST(OoxmlSheetFeaturesBatch, MinimalProtectionAppliesLockByDefault) {
  // A protected sheet that omits the lock-by-default actions must be read
  // as those actions LOCKED (true), matching Excel, and re-emit them as
  // defaults (i.e. omitted), not as explicit "0".
  constexpr std::string_view kMinimalProtection =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "  <sheetProtection sheet=\"1\"/>\n"
      "</worksheet>\n";
  auto load_or = io::read_ooxml(SpanOf(BuildPackage(kMinimalProtection)));
  ASSERT_TRUE(static_cast<bool>(load_or)) << load_or.error().message;
  const SheetProtection& p = load_or.value().workbook.sheet(0).protection();
  EXPECT_TRUE(p.enabled);
  // Lock-by-default actions read as locked despite being omitted.
  EXPECT_TRUE(p.format_cells);
  EXPECT_TRUE(p.sort);
  EXPECT_TRUE(p.auto_filter);
  // False-default actions stay false.
  EXPECT_FALSE(p.objects);
  EXPECT_FALSE(p.select_locked_cells);

  pugi::xml_document doc;
  std::vector<std::uint8_t> bytes;
  pugi::xml_node ws = RoundTripWorksheet(BuildPackage(kMinimalProtection), doc, bytes);
  ASSERT_TRUE(ws);
  pugi::xml_node sp = ws.child("sheetProtection");
  ASSERT_TRUE(sp);
  EXPECT_STREQ(sp.attribute("sheet").value(), "1");
  // At-default lock flags are not re-emitted.
  EXPECT_TRUE(sp.attribute("formatCells").empty());
  EXPECT_TRUE(sp.attribute("sort").empty());
}

TEST(OoxmlSheetFeaturesBatch, WorksheetExtLstRoundTrips) {
  // The worksheet-level <extLst> holds x14 conditional-formatting data
  // (DataBar 2010+ negative-fill / axis / gradient), linked to the legacy
  // cfRule by an x14:id GUID. It must survive load->save verbatim and be
  // re-emitted after <tableParts> at the worksheet tail.
  constexpr std::string_view kSheetWithExtLst =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "  <conditionalFormatting sqref=\"A1:A5\"><cfRule type=\"dataBar\" priority=\"1\"><dataBar>"
      "<cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF638EC6\"/></dataBar>"
      "<extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\">"
      "<x14:id>{DB000000-0000-0000-0000-000000000001}</x14:id></ext></extLst></cfRule></conditionalFormatting>\n"
      "  <extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\">"
      "<x14:conditionalFormattings "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
      "<x14:conditionalFormatting "
      "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
      "<x14:cfRule type=\"dataBar\" id=\"{DB000000-0000-0000-0000-000000000001}\">"
      "<x14:dataBar minLength=\"0\" maxLength=\"100\" negativeBarColorSameAsPositive=\"0\">"
      "<x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/>"
      "<x14:negativeFillColor rgb=\"FFFF0000\"/><x14:axisColor rgb=\"FF000000\"/></x14:dataBar>"
      "</x14:cfRule><xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>"
      "</x14:conditionalFormattings></ext></extLst>\n"
      "</worksheet>\n";
  pugi::xml_document doc;
  std::vector<std::uint8_t> bytes;
  pugi::xml_node ws = RoundTripWorksheet(BuildPackage(kSheetWithExtLst), doc, bytes);
  ASSERT_TRUE(ws);

  // The worksheet-level extLst survives and sits at the tail.
  pugi::xml_node ext = ws.child("extLst");
  ASSERT_TRUE(ext) << "worksheet <extLst> dropped on save";
  // The x14:id GUID linking to the legacy cfRule is preserved verbatim.
  const std::string_view out(reinterpret_cast<const char*>(bytes.data()), bytes.size());
  EXPECT_NE(out.find("{DB000000-0000-0000-0000-000000000001}"), std::string_view::npos) << "x14:id GUID linkage lost";
  EXPECT_NE(out.find("negativeFillColor"), std::string_view::npos) << "x14 DataBar negative-fill data lost";
  EXPECT_NE(out.find("x14:conditionalFormattings"), std::string_view::npos);
}

TEST(OoxmlSheetFeaturesBatch, TableStyleInfoRoundTrips) {
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
      "  <Override PartName=\"/xl/tables/table1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>\n"
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
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "  <tableParts count=\"1\">\n"
      "    <tablePart r:id=\"rId1\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>\n"
      "  </tableParts>\n"
      "</worksheet>\n";
  const std::string_view sheet_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" "
      "Target=\"../tables/table1.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view table_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Sales\" "
      "displayName=\"Sales\" ref=\"A1:B2\">\n"
      "  <tableColumns count=\"2\">\n"
      "    <tableColumn id=\"1\" name=\"A\"/>\n"
      "    <tableColumn id=\"2\" name=\"B\"/>\n"
      "  </tableColumns>\n"
      "  <tableStyleInfo name=\"TableStyleMedium2\" showFirstColumn=\"0\" showLastColumn=\"0\" "
      "showRowStripes=\"1\" showColumnStripes=\"0\"/>\n"
      "</table>\n";
  const std::vector<std::uint8_t> source = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/worksheets/_rels/sheet1.xml.rels", sheet_rels},
      {"xl/tables/table1.xml", table_xml},
  });

  auto load_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(load_or)) << load_or.error().message;
  ASSERT_EQ(load_or.value().workbook.tables().size(), 1U);
  EXPECT_NE(load_or.value().workbook.tables()[0].table_style_info_xml.find("TableStyleMedium2"), std::string::npos);

  auto save_or = load_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << save_or.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  const std::vector<std::uint8_t> out = ReadPart(zip, "xl/tables/table1.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_buffer(out.data(), out.size()));
  pugi::xml_node style = doc.child("table").child("tableStyleInfo");
  ASSERT_TRUE(style) << "tableStyleInfo dropped on save";
  EXPECT_STREQ(style.attribute("name").value(), "TableStyleMedium2");
  EXPECT_STREQ(style.attribute("showRowStripes").value(), "1");
}

}  // namespace
}  // namespace formulon
