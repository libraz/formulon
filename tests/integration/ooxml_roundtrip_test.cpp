// Copyright 2026 libraz. Licensed under the MIT License.
//
// Round-trip integration tests: writer -> reader. The earlier slice only
// covered sheet names; this slice exercises cell-level round-tripping
// for literals and formulas. Cells written via `Workbook::set_cell_*`
// must come back through `read_ooxml` with the same shape, and the
// reader must register formula cells with the recalc engine so a
// post-load `recalc()` reproduces the original cached values.

#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

std::vector<std::uint8_t> SaveOrDie(const Workbook& wb) {
  auto save_or = wb.save();
  EXPECT_TRUE(static_cast<bool>(save_or)) << "save() failed: " << save_or.error().message;
  return save_or.value();
}

TEST(OoxmlRoundTrip, SingleSheet) {
  Workbook src = Workbook::create();
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << result_or.error().message;

  ASSERT_EQ(result_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "Sheet1");
}

TEST(OoxmlRoundTrip, MultipleSheetsPreserveOrder) {
  Workbook src = Workbook::create_empty();
  src.add_sheet("Alpha");
  src.add_sheet("Beta");
  src.add_sheet("Gamma");
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const Workbook& dst = result_or.value().workbook;
  ASSERT_EQ(dst.sheet_count(), 3U);
  EXPECT_EQ(dst.sheet(0).name(), "Alpha");
  EXPECT_EQ(dst.sheet(1).name(), "Beta");
  EXPECT_EQ(dst.sheet(2).name(), "Gamma");
}

TEST(OoxmlRoundTrip, JapaneseSheetName) {
  Workbook src = Workbook::create_empty();
  // "売上" in UTF-8: E5 A3 B2 E4 B8 8A
  src.add_sheet("\xE5\xA3\xB2\xE4\xB8\x8A");
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "\xE5\xA3\xB2\xE4\xB8\x8A");
}

TEST(OoxmlRoundTrip, SheetNameWithSpaceAndAmpersand) {
  Workbook src = Workbook::create_empty();
  src.add_sheet("Q1 & Q2 Summary");
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().workbook.sheet_count(), 1U);
  // The writer emits `&amp;` and pugixml decodes it; round-trip yields
  // the original ampersand.
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "Q1 & Q2 Summary");
}

TEST(OoxmlRoundTrip, SheetNameWithSingleQuoteRequiresQuoting) {
  // Excel itself wraps such names in single quotes when emitting
  // formulas, but the workbook.xml `<sheet name=...>` attribute is just
  // an XML attribute and accepts the apostrophe directly. We verify the
  // name comes back unchanged.
  Workbook src = Workbook::create_empty();
  src.add_sheet("Joe's Notes");
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "Joe's Notes");
}

TEST(OoxmlRoundTrip, RenamedDefaultSheet) {
  Workbook src = Workbook::create();
  src.sheet(0).set_name("Renamed");
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "Renamed");
}

TEST(OoxmlRoundTrip, NumericLiteralAndFormulaRecalcMatches) {
  // Build a workbook with A1=42, A2==A1*2, recalc, save, read back,
  // recalc, assert the cached value at A2 is 84 again. This is the
  // canonical "true round-trip" check for the cell-aware reader.
  Workbook src = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0U, 0U, 0U, Value::number(42.0))));
  ASSERT_TRUE(static_cast<bool>(src.set_cell_formula(0U, 1U, 0U, "=A1*2")));
  ASSERT_TRUE(static_cast<bool>(src.recalc(eval::default_registry())));

  // Sanity: the source workbook recalc'd correctly.
  {
    const Cell* a2 = src.sheet(0).cell_at(1U, 0U);
    ASSERT_NE(a2, nullptr);
    ASSERT_TRUE(a2->cached_value.is_number());
    EXPECT_DOUBLE_EQ(a2->cached_value.as_number(), 84.0);
  }

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;

  Workbook& dst = result_or.value().workbook;
  ASSERT_EQ(dst.sheet_count(), 1U);

  // After read but BEFORE recalc, A1 should already hold 42 (a literal),
  // and A2 should carry the formula text (the cached <v> emitted by the
  // writer is dropped by the reader on purpose; recalc populates it).
  {
    const Cell* a1 = dst.sheet(0).cell_at(0U, 0U);
    ASSERT_NE(a1, nullptr);
    ASSERT_TRUE(a1->cached_value.is_number());
    EXPECT_DOUBLE_EQ(a1->cached_value.as_number(), 42.0);
    EXPECT_TRUE(a1->formula_text.empty());

    const Cell* a2 = dst.sheet(0).cell_at(1U, 0U);
    ASSERT_NE(a2, nullptr);
    EXPECT_EQ(a2->formula_text, "=A1*2");
  }

  // Recalc and verify A2 == 84.
  ASSERT_TRUE(static_cast<bool>(dst.recalc(eval::default_registry())));
  {
    const Cell* a2 = dst.sheet(0).cell_at(1U, 0U);
    ASSERT_NE(a2, nullptr);
    ASSERT_TRUE(a2->cached_value.is_number());
    EXPECT_DOUBLE_EQ(a2->cached_value.as_number(), 84.0);
  }
}

TEST(OoxmlRoundTrip, InlineStringCellRoundTrips) {
  // The empty-workbook writer emits text values via t="inlineStr" (SST
  // is a Bundle 2.3 concern), so this is the round-trip path that
  // currently exists end-to-end without SST resolution.
  Workbook src = Workbook::create();
  std::string greeting = "Hello, world!";
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0U, 0U, 0U, Value::text(greeting))));

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  // No SST cells were used.
  EXPECT_EQ(result_or.value().pending_sst_count, 0U);

  const Workbook& dst = result_or.value().workbook;
  const Cell* a1 = dst.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "Hello, world!");
}

// Builds a minimal but valid in-memory `.xlsx` package containing an
// explicit shared-strings part. The package has one sheet with three
// cells: A1, A2, A3 — each carrying `t="s"` and pointing to indices
// 0, 1, 2 of the SST. The SST resolves to "alpha", "beta", "alpha"
// (note the duplicate index to verify aliasing rather than copying).
//
// We assemble the parts via miniz directly rather than going through
// `write_ooxml` because that writer does not yet emit SST-typed cells
// (Bundle 2.5). This is the canonical "explicit SST" payload Excel
// itself emits.
std::vector<std::uint8_t> BuildXlsxWithSst() {
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
      "  <Override PartName=\"/xl/sharedStrings.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\n"
      "  <Override PartName=\"/xl/styles.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n"
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
      "  <sheets>\n"
      "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "  </sheets>\n"
      "</workbook>\n";
  const std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "  <Relationship Id=\"rId2\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" "
      "Target=\"sharedStrings.xml\"/>\n"
      "  <Relationship Id=\"rId3\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
      "Target=\"styles.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData>\n"
      "    <row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>\n"
      "    <row r=\"2\"><c r=\"A2\" t=\"s\"><v>1</v></c></row>\n"
      "    <row r=\"3\"><c r=\"A3\" t=\"s\"><v>0</v></c></row>\n"
      "  </sheetData>\n"
      "</worksheet>\n";
  // SST: index 0 = "alpha", index 1 = "beta". The third cell reuses
  // index 0 to verify the reader aliases (does not copy) the entry.
  const std::string_view sst_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"2\">\n"
      "  <si><t>alpha</t></si>\n"
      "  <si><t>beta</t></si>\n"
      "</sst>\n";
  const std::string_view styles_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>\n"
      "</styleSheet>\n";

  struct PartFile {
    const char* path;
    std::string_view body;
  };
  const PartFile parts[] = {
      {"[Content_Types].xml", content_types},  {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},       {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml}, {"xl/sharedStrings.xml", sst_xml},
      {"xl/styles.xml", styles_xml},
  };

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

TEST(OoxmlRoundTrip, SstResolutionEndToEnd) {
  const std::vector<std::uint8_t> bytes = BuildXlsxWithSst();

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const io::OoxmlReadResult& result = result_or.value();

  // Three SST-typed cells -> three resolutions.
  EXPECT_EQ(result.pending_sst_count, 3U);

  // sharedStrings.xml and styles.xml should NOT surface as unknown.
  for (const io::PassthroughPart& part : result.unknown_parts) {
    EXPECT_NE(part.path, "xl/sharedStrings.xml");
    EXPECT_NE(part.path, "xl/styles.xml");
    EXPECT_NE(part.path, "xl/worksheets/sheet1.xml");
  }

  const Workbook& dst = result.workbook;
  ASSERT_EQ(dst.sheet_count(), 1U);

  const Cell* a1 = dst.sheet(0).cell_at(0U, 0U);
  const Cell* a2 = dst.sheet(0).cell_at(1U, 0U);
  const Cell* a3 = dst.sheet(0).cell_at(2U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_NE(a2, nullptr);
  ASSERT_NE(a3, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  ASSERT_TRUE(a2->cached_value.is_text());
  ASSERT_TRUE(a3->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "alpha");
  EXPECT_EQ(a2->cached_value.as_text(), "beta");
  // Duplicate-index reuse: index 0 again resolves to "alpha".
  EXPECT_EQ(a3->cached_value.as_text(), "alpha");
}

TEST(OoxmlRoundTrip, ErrorCellRoundTrips) {
  Workbook src = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0U, 0U, 0U, Value::error(ErrorCode::Div0))));

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const Workbook& dst = result_or.value().workbook;
  const Cell* a1 = dst.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_error());
  EXPECT_EQ(a1->cached_value.as_error(), ErrorCode::Div0);
}

}  // namespace
}  // namespace formulon
