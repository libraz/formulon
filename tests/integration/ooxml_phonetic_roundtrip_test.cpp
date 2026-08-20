//
// Round-trip integration test for the OOXML <rPh> (phonetic guide)
// pipeline. Covers two halves:
//
//   1. The reader: a synthetic xlsx package whose `xl/sharedStrings.xml`
//      contains an annotated `<si>` and whose `xl/worksheets/sheet1.xml`
//      contains an inline-string cell with an inline `<rPh>` block. After
//      `read_ooxml`, both cells must carry a non-empty
//      `Cell::phonetic_text`.
//
//   2. The writer: a workbook constructed in memory with `phonetic_text`
//      set on a Text cell, saved via `Workbook::save`, and re-read. The
//      annotation must round-trip back to the destination cell. The
//      writer uses inline strings for literal Text cells, so the kana
//      lands in the worksheet's `<is>` block on save.
//
// Multi-block <rPh> structure (sb/eb spans) is intentionally lossy: we
// collapse the original blocks into a single concatenated kana string on
// read, and emit a single `<rPh sb="0" eb="N">` block on write. This is
// adequate because PHONETIC's observable behaviour depends only on the
// concatenated kana, not the per-span boundaries.

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "phonetic.h"
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

/// Builds a heap-allocated zip archive containing the supplied parts.
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

// ---------------------------------------------------------------------------
// Reader: <rPh> on SST entry lands on the cell
// ---------------------------------------------------------------------------

TEST(OoxmlPhoneticRoundTrip, ReaderSstAnnotationLandsOnCell) {
  // Single-sheet workbook. A1 references SST index 0 which carries
  // `<si><t>kanji</t><rPh sb="0" eb="2"><t>furigana</t></rPh></si>`.
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
      "</Relationships>\n";
  const std::string_view shared_strings =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <si><t>kanji</t><rPh sb=\"0\" eb=\"2\"><t>furigana</t></rPh></si>\n"
      "</sst>\n";
  const std::string_view sheet1_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData>\n"
      "    <row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>\n"
      "  </sheetData>\n"
      "</worksheet>\n";

  const std::vector<std::uint8_t> bytes = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/sharedStrings.xml", shared_strings},
      {"xl/worksheets/sheet1.xml", sheet1_xml},
  });

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& wb = result_or.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 1U);

  const Cell* cell = wb.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_text());
  EXPECT_EQ(cell->cached_value.as_text(), "kanji");
  ASSERT_EQ(cell->phonetic_runs.size(), 1U);
  EXPECT_EQ(cell->phonetic_runs[0].sb, 0U);
  EXPECT_EQ(cell->phonetic_runs[0].eb, 2U);
  EXPECT_EQ(cell->phonetic_runs[0].text, "furigana");
}

// ---------------------------------------------------------------------------
// Reader: inline-string <rPh> lands on the cell
// ---------------------------------------------------------------------------

TEST(OoxmlPhoneticRoundTrip, ReaderInlineStringAnnotationLandsOnCell) {
  // A1 carries an inline-string body with both surface text and an
  // <rPh> block, no SST involvement.
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
      "</Relationships>\n";
  const std::string_view sheet1_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData>\n"
      "    <row r=\"1\"><c r=\"A1\" t=\"inlineStr\">"
      "<is><t>kanji</t><rPh sb=\"0\" eb=\"2\"><t>furigana</t></rPh></is>"
      "</c></row>\n"
      "  </sheetData>\n"
      "</worksheet>\n";

  const std::vector<std::uint8_t> bytes = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet1_xml},
  });

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& wb = result_or.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 1U);

  const Cell* cell = wb.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_text());
  EXPECT_EQ(cell->cached_value.as_text(), "kanji");
  ASSERT_EQ(cell->phonetic_runs.size(), 1U);
  EXPECT_EQ(cell->phonetic_runs[0].sb, 0U);
  EXPECT_EQ(cell->phonetic_runs[0].eb, 2U);
  EXPECT_EQ(cell->phonetic_runs[0].text, "furigana");
}

// ---------------------------------------------------------------------------
// End to end: PHONETIC over a partially annotated shared string
// ---------------------------------------------------------------------------

TEST(OoxmlPhoneticRoundTrip, PhoneticSubstitutesAnnotatedSpanOnly) {
  // The package Mac Excel 365 ja-JP 16.112.1 was asked to open directly:
  // A1 is fully annotated, A2 is annotated over its first two characters
  // only, A3 carries the same text as A1 with no annotation. Excel
  // answered ニホンゴ / トウキョウ都 / 日本語 -- the middle one is the
  // case that pins the substitution to the span it names.
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
      "</Relationships>\n";
  // 日本語 / ニホンゴ over all three characters, 東京都 / トウキョウ over
  // the first two, then 日本語 bare.
  const std::string_view shared_strings =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <si><t>\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E</t>"
      "<rPh sb=\"0\" eb=\"3\"><t>\xE3\x83\x8B\xE3\x83\x9B\xE3\x83\xB3\xE3\x82\xB4</t></rPh></si>\n"
      "  <si><t>\xE6\x9D\xB1\xE4\xBA\xAC\xE9\x83\xBD</t>"
      "<rPh sb=\"0\" eb=\"2\"><t>\xE3\x83\x88\xE3\x82\xA6\xE3\x82\xAD\xE3\x83\xA7\xE3\x82\xA6</t></rPh></si>\n"
      "  <si><t>\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E</t></si>\n"
      "</sst>\n";
  const std::string_view sheet1_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData>\n"
      "    <row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"C1\"><f>PHONETIC(A1)</f></c></row>\n"
      "    <row r=\"2\"><c r=\"A2\" t=\"s\"><v>1</v></c><c r=\"C2\"><f>PHONETIC(A2)</f></c></row>\n"
      "    <row r=\"3\"><c r=\"A3\" t=\"s\"><v>2</v></c><c r=\"C3\"><f>PHONETIC(A3)</f></c></row>\n"
      "  </sheetData>\n"
      "</worksheet>\n";

  const std::vector<std::uint8_t> bytes = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/sharedStrings.xml", shared_strings},
      {"xl/worksheets/sheet1.xml", sheet1_xml},
  });

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  Workbook& wb = result_or.value().workbook;
  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << "recalc: " << recalc_or.error().message;

  const Cell* full = wb.sheet(0).cell_at(0, 2);
  const Cell* partial = wb.sheet(0).cell_at(1, 2);
  const Cell* bare = wb.sheet(0).cell_at(2, 2);
  ASSERT_NE(full, nullptr);
  ASSERT_NE(partial, nullptr);
  ASSERT_NE(bare, nullptr);
  // 日本語 -> ニホンゴ: the span covers the whole string, nothing left.
  EXPECT_EQ(full->cached_value.as_text(), "\xE3\x83\x8B\xE3\x83\x9B\xE3\x83\xB3\xE3\x82\xB4");
  // 東京都 -> トウキョウ都: 都 sits outside the span and passes through.
  EXPECT_EQ(partial->cached_value.as_text(),
            "\xE3\x83\x88\xE3\x82\xA6\xE3\x82\xAD\xE3\x83\xA7\xE3\x82\xA6\xE9\x83\xBD");
  // No annotation at all -> the surface text.
  EXPECT_EQ(bare->cached_value.as_text(), "\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E");
}

// ---------------------------------------------------------------------------
// Round-trip: writer emits <rPh>, reader reads it back
// ---------------------------------------------------------------------------

TEST(OoxmlPhoneticRoundTrip, WriterEmitsAnnotationAndReaderReadsItBack) {
  Workbook src = Workbook::create();
  src.sheet(0).set_cell_value(0, 0, Value::text("kanji"));
  src.sheet(0).set_cell_phonetic(0, 0, "furigana");

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save: " << save_or.error().message;
  const std::vector<std::uint8_t> bytes = save_or.value();

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& dst = result_or.value().workbook;
  ASSERT_EQ(dst.sheet_count(), 1U);

  const Cell* cell = dst.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_text());
  EXPECT_EQ(cell->cached_value.as_text(), "kanji");
  // `set_cell_phonetic` records a whole-string run, so the span covers
  // all five characters of "kanji" rather than the two the reader test
  // above supplies.
  ASSERT_EQ(cell->phonetic_runs.size(), 1U);
  EXPECT_EQ(cell->phonetic_runs[0].sb, 0U);
  EXPECT_EQ(cell->phonetic_runs[0].eb, 5U);
  EXPECT_EQ(cell->phonetic_runs[0].text, "furigana");
}

TEST(OoxmlPhoneticRoundTrip, WriterPreservesUnicodeAnnotation) {
  // Surface "山田" (UTF-8: E5 B1 B1 E7 94 B0), kana "やまだ" (UTF-8:
  // E3 82 84 E3 81 BE E3 81 A0). The surface "山田" is two BMP
  // codepoints, so the writer should emit `eb="2"`.
  Workbook src = Workbook::create();
  src.sheet(0).set_cell_value(0, 0, Value::text("\xE5\xB1\xB1\xE7\x94\xB0"));
  src.sheet(0).set_cell_phonetic(0, 0, "\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0");

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  const std::vector<std::uint8_t> bytes = save_or.value();

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const Workbook& dst = result_or.value().workbook;

  const Cell* cell = dst.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->cached_value.as_text(), "\xE5\xB1\xB1\xE7\x94\xB0");
  ASSERT_EQ(cell->phonetic_runs.size(), 1U);
  EXPECT_EQ(cell->phonetic_runs[0].sb, 0U);
  EXPECT_EQ(cell->phonetic_runs[0].eb, 2U);
  EXPECT_EQ(cell->phonetic_runs[0].text, "\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0");
}

TEST(OoxmlPhoneticRoundTrip, PartialAnnotationSurvivesWithItsSpan) {
  // 東京都 read over 東京 alone. The writer must emit the span as stored
  // rather than widening it to the whole string: Excel reads the kana
  // only over the characters the span names, so a widened block would
  // put トウキョウ over 都 as well.
  Workbook src = Workbook::create();
  src.sheet(0).set_cell_value(0, 0, Value::text("\xE6\x9D\xB1\xE4\xBA\xAC\xE9\x83\xBD"));
  src.sheet(0).set_cell_phonetic_runs(
      0, 0, {PhoneticRun{0U, 2U, "\xE3\x83\x88\xE3\x82\xA6\xE3\x82\xAD\xE3\x83\xA7\xE3\x82\xA6"}});

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save: " << save_or.error().message;

  auto result_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;

  const Cell* cell = result_or.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->cached_value.as_text(), "\xE6\x9D\xB1\xE4\xBA\xAC\xE9\x83\xBD");
  ASSERT_EQ(cell->phonetic_runs.size(), 1U);
  EXPECT_EQ(cell->phonetic_runs[0].sb, 0U);
  EXPECT_EQ(cell->phonetic_runs[0].eb, 2U);
  EXPECT_EQ(cell->phonetic_runs[0].text, "\xE3\x83\x88\xE3\x82\xA6\xE3\x82\xAD\xE3\x83\xA7\xE3\x82\xA6");
}

TEST(OoxmlPhoneticRoundTrip, MultipleSpansStayDistinctAcrossSave) {
  // Two blocks on one string item must not merge into one: their spans
  // are what tells Excel which reading covers which characters.
  Workbook src = Workbook::create();
  src.sheet(0).set_cell_value(0, 0, Value::text("\xE5\xB1\xB1\xE7\x94\xB0\xE5\xA4\xAA\xE9\x83\x8E"));
  src.sheet(0).set_cell_phonetic_runs(0, 0,
                                      {PhoneticRun{0U, 2U, "\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0"},
                                       PhoneticRun{2U, 4U, "\xE3\x81\x9F\xE3\x82\x8D\xE3\x81\x86"}});

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save: " << save_or.error().message;

  auto result_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;

  const Cell* cell = result_or.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  ASSERT_EQ(cell->phonetic_runs.size(), 2U);
  EXPECT_EQ(cell->phonetic_runs[0].sb, 0U);
  EXPECT_EQ(cell->phonetic_runs[0].eb, 2U);
  EXPECT_EQ(cell->phonetic_runs[0].text, "\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0");
  EXPECT_EQ(cell->phonetic_runs[1].sb, 2U);
  EXPECT_EQ(cell->phonetic_runs[1].eb, 4U);
  EXPECT_EQ(cell->phonetic_runs[1].text, "\xE3\x81\x9F\xE3\x82\x8D\xE3\x81\x86");
}

TEST(OoxmlPhoneticRoundTrip, DistinctAnnotationsOnEqualTextDoNotShareAnEntry) {
  // Two cells with the same surface text but different spans must intern
  // to different <si> entries, otherwise one cell's reading would be
  // applied to the other's characters.
  Workbook src = Workbook::create();
  src.sheet(0).set_cell_value(0, 0, Value::text("\xE6\x9D\xB1\xE4\xBA\xAC\xE9\x83\xBD"));
  src.sheet(0).set_cell_phonetic_runs(0, 0, {PhoneticRun{0U, 2U, "AB"}});
  src.sheet(0).set_cell_value(1, 0, Value::text("\xE6\x9D\xB1\xE4\xBA\xAC\xE9\x83\xBD"));
  src.sheet(0).set_cell_phonetic_runs(1, 0, {PhoneticRun{0U, 3U, "AB"}});

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save: " << save_or.error().message;

  auto result_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;

  const Sheet& sheet = result_or.value().workbook.sheet(0);
  const Cell* first = sheet.cell_at(0, 0);
  const Cell* second = sheet.cell_at(1, 0);
  ASSERT_NE(first, nullptr);
  ASSERT_NE(second, nullptr);
  ASSERT_EQ(first->phonetic_runs.size(), 1U);
  ASSERT_EQ(second->phonetic_runs.size(), 1U);
  EXPECT_EQ(first->phonetic_runs[0].eb, 2U);
  EXPECT_EQ(second->phonetic_runs[0].eb, 3U);
}

TEST(OoxmlPhoneticRoundTrip, WriterSkipsAnnotationOnUnannotatedCells) {
  // A Text cell without an annotation must round-trip with none; the
  // writer must not synthesise an <rPh> block out of thin air.
  Workbook src = Workbook::create();
  src.sheet(0).set_cell_value(0, 0, Value::text("plain"));

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  const std::vector<std::uint8_t> bytes = save_or.value();

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const Workbook& dst = result_or.value().workbook;

  const Cell* cell = dst.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->cached_value.as_text(), "plain");
  EXPECT_TRUE(cell->phonetic_runs.empty());
}

}  // namespace
}  // namespace formulon
