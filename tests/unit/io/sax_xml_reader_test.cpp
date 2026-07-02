// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for `formulon::io::scan_sheet_data` (the SAX-style sheet
// XML scanner). Each test feeds the scanner a small synthetic
// `<worksheet>` document and asserts the resulting callback stream.

#include "io/sax_xml_reader.h"

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace {

ByteSpan SpanOf(const std::string& s) {
  return ByteSpan{reinterpret_cast<const std::uint8_t*>(s.data()), s.size()};
}

/// Convenience: collect every `CellRecord` into a vector of
/// (row, col, t, formula, value, is_inline_string) tuples so the
/// assertions stay terse.
struct CapturedCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string t;
  std::string s;
  std::string formula;
  std::string value;
  bool is_inline_string = false;
  std::string phonetic;
};

struct Capture {
  std::vector<std::uint32_t> row_starts;
  std::vector<std::uint32_t> row_ends;
  std::vector<CapturedCell> cells;
};

Expected<void, Error> CaptureRowStart(void* ud, const RowRecord& row) {
  static_cast<Capture*>(ud)->row_starts.push_back(row.row_1based);
  return Expected<void, Error>::Ok();
}
Expected<void, Error> CaptureRowEnd(void* ud, std::uint32_t r) {
  static_cast<Capture*>(ud)->row_ends.push_back(r);
  return Expected<void, Error>::Ok();
}
Expected<void, Error> CaptureCell(void* ud, const CellRecord& rec) {
  static_cast<Capture*>(ud)->cells.push_back(CapturedCell{rec.row, rec.col, std::string(rec.t), std::string(rec.s),
                                                          std::string(rec.formula), std::string(rec.value),
                                                          rec.is_inline_string, std::string(rec.phonetic)});
  return Expected<void, Error>::Ok();
}

SheetSaxCallbacks MakeCallbacks(Capture* cap) {
  SheetSaxCallbacks cb;
  cb.user_data = cap;
  cb.on_row_start = &CaptureRowStart;
  cb.on_row_end = &CaptureRowEnd;
  cb.on_cell = &CaptureCell;
  return cb;
}

// ---------------------------------------------------------------------------
// 1. Empty / structural shapes.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, EmptySheetDataNoCallbacks) {
  const std::string xml = "<worksheet><sheetData/></worksheet>";
  Capture cap;
  auto rs = scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap));
  ASSERT_TRUE(static_cast<bool>(rs));
  EXPECT_TRUE(cap.row_starts.empty());
  EXPECT_TRUE(cap.row_ends.empty());
  EXPECT_TRUE(cap.cells.empty());
}

TEST(SaxXmlReader, AbsentSheetDataNoCallbacks) {
  const std::string xml = "<worksheet><dimension ref=\"A1\"/></worksheet>";
  Capture cap;
  auto rs = scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap));
  ASSERT_TRUE(static_cast<bool>(rs));
  EXPECT_TRUE(cap.cells.empty());
}

TEST(SaxXmlReader, NullByteSpanNoCallbacks) {
  Capture cap;
  auto rs = scan_sheet_data(ByteSpan{}, MakeCallbacks(&cap));
  ASSERT_TRUE(static_cast<bool>(rs));
  EXPECT_TRUE(cap.cells.empty());
}

// ---------------------------------------------------------------------------
// 2. Single cells of every recognised type.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, SingleCellNumber) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><v>42</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].row, 0U);
  EXPECT_EQ(cap.cells[0].col, 0U);
  EXPECT_EQ(cap.cells[0].t, "");
  EXPECT_EQ(cap.cells[0].value, "42");
  EXPECT_FALSE(cap.cells[0].is_inline_string);
}

TEST(SaxXmlReader, SingleCellSharedString) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"B1\" t=\"s\"><v>5</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].col, 1U);
  EXPECT_EQ(cap.cells[0].t, "s");
  EXPECT_EQ(cap.cells[0].value, "5");
}

TEST(SaxXmlReader, SingleCellInlineString) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>hello</t></is></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_TRUE(cap.cells[0].is_inline_string);
  EXPECT_EQ(cap.cells[0].t, "inlineStr");
  EXPECT_EQ(cap.cells[0].value, "hello");
}

TEST(SaxXmlReader, SingleCellInlineStringRichText) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\">"
      "<is><r><t>foo</t></r><r><t>bar</t></r></is>"
      "</c></row></sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_TRUE(cap.cells[0].is_inline_string);
  EXPECT_EQ(cap.cells[0].value, "foobar");
}

// <rPh> wraps a kana <t> that must NOT be folded into the surface text;
// otherwise large sheets (>= kSaxThresholdBytes) would silently corrupt
// inline-string cells whenever they carry a phonetic guide. The kana is
// captured separately in `phonetic` so PHONETIC works against SAX-
// materialised cells.
TEST(SaxXmlReader, InlineStringWithRPhSkipsKana) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\">"
      "<is><t>\xE5\xB1\xB1\xE7\x94\xB0</t>"
      "<rPh sb=\"0\" eb=\"2\"><t>\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0</t></rPh></is>"
      "</c></row></sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_TRUE(cap.cells[0].is_inline_string);
  EXPECT_EQ(cap.cells[0].value, "\xE5\xB1\xB1\xE7\x94\xB0");                 // 山田
  EXPECT_EQ(cap.cells[0].phonetic, "\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0");  // やまだ
}

// Multi-block <rPh> (one per kanji span) plus rich-text <r><t> runs in
// arbitrary order: surface concatenates only the <r><t> and bare <t>
// payloads; phonetic concatenates every <rPh><t> in document order.
TEST(SaxXmlReader, InlineStringWithMultipleRPhAndRichTextSkipsKana) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\">"
      "<is>"
      "<r><t>\xE5\xB1\xB1\xE7\x94\xB0</t></r>"
      "<rPh sb=\"0\" eb=\"2\"><t>\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0</t></rPh>"
      "<r><t>\xE5\xA4\xAA\xE9\x83\x8E</t></r>"
      "<rPh sb=\"2\" eb=\"4\"><t>\xE3\x81\x9F\xE3\x82\x8D\xE3\x81\x86</t></rPh>"
      "</is></c></row></sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].value, "\xE5\xB1\xB1\xE7\x94\xB0\xE5\xA4\xAA\xE9\x83\x8E");  // 山田太郎
  EXPECT_EQ(cap.cells[0].phonetic,
            "\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0\xE3\x81\x9F\xE3\x82\x8D\xE3\x81\x86");  // やまだたろう
}

// Inline string without <rPh>: phonetic stays empty.
TEST(SaxXmlReader, InlineStringWithoutRPhPhoneticEmpty) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>plain</t></is></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].value, "plain");
  EXPECT_TRUE(cap.cells[0].phonetic.empty());
}

TEST(SaxXmlReader, SingleCellBoolean) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"b\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].t, "b");
  EXPECT_EQ(cap.cells[0].value, "1");
}

TEST(SaxXmlReader, SingleCellError) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"e\"><v>#N/A</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].t, "e");
  EXPECT_EQ(cap.cells[0].value, "#N/A");
}

// ---------------------------------------------------------------------------
// 3. Multi-row, sparse-column, A1 decoding.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, MultipleRowsSparseColumns) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"C1\"><v>3</v></c></row>"
      "<row r=\"2\"><c r=\"B2\"><v>22</v></c></row>"
      "<row r=\"4\"><c r=\"E4\"><v>44</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 4U);
  EXPECT_EQ(cap.cells[0].col, 0U);
  EXPECT_EQ(cap.cells[1].col, 2U);
  EXPECT_EQ(cap.cells[2].col, 1U);
  EXPECT_EQ(cap.cells[2].row, 1U);
  EXPECT_EQ(cap.cells[3].col, 4U);
  EXPECT_EQ(cap.cells[3].row, 3U);
  EXPECT_EQ(cap.row_starts.size(), 3U);
  EXPECT_EQ(cap.row_ends.size(), 3U);
  EXPECT_EQ(cap.row_starts[0], 1U);
  EXPECT_EQ(cap.row_starts[2], 4U);
}

TEST(SaxXmlReader, ColumnLettersBeyondZ) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"AA1\"><v>1</v></c><c r=\"XFD1\"><v>2</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 2U);
  EXPECT_EQ(cap.cells[0].col, 26U);     // AA1 -> 26
  EXPECT_EQ(cap.cells[1].col, 16383U);  // XFD1 -> 16383 (Excel max)
}

// ---------------------------------------------------------------------------
// 4. Formula handling.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, FormulaWithCachedValue) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><f>A1+1</f><v>2</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].formula, "A1+1");
  EXPECT_EQ(cap.cells[0].value, "2");
}

TEST(SaxXmlReader, FormulaLeadingEqualsStripped) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><f>=SUM(B1:B3)</f></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].formula, "SUM(B1:B3)");
}

// ---------------------------------------------------------------------------
// 5. Self-closing variants.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, SelfClosingCellNoValue) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"/></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_TRUE(cap.cells[0].formula.empty());
  EXPECT_TRUE(cap.cells[0].value.empty());
}

TEST(SaxXmlReader, SelfClosingRow) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"/>"
      "<row r=\"2\"><c r=\"A2\"><v>5</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].row, 1U);
  EXPECT_EQ(cap.row_starts.size(), 2U);
  EXPECT_EQ(cap.row_ends.size(), 2U);
}

TEST(SaxXmlReader, EmptyValueElement) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><v></v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_TRUE(cap.cells[0].value.empty());
}

// ---------------------------------------------------------------------------
// 6. Attribute order and miscellaneous attributes.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, AttributesInArbitraryOrder) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c s=\"3\" t=\"n\" r=\"A1\"><v>9</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].s, "3");
  EXPECT_EQ(cap.cells[0].t, "n");
  EXPECT_EQ(cap.cells[0].col, 0U);
}

TEST(SaxXmlReader, SingleQuotedAttributes) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r='1'><c r='A1' t='n'><v>7</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].value, "7");
}

TEST(SaxXmlReader, IgnoresUnknownAttributes) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\" customHeight=\"true\"><c r=\"A1\" foo=\"bar\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].value, "1");
}

// ---------------------------------------------------------------------------
// 7. Whitespace and irrelevant siblings.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, WhitespaceBetweenTagsIgnored) {
  const std::string xml =
      "<worksheet>\n  <sheetData>\n"
      "    <row r=\"1\">\n      <c r=\"A1\"><v>1</v></c>\n    </row>\n"
      "  </sheetData>\n</worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].value, "1");
}

TEST(SaxXmlReader, SkipsUnknownSiblingsBeforeAndAfterSheetData) {
  const std::string xml =
      "<worksheet>"
      "<dimension ref=\"A1:B2\"/>"
      "<sheetViews><sheetView><selection activeCell=\"A1\"/></sheetView></sheetViews>"
      "<cols><col min=\"1\" max=\"1\" width=\"10\"/></cols>"
      "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>"
      "<mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells>"
      "<extLst><ext uri=\"foo\"/></extLst>"
      "</worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].value, "1");
}

TEST(SaxXmlReader, XmlDeclarationAndCommentSkipped) {
  const std::string xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
      "<!-- author: test -->"
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><v>5</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].value, "5");
}

// ---------------------------------------------------------------------------
// 8. Entity decoding.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, EntityDecodingInValue) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"str\"><v>a&amp;b&lt;c&gt;d&quot;e&apos;f</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].value, "a&b<c>d\"e'f");
}

TEST(SaxXmlReader, EntityDecodingInInlineString) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>x &amp; y</t></is></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].value, "x & y");
}

TEST(SaxXmlReader, NumericEntityDecodes) {
  // Numeric character references decode through the same UTF-8 encoder
  // that the attribute path uses (added alongside attribute entity
  // decoding so the two surfaces stay symmetric). `&#65;` -> 'A',
  // `&#x41;` -> 'A', and a non-ASCII reference exercises the multi-byte
  // path.
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>&#65;-&#x41;-&#x3042;</t></is></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  // Hiragana A (U+3042) encodes to UTF-8 0xE3 0x81 0x82.
  EXPECT_EQ(cap.cells[0].value, std::string("A-A-") + "\xE3\x81\x82");
}

TEST(SaxXmlReader, MalformedEntityPassesThroughVerbatim) {
  // Stray `&` without a matching `;`, and unknown named entities, are
  // emitted verbatim — the decoder is permissive on malformed input
  // because Excel-emitted payloads never trip these branches and
  // hand-rolled fixtures occasionally rely on the raw bytes.
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>raw&amp;and&unknown;</t></is></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].value, "raw&and&unknown;");
}

// ---------------------------------------------------------------------------
// 9. Malformed input -> kIoXmlParse, no crash.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, UnterminatedTagFails) {
  const std::string xml = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"";
  Capture cap;
  auto rs = scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap));
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoXmlParse);
}

TEST(SaxXmlReader, MalformedCellRefFails) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  auto rs = scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap));
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoXmlParse);
}

TEST(SaxXmlReader, MissingCellRefFails) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c><v>1</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  auto rs = scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap));
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoXmlParse);
}

TEST(SaxXmlReader, OversizedColumnLettersFail) {
  // ZZZZ is 4 letters; Excel max is XFD (3 letters).
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"ZZZZ1\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  auto rs = scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap));
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoXmlParse);
}

// ---------------------------------------------------------------------------
// 10. Stress / scaling.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, TenThousandCellsCallbackCount) {
  // Build 10000 cells across 100 rows, 100 cols. Confirm exactly that
  // many on_cell callbacks fire and that capture memory stays bounded
  // (we only retain the integer count, not the records).
  std::string xml = "<worksheet><sheetData>";
  constexpr std::uint32_t kRows = 100;
  constexpr std::uint32_t kCols = 100;
  for (std::uint32_t r = 1; r <= kRows; ++r) {
    xml.append("<row r=\"");
    xml.append(std::to_string(r));
    xml.append("\">");
    for (std::uint32_t c = 0; c < kCols; ++c) {
      xml.append("<c r=\"");
      // Column letters: simple A..Z then AA..CV (up to 100).
      if (c < 26U) {
        xml.push_back(static_cast<char>('A' + c));
      } else {
        xml.push_back(static_cast<char>('A' + (c / 26U) - 1U));
        xml.push_back(static_cast<char>('A' + (c % 26U)));
      }
      xml.append(std::to_string(r));
      xml.append("\"><v>");
      xml.append(std::to_string(r * 100 + c));
      xml.append("</v></c>");
    }
    xml.append("</row>");
  }
  xml.append("</sheetData></worksheet>");

  struct Counter {
    std::uint64_t cells = 0;
    std::uint64_t rows = 0;
  };
  Counter counter;
  SheetSaxCallbacks cb;
  cb.user_data = &counter;
  cb.on_row_start = +[](void* ud, const RowRecord&) -> Expected<void, Error> {
    ++static_cast<Counter*>(ud)->rows;
    return Expected<void, Error>::Ok();
  };
  cb.on_cell = +[](void* ud, const CellRecord&) -> Expected<void, Error> {
    ++static_cast<Counter*>(ud)->cells;
    return Expected<void, Error>::Ok();
  };
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), cb)));
  EXPECT_EQ(counter.rows, kRows);
  EXPECT_EQ(counter.cells, static_cast<std::uint64_t>(kRows) * kCols);
}

TEST(SaxXmlReader, CallbackErrorPropagates) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  SheetSaxCallbacks cb;
  cb.on_cell = +[](void* /*ud*/, const CellRecord&) -> Expected<void, Error> {
    return make_error(FormulonErrorCode::kInternalError, "test-injected", "context=test");
  };
  auto rs = scan_sheet_data(SpanOf(xml), cb);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kInternalError);
}

TEST(SaxXmlReader, NullCallbacksAcceptedForUntrackedFields) {
  // Only on_cell set: row callbacks are nullptr; should still scan.
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  std::uint64_t cells = 0;
  SheetSaxCallbacks cb;
  cb.user_data = &cells;
  cb.on_cell = +[](void* ud, const CellRecord&) -> Expected<void, Error> {
    ++(*static_cast<std::uint64_t*>(ud));
    return Expected<void, Error>::Ok();
  };
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), cb)));
  EXPECT_EQ(cells, 1U);
}

TEST(SaxXmlReader, AllCallbacksNullStillScansSuccessfully) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  SheetSaxCallbacks cb;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), cb)));
}

TEST(SaxXmlReader, RowWithoutAttributeStillFiresCallbacks) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row><c r=\"A1\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  EXPECT_EQ(cap.row_starts.size(), 1U);
  EXPECT_EQ(cap.row_starts[0], 0U);  // sentinel for "no r= attr"
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].col, 0U);
}

// ---------------------------------------------------------------------------
// Attribute entity decoding. The `<c>` / `<row>` attribute values must
// route through the same predefined-entity + numeric-character-reference
// decoder as `<v>` / `<t>` text content. Hand-rolled fixtures (and some
// non-Excel producers) emit values like `t="&quot;str&quot;"` or
// `r="A&#x31;"` that must canonicalise before A1 / type-token decoding.
// ---------------------------------------------------------------------------

TEST(SaxXmlReader, AttributeAmpEntityDecodes) {
  // `t="inlineStr"` written as `t="inline&amp;Str"` would not be a
  // recognised t-token after decoding ("inline&Str"), but the decoded
  // text must reach the record so a downstream check can diagnose it.
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inline&amp;Str\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].t, "inline&Str");
}

TEST(SaxXmlReader, AttributeAllPredefinedEntitiesDecode) {
  // Cell's s= attribute is a free-form string in our SAX layer (callers
  // parse the integer index later); use it as a transport for the five
  // predefined entities to assert end-to-end decoding without tripping
  // r= validation.
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" s=\"&amp;&lt;&gt;&quot;&apos;\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].s, "&<>\"'");
}

TEST(SaxXmlReader, AttributeNumericCharRefHexDecodes) {
  // `&#x31;` -> '1'. The cell reference `r="A&#x31;"` must canonicalise
  // to `A1` before A1 decoding runs, otherwise the parser rejects the
  // cell with a malformed-r= error.
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A&#x31;\"><v>42</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  auto rs = scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap));
  ASSERT_TRUE(static_cast<bool>(rs)) << "scan failed: " << rs.error().message;
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].row, 0U);
  EXPECT_EQ(cap.cells[0].col, 0U);
}

TEST(SaxXmlReader, AttributeNumericCharRefDecimalDecodes) {
  // `&#65;` -> 'A'. Pair with a literal digit so the resulting cell
  // reference is `A1`.
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"&#65;1\"><v>7</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].col, 0U);
  EXPECT_EQ(cap.cells[0].row, 0U);
}

TEST(SaxXmlReader, AttributeWithoutEntitiesIsZeroCopy) {
  // Sanity check: pure-ASCII attribute values do not invoke the decoder
  // (and the resulting view aliases the input). The observable behaviour
  // is identical either way; the test exists to lock down the common
  // path.
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"n\" s=\"3\"><v>9</v></c></row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 1U);
  EXPECT_EQ(cap.cells[0].t, "n");
  EXPECT_EQ(cap.cells[0].s, "3");
}

TEST(SaxXmlReader, MixedFormulaAndLiteralCellsInOneRow) {
  const std::string xml =
      "<worksheet><sheetData>"
      "<row r=\"1\">"
      "<c r=\"A1\"><v>1</v></c>"
      "<c r=\"B1\"><f>A1*2</f><v>2</v></c>"
      "<c r=\"C1\" t=\"inlineStr\"><is><t>x</t></is></c>"
      "</row>"
      "</sheetData></worksheet>";
  Capture cap;
  ASSERT_TRUE(static_cast<bool>(scan_sheet_data(SpanOf(xml), MakeCallbacks(&cap))));
  ASSERT_EQ(cap.cells.size(), 3U);
  EXPECT_EQ(cap.cells[0].value, "1");
  EXPECT_EQ(cap.cells[1].formula, "A1*2");
  EXPECT_EQ(cap.cells[1].value, "2");
  EXPECT_EQ(cap.cells[2].value, "x");
  EXPECT_TRUE(cap.cells[2].is_inline_string);
}

}  // namespace
}  // namespace io
}  // namespace formulon
