//
// DOM / SAX parity over malformed numeric attributes.
//
// `read_ooxml` chooses between `read_sheet_data` (pugixml DOM) and
// `read_sheet_data_sax` (streaming) by sheet size alone, so the two must
// turn any byte sequence — including values no producer should have
// written — into the same workbook and the same success / failure. The
// corpus parity gate covers well-formed input; the fixtures here pin the
// attribute values whose lexing used to differ between the paths, so a
// sheet crossing `kSaxThresholdBytes` cannot change what it means.
//
// Each test asserts the concrete decoded result as well as the equality,
// so the chosen disposition per attribute (`si=` hard error, `s=`
// degrades to 0, `<row r=>` drops the override) is pinned rather than
// merely shared.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/sheet_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

/// Everything a sheet read observably produces, flattened so the two
/// paths can be compared without keeping either workbook alive (text
/// payloads are copied out of the reader's storage).
struct CellSnapshot {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string formula;
  ValueKind kind = ValueKind::Blank;
  double number = 0.0;
  std::string text;
  std::uint32_t xf_index = 0;
};

struct LoadResult {
  bool ok = false;
  FormulonErrorCode code = FormulonErrorCode::kUnknownError;
  std::vector<CellSnapshot> cells;
  std::vector<RowLayout> rows;
};

void SnapshotSheet(const Sheet& sheet, LoadResult* out) {
  for (const auto& [row, cells] : sheet.rows()) {
    for (std::uint32_t col = 0; col < cells.size(); ++col) {
      const Cell& cell = cells[col];
      if (cell.formula_text.empty() && cell.cached_value.is_blank() && cell.xf_index == 0U) {
        continue;
      }
      CellSnapshot snap;
      snap.row = row;
      snap.col = col;
      snap.formula = cell.formula_text;
      snap.kind = cell.cached_value.kind();
      if (snap.kind == ValueKind::Number) {
        snap.number = cell.cached_value.as_number();
      } else if (snap.kind == ValueKind::Text) {
        snap.text = std::string(cell.cached_value.as_text());
      }
      snap.xf_index = cell.xf_index;
      out->cells.push_back(std::move(snap));
    }
  }
  out->rows = sheet.layout().row_overrides;
}

/// Reads `sheet_xml` through the DOM path. Row overrides live in
/// `read_sheet_view_and_layout` on this path, which the OOXML reader
/// calls over the same document, so both are run here.
LoadResult LoadDom(std::string_view sheet_xml) {
  LoadResult result;
  pugi::xml_document doc;
  if (!doc.load_buffer(sheet_xml.data(), sheet_xml.size())) {
    result.code = FormulonErrorCode::kIoXmlParse;
    return result;
  }
  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  auto rs = read_sheet_data(doc, 0U, wb, ctx, text_storage);
  if (!rs) {
    result.code = rs.error().code;
    return result;
  }
  auto vl = read_sheet_view_and_layout(doc, 0U, wb);
  if (!vl) {
    result.code = vl.error().code;
    return result;
  }
  result.ok = true;
  SnapshotSheet(wb.sheet(0), &result);
  return result;
}

/// Reads `sheet_xml` through the streaming path, which reports cells and
/// row overrides from the one scan.
LoadResult LoadSax(std::string_view sheet_xml) {
  LoadResult result;
  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  const ByteSpan span{reinterpret_cast<const std::uint8_t*>(sheet_xml.data()), sheet_xml.size()};
  auto rs = read_sheet_data_sax(span, 0U, wb, ctx, text_storage);
  if (!rs) {
    result.code = rs.error().code;
    return result;
  }
  result.ok = true;
  SnapshotSheet(wb.sheet(0), &result);
  return result;
}

::testing::AssertionResult SameOutcome(const LoadResult& dom, const LoadResult& sax) {
  if (dom.ok != sax.ok) {
    return ::testing::AssertionFailure() << "dom ok=" << dom.ok << " but sax ok=" << sax.ok;
  }
  if (!dom.ok) {
    return ::testing::AssertionSuccess();
  }
  if (dom.cells.size() != sax.cells.size()) {
    return ::testing::AssertionFailure() << "cell count differs: " << dom.cells.size() << " vs " << sax.cells.size();
  }
  for (std::size_t i = 0; i < dom.cells.size(); ++i) {
    const CellSnapshot& a = dom.cells[i];
    const CellSnapshot& b = sax.cells[i];
    if (a.row != b.row || a.col != b.col) {
      return ::testing::AssertionFailure() << "cell " << i << " address differs";
    }
    if (a.formula != b.formula) {
      return ::testing::AssertionFailure()
             << "cell " << i << " formula differs: '" << a.formula << "' vs '" << b.formula << "'";
    }
    if (a.kind != b.kind || a.number != b.number || a.text != b.text) {
      return ::testing::AssertionFailure() << "cell " << i << " value differs";
    }
    if (a.xf_index != b.xf_index) {
      return ::testing::AssertionFailure()
             << "cell " << i << " xf index differs: " << a.xf_index << " vs " << b.xf_index;
    }
  }
  if (dom.rows.size() != sax.rows.size()) {
    return ::testing::AssertionFailure() << "row override count differs: " << dom.rows.size() << " vs "
                                         << sax.rows.size();
  }
  for (std::size_t i = 0; i < dom.rows.size(); ++i) {
    const RowLayout& a = dom.rows[i];
    const RowLayout& b = sax.rows[i];
    if (a.row != b.row || a.height != b.height || a.hidden != b.hidden || a.outline_level != b.outline_level ||
        a.has_height != b.has_height) {
      return ::testing::AssertionFailure() << "row override " << i << " differs (row " << a.row << " vs " << b.row
                                           << ", hidden " << a.hidden << " vs " << b.hidden << ")";
    }
  }
  return ::testing::AssertionSuccess();
}

std::string WrapSheet(std::string_view body) {
  return std::string("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>")
      .append(body)
      .append("</sheetData></worksheet>");
}

TEST(SheetReaderLexingParity, WellFormedStyleAndRowOverrideAgree) {
  const std::string xml = WrapSheet("<row r=\"5\" ht=\"20\"><c r=\"A5\" s=\"5\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_EQ(dom.cells[0].xf_index, 5U);
  ASSERT_EQ(dom.rows.size(), 1U);
  EXPECT_EQ(dom.rows[0].row, 4U);
  EXPECT_DOUBLE_EQ(dom.rows[0].height, 20.0);
}

TEST(SheetReaderLexingParity, EntityDecodedFormulaAndRowAttributesAgree) {
  const std::string xml = WrapSheet(
      "<row r=\"&#53;\" ht=\"&#50;&#48;\" hidden=\"&#116;&#114;&#117;&#101;\" "
      "customHeight=\"&#116;&#114;&#117;&#101;\" outlineLevel=\"&#50;\">"
      "<c r=\"A5\" s=\"&#53;\"><f t=\"shar&#101;d\" si=\"&#49;\" ref=\"A5&#58;A6\">&#61;A4+1</f></c>"
      "</row>"
      "<row r=\"6\"><c r=\"A6\"><f t=\"shar&#101;d\" si=\"&#49;\"/></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  ASSERT_TRUE(sax.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 2U);
  const auto master = std::find_if(dom.cells.begin(), dom.cells.end(),
                                   [](const CellSnapshot& cell) { return cell.row == 4U && cell.col == 0U; });
  const auto follower = std::find_if(dom.cells.begin(), dom.cells.end(),
                                     [](const CellSnapshot& cell) { return cell.row == 5U && cell.col == 0U; });
  ASSERT_NE(master, dom.cells.end());
  ASSERT_NE(follower, dom.cells.end());
  EXPECT_EQ(master->formula, "=A4+1");
  EXPECT_EQ(follower->formula, "=A5+1");
  ASSERT_EQ(master->xf_index, 5U);
  ASSERT_EQ(dom.rows.size(), 1U);
  EXPECT_EQ(dom.rows[0].row, 4U);
  EXPECT_DOUBLE_EQ(dom.rows[0].height, 20.0);
  EXPECT_TRUE(dom.rows[0].hidden);
  EXPECT_EQ(dom.rows[0].outline_level, 2U);
}

TEST(SheetReaderLexingParity, EntityDecodedMalformedValuesKeepDispositions) {
  const std::string xml = WrapSheet("<row r=\"&#53;x\" ht=\"20\"><c r=\"A5\" s=\"&#53;x\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_EQ(dom.cells[0].xf_index, 0U);
  EXPECT_TRUE(dom.rows.empty());
}

TEST(SheetReaderLexingParity, EntityDecodedOutlinePrefixAgrees) {
  const std::string xml = WrapSheet("<row r=\"1\" outlineLevel=\"&#53;x\"><c r=\"A1\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  ASSERT_TRUE(sax.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.rows.size(), 1U);
  EXPECT_EQ(dom.rows[0].row, 0U);
  EXPECT_EQ(dom.rows[0].outline_level, 5U);
}

TEST(SheetReaderLexingParity, LiteralAttributeWhitespaceAgrees) {
  const std::string xml =
      WrapSheet("<row r=\"5\" ht=\"20\t\" hidden=\"true\r\n\"><c r=\"A5\" s=\"5\r\n\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_EQ(dom.cells[0].xf_index, 0U);
  ASSERT_EQ(dom.rows.size(), 1U);
  EXPECT_DOUBLE_EQ(dom.rows[0].height, 20.0);
  EXPECT_FALSE(dom.rows[0].hidden);
}

TEST(SheetReaderLexingParity, LeadingSpaceInStyleIndexDegradesToDefault) {
  // Used to read as xf=5 under the DOM path (pugixml `as_uint` skips
  // leading whitespace) and xf=0 under the streaming path, so the same
  // cell was formatted differently depending on the sheet's size.
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\" s=\" 5\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_EQ(dom.cells[0].xf_index, 0U);
}

TEST(SheetReaderLexingParity, TrailingGarbageInStyleIndexDegradesToDefault) {
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\" s=\"5x\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_EQ(dom.cells[0].xf_index, 0U);
}

TEST(SheetReaderLexingParity, OverflowingStyleIndexDegradesToDefault) {
  // Past UINT32_MAX. The streaming path used to wrap this into a small
  // plausible index; both paths now fall back to the schema default.
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\" s=\"4294967296\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_EQ(dom.cells[0].xf_index, 0U);
}

TEST(SheetReaderLexingParity, MalformedSharedFormulaIndexFailsBothPaths) {
  // Used to succeed under the DOM path (`as_llong` stops at the 'x' and
  // yields 1) and fail the whole sheet under the streaming path.
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"1x\">1+1</f><v>2</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  EXPECT_FALSE(dom.ok);
  EXPECT_FALSE(sax.ok);
  EXPECT_EQ(dom.code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_EQ(sax.code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_TRUE(SameOutcome(dom, sax));
}

TEST(SheetReaderLexingParity, OverflowingSharedFormulaIndexFailsBothPaths) {
  const std::string xml =
      WrapSheet("<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"4294967296\">1+1</f><v>2</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  EXPECT_FALSE(dom.ok);
  EXPECT_FALSE(sax.ok);
  EXPECT_EQ(dom.code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_EQ(sax.code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(SheetReaderLexingParity, MalformedRowNumberDropsTheOverride) {
  // Used to drop the override under the DOM path and apply it to row 4
  // (the `5` prefix, minus one) under the streaming path.
  const std::string xml = WrapSheet("<row r=\"5x\" ht=\"20\"><c r=\"A5\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  EXPECT_TRUE(dom.rows.empty());
  // The cell itself is addressed by `<c r=>`, not by the row number, so
  // it still lands at A5 on both paths.
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_EQ(dom.cells[0].row, 4U);
}

TEST(SheetReaderLexingParity, LeadingSpaceInRowNumberDropsTheOverride) {
  const std::string xml = WrapSheet("<row r=\" 5\" ht=\"20\"><c r=\"A5\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  EXPECT_TRUE(dom.rows.empty());
}

TEST(SheetReaderLexingParity, UppercaseHiddenBooleanAgrees) {
  // `hidden` is an XSD boolean, whose alphabetic forms are
  // case-insensitive. The DOM path already accepted `TRUE`; the
  // streaming path only matched the lowercase spelling.
  const std::string xml = WrapSheet("<row r=\"3\" hidden=\"TRUE\"><c r=\"A3\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.rows.size(), 1U);
  EXPECT_TRUE(dom.rows[0].hidden);
}

TEST(SheetReaderLexingParity, NonFiniteCellValueFailsBothPaths) {
  // A magnitude IEEE 754 cannot hold is rejected at load rather than
  // stored as +inf and re-emitted as `#NUM!` by the next save.
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\"><v>1e999</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  EXPECT_FALSE(dom.ok);
  EXPECT_FALSE(sax.ok);
  EXPECT_EQ(dom.code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_EQ(sax.code, FormulonErrorCode::kIoSheetCorrupt);
}

}  // namespace
}  // namespace io
}  // namespace formulon
