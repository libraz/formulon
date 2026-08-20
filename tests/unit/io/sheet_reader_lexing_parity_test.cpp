//
// DOM / SAX parity over attribute lexing and document node shape.
//
// `read_ooxml` chooses between `read_sheet_data` (pugixml DOM) and
// `read_sheet_data_sax` (streaming) by sheet size alone, so the two must
// turn any byte sequence — including values no producer should have
// written — into the same workbook and the same success / failure. The
// corpus parity gate covers well-formed input; the fixtures here pin the
// attribute values whose lexing used to differ between the paths, plus the
// document shapes -- quoted attribute delimiters, CDATA / comments /
// processing instructions in a text body, stray character data, the root
// element's name -- where the two used to disagree on the node tree. So a
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
#include "io/xml_utils.h"
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
/// A one-sheet workbook with no style table.
///
/// These fixtures feed a bare `<sheetData>` fragment to the sheet readers
/// with no `xl/styles.xml` alongside it, which is the case `ApplyParsedCell`
/// exempts from its dangling-`s=` clamp. Both factories seed the
/// Excel-canonical minimum style table, so the host has to be emptied the
/// way the OOXML reader empties it; leaving the seed in place would clamp
/// `s="5"` to 0 and pin the wrong disposition here.
Workbook MakeStylelessHost() {
  Workbook wb = Workbook::create_empty();
  wb.set_styles(StylesTable{});
  wb.add_sheet("Sheet1");
  return wb;
}

LoadResult LoadDom(std::string_view sheet_xml) {
  LoadResult result;
  pugi::xml_document doc;
  // `load_xml_buffer` is the loader every part reader goes through, and it
  // adds `parse_ws_pcdata_single` to the defaults. Parsing these fixtures
  // any other way would compare the streaming path against a parser
  // configuration that never runs.
  const std::vector<std::uint8_t> bytes(sheet_xml.begin(), sheet_xml.end());
  if (!load_xml_buffer(doc, bytes, "sheet_reader_parity", "sheet.xml")) {
    result.code = FormulonErrorCode::kIoXmlParse;
    return result;
  }
  Workbook wb = MakeStylelessHost();
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
  Workbook wb = MakeStylelessHost();
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

TEST(SheetReaderLexingParity, NonFiniteRowHeightDropsTheHeightOnBothPaths) {
  // `strtod` turns these four spellings into +inf, a quiet NaN, 16, and 0
  // respectively. An infinite or NaN track height reaches the paginator,
  // which then breaks a page at every row or none at all, so a height
  // outside the shared lexical space is treated as absent instead.
  for (const char* spelling : {"INF", "NaN", "0x10", "abc"}) {
    const std::string xml =
        WrapSheet(std::string("<row r=\"1\" ht=\"").append(spelling).append("\"><c r=\"A1\"><v>1</v></c></row>"));
    const LoadResult dom = LoadDom(xml);
    const LoadResult sax = LoadSax(xml);
    ASSERT_TRUE(dom.ok) << spelling;
    ASSERT_TRUE(sax.ok) << spelling;
    EXPECT_TRUE(SameOutcome(dom, sax)) << spelling;
    // The row carried nothing but the bad height, so it contributes no
    // override at all rather than an all-defaults one.
    EXPECT_TRUE(dom.rows.empty()) << spelling;
  }
}

TEST(SheetReaderLexingParity, OverflowingRowHeightDropsTheHeightOnBothPaths) {
  const std::string xml = WrapSheet("<row r=\"1\" ht=\"1e999\"><c r=\"A1\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  EXPECT_TRUE(dom.rows.empty());
}

TEST(SheetReaderLexingParity, NegativeRowHeightDropsTheHeightOnBothPaths) {
  // A track cannot be shorter than nothing; a negative height subtracts
  // from the running page extent instead of adding to it.
  const std::string xml = WrapSheet("<row r=\"1\" ht=\"-20\"><c r=\"A1\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  EXPECT_TRUE(dom.rows.empty());
}

TEST(SheetReaderLexingParity, RowWithBadHeightKeepsItsOtherOverridesOnBothPaths) {
  // Dropping the height must not drop the row: `hidden` still applies,
  // and the entry reports no height so the paginator uses the default.
  const std::string xml = WrapSheet("<row r=\"3\" ht=\"NaN\" hidden=\"1\"><c r=\"A3\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.rows.size(), 1U);
  EXPECT_EQ(dom.rows[0].row, 2U);
  EXPECT_TRUE(dom.rows[0].hidden);
  EXPECT_FALSE(dom.rows[0].has_height);
}

TEST(SheetReaderLexingParity, ZeroRowHeightIsKeptOnBothPaths) {
  // Zero is a legitimate height Excel writes for a collapsed row, so the
  // non-negative gate must not swallow it.
  const std::string xml = WrapSheet("<row r=\"1\" ht=\"0\"><c r=\"A1\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.rows.size(), 1U);
  EXPECT_TRUE(dom.rows[0].has_height);
  EXPECT_DOUBLE_EQ(dom.rows[0].height, 0.0);
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

// ---------------------------------------------------------------------------
// Node shape. The cases above vary the *bytes* of an attribute value; these
// vary the *structure* of the document, which the lexing fixtures never
// reach. XML 1.0 permits a literal `>` inside an attribute value (only `<`
// and `&` are forbidden there), and permits CDATA / comments / processing
// instructions wherever character data may appear.
// ---------------------------------------------------------------------------

/// Wraps `body` in a worksheet whose root carries `extra_root_markup`
/// before `<sheetData>`, so a start tag ahead of the cell data can be
/// varied independently of the cells themselves.
std::string WrapSheetWithPrefix(std::string_view extra_root_markup, std::string_view body) {
  return std::string("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">")
      .append(extra_root_markup)
      .append("<sheetData>")
      .append(body)
      .append("</sheetData></worksheet>");
}

TEST(SheetReaderNodeShapeParity, GreaterThanInAttributeBeforeSheetDataKeepsTheCells) {
  // A literal `>` is legal in an attribute value. Treating it as the end of
  // the start tag makes the rest of `<sheetPr .../>` look like character
  // data, and the recovery scan then swallows `<sheetData>` whole -- so the
  // sheet reads as empty and still reports success.
  const std::string xml =
      WrapSheetWithPrefix("<sheetPr codeName=\"a&gt;b\"/>", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>");
  const std::string literal =
      WrapSheetWithPrefix("<sheetPr codeName=\"a>b\"/>", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>");
  const LoadResult dom = LoadDom(literal);
  const LoadResult sax = LoadSax(literal);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  // The escaped spelling is the same document; both must agree with it too.
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_EQ(LoadDom(xml).cells.size(), 1U);
  EXPECT_DOUBLE_EQ(dom.cells[0].number, 1.0);
}

TEST(SheetReaderNodeShapeParity, GreaterThanInAttributeInsideSheetDataKeepsTheCells) {
  const std::string xml = WrapSheet("<row r=\"1\" spans=\"1:2>\"><c r=\"A1\"><v>7</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_DOUBLE_EQ(dom.cells[0].number, 7.0);
}

TEST(SheetReaderNodeShapeParity, GreaterThanInSingleQuotedAttributeKeepsTheCells) {
  // Both quote characters delimit an attribute value, so tracking only the
  // double quote would leave this shape truncating the tag.
  const std::string xml = WrapSheet("<row r='1' spans='1:2>'><c r='A1'><v>9</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_DOUBLE_EQ(dom.cells[0].number, 9.0);
}

TEST(SheetReaderNodeShapeParity, SlashInsideAttributeDoesNotSelfCloseTheTag) {
  // `/` only closes a tag as `/>` outside a quoted value; inside one it is
  // an ordinary character, as in a path or URL.
  const std::string xml = WrapSheet("<row r=\"1\" spans=\"a/>b\"><c r=\"A1\"><v>3</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_DOUBLE_EQ(dom.cells[0].number, 3.0);
}

TEST(SheetReaderNodeShapeParity, UnterminatedAttributeQuoteFailsBothPaths) {
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\" s=\"0><v>1</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  EXPECT_FALSE(dom.ok);
  EXPECT_FALSE(sax.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
}

TEST(SheetReaderNodeShapeParity, CdataInsideValueAgrees) {
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\"><v><![CDATA[42]]></v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_DOUBLE_EQ(dom.cells[0].number, 42.0);
}

TEST(SheetReaderNodeShapeParity, CommentAndProcessingInstructionInsideValueAgree) {
  // `parse_default` keeps neither node type, so the surviving character
  // data is the whole value.
  for (const char* markup : {"<!--note-->", "<?php echo 1;?>"}) {
    const std::string xml =
        WrapSheet(std::string("<row r=\"1\"><c r=\"A1\"><v>").append(markup).append("42</v></c></row>"));
    const LoadResult dom = LoadDom(xml);
    const LoadResult sax = LoadSax(xml);
    ASSERT_TRUE(dom.ok) << markup;
    EXPECT_TRUE(SameOutcome(dom, sax)) << markup;
    ASSERT_EQ(dom.cells.size(), 1U) << markup;
    EXPECT_DOUBLE_EQ(dom.cells[0].number, 42.0) << markup;
  }
}

TEST(SheetReaderNodeShapeParity, CdataInsideInlineStringAgrees) {
  const std::string xml =
      WrapSheet("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t><![CDATA[a<b&c]]></t></is></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  // CDATA content is not entity-decoded and may carry XML-critical bytes.
  EXPECT_EQ(dom.cells[0].text, "a<b&c");
}

TEST(SheetReaderNodeShapeParity, CommentAndProcessingInstructionInsideInlineStringAgree) {
  for (const char* markup : {"<!--note-->", "<?php echo 1;?>"}) {
    const std::string xml = WrapSheet(std::string("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>")
                                          .append(markup)
                                          .append("hello</t></is></c></row>"));
    const LoadResult dom = LoadDom(xml);
    const LoadResult sax = LoadSax(xml);
    ASSERT_TRUE(dom.ok) << markup;
    EXPECT_TRUE(SameOutcome(dom, sax)) << markup;
    ASSERT_EQ(dom.cells.size(), 1U) << markup;
    EXPECT_EQ(dom.cells[0].text, "hello") << markup;
  }
}

TEST(SheetReaderNodeShapeParity, TextRunEndsAtTheFirstNonCharacterNode) {
  // pugixml exposes the first PCDATA-or-CDATA child, and a comment, a
  // processing instruction or a CDATA section all end the run that precedes
  // them. So the contribution here is "4", not "42" -- the streaming path
  // has to reproduce that, not the concatenation the spec might suggest.
  for (const char* markup : {"<!--x-->", "<?pi x?>", "<![CDATA[2]]>"}) {
    const std::string xml =
        WrapSheet(std::string("<row r=\"1\"><c r=\"A1\"><v>4").append(markup).append("2</v></c></row>"));
    const LoadResult dom = LoadDom(xml);
    const LoadResult sax = LoadSax(xml);
    ASSERT_TRUE(dom.ok) << markup;
    EXPECT_TRUE(SameOutcome(dom, sax)) << markup;
    ASSERT_EQ(dom.cells.size(), 1U) << markup;
    EXPECT_DOUBLE_EQ(dom.cells[0].number, 4.0) << markup;
  }
}

TEST(SheetReaderNodeShapeParity, WhitespaceOnlyRunBeforeCdataIsDropped) {
  // Whitespace-only character data forms no node under `parse_default`, so
  // the CDATA section that follows it is still the first one.
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\"><v>\n  <![CDATA[42]]>\n</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_DOUBLE_EQ(dom.cells[0].number, 42.0);
}

TEST(SheetReaderNodeShapeParity, ChildElementInsideValueAgrees) {
  // An element child is not part of the OOXML vocabulary here, but pugixml
  // accepts it and still reports the first character-data node, so the
  // streaming path may not reject the sheet over it.
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\"><v><b/>42</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_DOUBLE_EQ(dom.cells[0].number, 42.0);
}

TEST(SheetReaderNodeShapeParity, WhitespaceOnlyRunSurvivesOnlyAsTheSoleChild) {
  // `parse_ws_pcdata_single` keeps a whitespace-only run when it is the
  // element's only child and drops it otherwise, so an inline string reads
  // as a space in the first shape and as empty in the second.
  const std::string alone = WrapSheet("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t> </t></is></c></row>");
  const LoadResult alone_dom = LoadDom(alone);
  ASSERT_TRUE(alone_dom.ok);
  EXPECT_TRUE(SameOutcome(alone_dom, LoadSax(alone)));
  ASSERT_EQ(alone_dom.cells.size(), 1U);
  EXPECT_EQ(alone_dom.cells[0].text, " ");

  const std::string with_sibling =
      WrapSheet("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t> <!--x--></t></is></c></row>");
  const LoadResult sibling_dom = LoadDom(with_sibling);
  ASSERT_TRUE(sibling_dom.ok);
  EXPECT_TRUE(SameOutcome(sibling_dom, LoadSax(with_sibling)));
  ASSERT_EQ(sibling_dom.cells.size(), 1U);
  EXPECT_EQ(sibling_dom.cells[0].text, "");
}

TEST(SheetReaderNodeShapeParity, CdataTakesEolConversionButNotEntityDecoding) {
  // Inside a CDATA section `&amp;` is five characters, while the
  // end-of-line conversion still applies to the raw bytes.
  const std::string xml =
      WrapSheet("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t><![CDATA[&amp;\r\nx]]></t></is></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_EQ(dom.cells[0].text, "&amp;\nx");
}

TEST(SheetReaderNodeShapeParity, CdataInsideFormulaAgrees) {
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\"><f><![CDATA[1+1]]></f><v>2</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_EQ(dom.cells[0].formula, "=1+1");
}

TEST(SheetReaderNodeShapeParity, UnterminatedCdataFailsBothPaths) {
  const std::string xml = WrapSheet("<row r=\"1\"><c r=\"A1\"><v><![CDATA[42</v></c></row>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  EXPECT_FALSE(dom.ok);
  EXPECT_FALSE(sax.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
}

TEST(SheetReaderNodeShapeParity, GreaterThanInAttributeAfterSheetDataKeepsTheCells) {
  // The same truncation one element later: here it is `</sheetData>` that
  // has already closed, so the damage lands on the metadata sibling rather
  // than on the cells -- both paths must still read the row.
  const std::string xml =
      std::string("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>")
          .append("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>")
          .append("</sheetData><mergeCells count=\"1>\"><mergeCell ref=\"B1:C1\"/></mergeCells></worksheet>");
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_DOUBLE_EQ(dom.cells[0].number, 1.0);
}

TEST(SheetReaderNodeShapeParity, GreaterThanInRootAttributeKeepsTheCells) {
  const std::string xml =
      "<worksheet codeName=\"a>b\"><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>";
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  ASSERT_TRUE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  ASSERT_EQ(dom.cells.size(), 1U);
  EXPECT_DOUBLE_EQ(dom.cells[0].number, 1.0);
}

TEST(SheetReaderNodeShapeParity, StrayCharacterDataBetweenChildrenIsIgnored) {
  // The DOM path reaches `<row>`, `<c>` and `<v>` by name, so a text node
  // sitting between them contributes nothing there. Rejecting it on the
  // streaming path would fail a workbook the DOM path opens.
  const struct {
    const char* label;
    const char* body;
  } kCases[] = {
      {"under sheetData", "junk<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"},
      {"under row", "<row r=\"1\">junk<c r=\"A1\"><v>1</v></c></row>"},
      {"under c", "<row r=\"1\"><c r=\"A1\">junk<v>1</v></c></row>"},
  };
  for (const auto& c : kCases) {
    const std::string xml = WrapSheet(c.body);
    const LoadResult dom = LoadDom(xml);
    const LoadResult sax = LoadSax(xml);
    ASSERT_TRUE(dom.ok) << c.label;
    EXPECT_TRUE(SameOutcome(dom, sax)) << c.label;
    ASSERT_EQ(dom.cells.size(), 1U) << c.label;
    EXPECT_DOUBLE_EQ(dom.cells[0].number, 1.0) << c.label;
  }
}

TEST(SheetReaderNodeShapeParity, CellsUnderAForeignRootAreNotRead) {
  // `read_sheet_data` resolves the cells through the document's
  // `<worksheet>` child, so a differently named root is a corrupt sheet
  // there. The streaming path used to stream the grid out of it anyway.
  const std::string xml = "<foo><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></foo>";
  const LoadResult dom = LoadDom(xml);
  const LoadResult sax = LoadSax(xml);
  EXPECT_FALSE(dom.ok);
  EXPECT_TRUE(SameOutcome(dom, sax));
  EXPECT_EQ(dom.code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_EQ(sax.code, FormulonErrorCode::kIoSheetCorrupt);
}

}  // namespace
}  // namespace io
}  // namespace formulon
