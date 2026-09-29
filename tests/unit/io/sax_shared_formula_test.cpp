//
// Unit tests for shared-formula resolution on the streaming SAX read
// path (`io::read_sheet_data_sax`). The SAX scanner now surfaces the
// `<f>` element's `t` / `si` / `ref` attributes, so a shared-formula
// group's followers (`<f t="shared" si="N"/>`, empty body) recover the
// master body shifted to their own cell — matching the DOM path. Before
// this, followers lost their formula entirely on the SAX path.

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "io/sheet_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

io::ByteSpan SpanOf(std::string_view s) {
  return io::ByteSpan{reinterpret_cast<const std::uint8_t*>(s.data()), s.size()};
}

// Drives the SAX cell reader over `sheet_xml` into sheet 0 of a fresh
// workbook and returns it.
Workbook LoadSax(std::string_view sheet_xml) {
  Workbook wb = Workbook::create();
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  SheetReadContext ctx;
  auto rs = read_sheet_data_sax(SpanOf(sheet_xml), 0U, wb, ctx, text_storage);
  EXPECT_TRUE(static_cast<bool>(rs)) << (rs ? "" : rs.error().message);
  return wb;
}

TEST(SaxSharedFormula, FollowersRecoverShiftedMasterBody) {
  // Master at B1 = A1*2 (si=0); followers at B2/B3 have empty bodies and
  // must resolve to the master body shifted down one / two rows.
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><v>10</v></c>"
      "<c r=\"B1\"><f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*2</f><v>20</v></c></row>\n"
      "<row r=\"2\"><c r=\"A2\"><v>11</v></c><c r=\"B2\"><f t=\"shared\" si=\"0\"/><v>22</v></c></row>\n"
      "<row r=\"3\"><c r=\"A3\"><v>12</v></c><c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>24</v></c></row>\n"
      "</sheetData></worksheet>\n";
  const Workbook wb = LoadSax(kXml);
  const Sheet& sheet = wb.sheet(0);

  const Cell* b1 = sheet.cell_at(0U, 1U);
  ASSERT_NE(b1, nullptr);
  EXPECT_EQ(b1->formula_text, "=A1*2");

  const Cell* b2 = sheet.cell_at(1U, 1U);
  ASSERT_NE(b2, nullptr);
  EXPECT_EQ(b2->formula_text, "=A2*2") << "shared follower must shift the master's relative refs";

  const Cell* b3 = sheet.cell_at(2U, 1U);
  ASSERT_NE(b3, nullptr);
  EXPECT_EQ(b3->formula_text, "=A3*2");
}

TEST(SaxSharedFormula, FollowerStripsStoragePrefixFromFutureFunctionMaster) {
  // The master body is the raw `<f>` text, which still carries Excel's
  // `_xlfn.` storage prefix; the general Call-node parse path only
  // recognises that prefix specially for LET/LAMBDA, so a follower must
  // have it stripped before shifting or it would keep the prefix glued to
  // the callee name in its own canonicalised formula_text.
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><v>3</v></c>"
      "<c r=\"B1\"><f t=\"shared\" ref=\"B1:B2\" si=\"0\">_xlfn.SEQUENCE(A1)</f></c></row>\n"
      "<row r=\"2\"><c r=\"A2\"><v>3</v></c><c r=\"B2\"><f t=\"shared\" si=\"0\"/></c></row>\n"
      "</sheetData></worksheet>\n";
  const Workbook wb = LoadSax(kXml);
  const Sheet& sheet = wb.sheet(0);

  const Cell* b1 = sheet.cell_at(0U, 1U);
  ASSERT_NE(b1, nullptr);
  EXPECT_EQ(b1->formula_text, "=SEQUENCE(A1)");

  const Cell* b2 = sheet.cell_at(1U, 1U);
  ASSERT_NE(b2, nullptr);
  EXPECT_EQ(b2->formula_text, "=SEQUENCE(A2)") << "shifted follower must drop the storage prefix, not keep it glued "
                                                  "to the callee name";
}

TEST(SaxSharedFormula, EntityEncodedAttributesResolveMasterAndFollower) {
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<sheetData>"
      "<row r=\"1\"><c r=\"B1\"><f t=\"shar&#101;d\" ref=\"B1&#58;B2\" si=\"&#49;\">A1*2</f></c></row>"
      "<row r=\"2\"><c r=\"A2\"><v>11</v></c><c r=\"B2\"><f t=\"shar&#101;d\" si=\"&#49;\"/></c></row>"
      "</sheetData></worksheet>";
  const Workbook wb = LoadSax(kXml);
  const Sheet& sheet = wb.sheet(0);
  const Cell* b1 = sheet.cell_at(0U, 1U);
  ASSERT_NE(b1, nullptr);
  EXPECT_EQ(b1->formula_text, "=A1*2");
  const Cell* b2 = sheet.cell_at(1U, 1U);
  ASSERT_NE(b2, nullptr);
  EXPECT_EQ(b2->formula_text, "=A2*2");
}

TEST(SaxSharedFormula, PlainFormulaUnaffected) {
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><f>1+2</f><v>3</v></c></row>\n"
      "</sheetData></worksheet>\n";
  const Workbook wb = LoadSax(kXml);
  const Cell* a1 = wb.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  EXPECT_EQ(a1->formula_text, "=1+2");
}

TEST(SaxSharedFormula, ArrayFormulaTreatedAsPlainBody) {
  // Array (CSE) formulas are read as plain formulas (body verbatim),
  // matching the DOM path.
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1\">SUM(B1:B3)</f><v>6</v></c></row>\n"
      "</sheetData></worksheet>\n";
  const Workbook wb = LoadSax(kXml);
  const Cell* a1 = wb.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  EXPECT_EQ(a1->formula_text, "=SUM(B1:B3)");
}

TEST(SaxSharedFormula, DataTableFormulaFallsBackToCachedValueAndIsCounted) {
  // `<f t="dataTable">` carries no body (its geometry lives in
  // attributes this reader does not decode), so the cell falls back to
  // its cached `<v>` -- matching the DOM path -- and the drop is
  // counted rather than silent.
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><f t=\"dataTable\" ref=\"A1:B2\" dt2D=\"1\" dtr=\"0\" r1=\"C1\" r2=\"C2\"/>"
      "<v>42</v></c></row>\n"
      "</sheetData></worksheet>\n";
  Workbook wb = Workbook::create();
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  SheetReadContext ctx;
  ReadDiagnostics diagnostics;
  auto rs = read_sheet_data_sax(SpanOf(kXml), 0U, wb, ctx, text_storage, &diagnostics);
  ASSERT_TRUE(static_cast<bool>(rs)) << (rs ? "" : rs.error().message);

  const Cell* a1 = wb.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  EXPECT_EQ(a1->formula_text, "");
  ASSERT_TRUE(a1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(a1->cached_value.as_number(), 42.0);
  EXPECT_EQ(diagnostics.skipped_feature_count, 1U);
}

// Mirrors `SheetReader.WideSparseStyledRowRejectedOverCellBudget`: the
// SAX path must enforce the same cell-materialisation budget as the DOM
// path, since either can be chosen for the same bytes depending on
// sheet size (see `read_sheet_data_sax`'s "identical" contract).
TEST(SaxSharedFormula, WideSparseStyledRowRejectedOverCellBudget) {
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\" s=\"1\"><v>1</v></c><c r=\"XFD1\" s=\"1\"><v>2</v></c></row>\n"
      "</sheetData></worksheet>\n";
  Workbook wb = Workbook::create();
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  SheetReadContext ctx;
  ReadDiagnostics diagnostics;
  auto rs = read_sheet_data_sax(SpanOf(kXml), 0U, wb, ctx, text_storage, &diagnostics, /*max_cell_bytes=*/1000U);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoFileTooLarge);
}

// Mirrors `SheetReader.TrulyEmptySparseCellsDoNotConsumeTheCellBudget`:
// a bare `<c>` with no value, style, or formula never materialises a
// `Cell`, so it must not be charged either.
TEST(SaxSharedFormula, TrulyEmptySparseCellsDoNotConsumeTheCellBudget) {
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"/><c r=\"XFD1\"/></row>\n"
      "</sheetData></worksheet>\n";
  Workbook wb = Workbook::create();
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  SheetReadContext ctx;
  auto rs =
      read_sheet_data_sax(SpanOf(kXml), 0U, wb, ctx, text_storage, /*diagnostics=*/nullptr, /*max_cell_bytes=*/1000U);
  ASSERT_TRUE(static_cast<bool>(rs)) << (rs ? "" : rs.error().message);
  EXPECT_EQ(wb.sheet(0).cell_count(), 0U);
}

TEST(SaxSharedFormula, UnknownSharedSiIsRejected) {
  // A follower referencing an unregistered `si` is corrupt, mirroring the
  // DOM path's diagnostic.
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"7\"/></c></row>\n"
      "</sheetData></worksheet>\n";
  Workbook wb = Workbook::create();
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  SheetReadContext ctx;
  auto rs = read_sheet_data_sax(SpanOf(kXml), 0U, wb, ctx, text_storage);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(SaxArrayAnchor, RejectsFullGridBeforeSpillWalk) {
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<sheetData><row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:XFD1048576\">1</f>"
      "</c></row></sheetData></worksheet>";
  Workbook wb = Workbook::create();
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  SheetReadContext ctx;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data_sax(SpanOf(kXml), 0U, wb, ctx, text_storage)));
  // Spill registration is the caller's job, run after SST resolution --
  // see `SheetReadContext::array_anchors`.
  auto rs = RegisterArraySpills(wb.sheet(0), ctx.array_anchors);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(0U, 0U), nullptr);
}

TEST(SaxArrayAnchor, RejectsCumulativeAnchorsOverSheetBudget) {
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<sheetData><row r=\"1\">"
      "<c r=\"A1\"><f t=\"array\" ref=\"A1:A600000\">1</f></c>"
      "<c r=\"B1\"><f t=\"array\" ref=\"B1:B600000\">1</f></c>"
      "</row></sheetData></worksheet>";
  Workbook wb = Workbook::create();
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  SheetReadContext ctx;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data_sax(SpanOf(kXml), 0U, wb, ctx, text_storage)));
  auto rs = RegisterArraySpills(wb.sheet(0), ctx.array_anchors);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_NE(rs.error().context.find("format=ooxml"), std::string::npos);
  EXPECT_NE(rs.error().context.find("anchor_row=0 anchor_col=1"), std::string::npos);
  EXPECT_NE(rs.error().context.find("used=600000 requested=600000 ceiling=1048576"), std::string::npos);
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(0U, 0U), nullptr);
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(0U, 1U), nullptr);
}

}  // namespace
}  // namespace io
}  // namespace formulon
