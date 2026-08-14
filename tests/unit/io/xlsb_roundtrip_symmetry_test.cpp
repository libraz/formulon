//
// Cross-format read symmetry for the binary XLSB path.
//
// `tests/unit/io/xlsb_fidelity_test.cpp` (bind-a) already checks the `.xlsb`
// reader against literal expected values and covers the `write_xlsb ->
// read_xlsb` round-trip for formula text and cell values. This file adds the
// complementary angle that no existing test covers: the SAME real workbook,
// authored once by Excel and exported to both `.xlsb` and `.xlsx`, must yield
// an equivalent in-memory model regardless of which reader parsed it. A
// format-specific reader divergence (a record the XLSB reader drops but the
// OOXML reader keeps, or vice versa) surfaces as a model mismatch here even
// when each reader passes its own literal-expectation fidelity suite.
//
// XLSB is a binary record stream, so the pugixml attribute-set helpers do not
// apply; the comparison is at the `Workbook` model level.

#include <cstdint>
#include <string>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/ooxml_reader.h"
#include "io/styles_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::string FixturePath(const char* name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
}

// Reads the shared fixture from both formats. Returns false (with a gtest
// failure) if either read fails; otherwise fills the two out-workbooks.
::testing::AssertionResult LoadBothFormats(Workbook* xlsb_out, Workbook* xlsx_out) {
  const std::vector<std::uint8_t> xlsb_bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsb"));
  const std::vector<std::uint8_t> xlsx_bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsx"));
  if (xlsb_bytes.empty() || xlsx_bytes.empty()) {
    return ::testing::AssertionFailure() << "fixture bytes empty";
  }
  auto xb = io::xlsb::read_xlsb(test::span_of(xlsb_bytes));
  if (!xb) {
    return ::testing::AssertionFailure() << "read_xlsb failed: " << xb.error().message;
  }
  auto xx = io::read_ooxml(test::span_of(xlsx_bytes));
  if (!xx) {
    return ::testing::AssertionFailure() << "read_ooxml failed: " << xx.error().message;
  }
  *xlsb_out = std::move(xb.value().workbook);
  *xlsx_out = std::move(xx.value().workbook);
  return ::testing::AssertionSuccess();
}

TEST(XlsbCrossFormatSymmetry, SheetStructureMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  ASSERT_EQ(xlsb.sheet_count(), xlsx.sheet_count());
  for (std::size_t i = 0; i < xlsb.sheet_count(); ++i) {
    EXPECT_EQ(xlsb.sheet(i).name(), xlsx.sheet(i).name()) << "sheet index " << i;
  }
}

TEST(XlsbCrossFormatSymmetry, DataSheetValuesMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const Sheet& sb = xlsb.sheet(0);
  const Sheet& sx = xlsx.sheet(0);
  // A1:A3 are text keys; B1:B3 are the numeric column.
  for (std::uint32_t r = 0; r < 3; ++r) {
    const Cell* ab = sb.cell_at(r, 0);
    const Cell* ax = sx.cell_at(r, 0);
    ASSERT_NE(ab, nullptr);
    ASSERT_NE(ax, nullptr);
    ASSERT_TRUE(ab->cached_value.is_text());
    ASSERT_TRUE(ax->cached_value.is_text());
    EXPECT_EQ(ab->cached_value.as_text(), ax->cached_value.as_text()) << "A" << (r + 1);

    const Cell* bb = sb.cell_at(r, 1);
    const Cell* bx = sx.cell_at(r, 1);
    ASSERT_NE(bb, nullptr);
    ASSERT_NE(bx, nullptr);
    ASSERT_TRUE(bb->cached_value.is_number());
    ASSERT_TRUE(bx->cached_value.is_number());
    EXPECT_DOUBLE_EQ(bb->cached_value.as_number(), bx->cached_value.as_number()) << "B" << (r + 1);
  }
}

TEST(XlsbCrossFormatSymmetry, FormulaTextMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  // F1 holds an XLOOKUP; both readers must restore identical formula text
  // (no `_xlfn.` prefix drift between the binary and OOXML paths).
  const Cell* fb = xlsb.sheet(0).cell_at(0, 5);
  const Cell* fx = xlsx.sheet(0).cell_at(0, 5);
  ASSERT_NE(fb, nullptr);
  ASSERT_NE(fx, nullptr);
  EXPECT_EQ(fb->formula_text, fx->formula_text);
  EXPECT_FALSE(fb->formula_text.empty());
}

TEST(XlsbCrossFormatSymmetry, StyleIndexMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  // D3 is a bold-red styled cell. Both readers must resolve it to the same
  // style (xf) index -- a cross-format check on the styles.bin vs styles.xml
  // parse producing equivalent style tables.
  const Cell* db = xlsb.sheet(0).cell_at(2, 3);
  const Cell* dx = xlsx.sheet(0).cell_at(2, 3);
  ASSERT_NE(db, nullptr);
  ASSERT_NE(dx, nullptr);
  EXPECT_EQ(db->xf_index, dx->xf_index);
  EXPECT_NE(db->xf_index, 0U) << "D3 should carry a non-default style";
}

TEST(XlsbCrossFormatSymmetry, DefinedNamesMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const std::vector<io::DefinedName>& sb = xlsb.defined_names();
  const std::vector<io::DefinedName>& sx = xlsx.defined_names();
  ASSERT_EQ(sb.size(), sx.size());
  // Field-level, not just count: the XLSB reader must fill the same
  // `io::DefinedName` field set the OOXML reader does (name, formula,
  // scope, hidden, comment) for the same source workbook, not merely
  // produce the same number of entries.
  for (std::size_t i = 0; i < sb.size(); ++i) {
    EXPECT_EQ(sb[i].name, sx[i].name) << "index " << i;
    EXPECT_EQ(sb[i].formula, sx[i].formula) << "index " << i;
    EXPECT_EQ(sb[i].local_sheet_id, sx[i].local_sheet_id) << "index " << i;
    EXPECT_EQ(sb[i].hidden, sx[i].hidden) << "index " << i;
    EXPECT_EQ(sb[i].comment, sx[i].comment) << "index " << i;
  }
}

// Resolves the numFmtId for cell (row, col) through `wb`'s style table, or
// SIZE_MAX-style 0xFFFFFFFF when the cell's xf index dangles past the table
// (which is exactly the failure mode a bare index-equality check would miss).
std::uint32_t ResolvedNumFmtId(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Cell* c = wb.sheet(0).cell_at(row, col);
  if (c == nullptr) {
    return 0xFFFFFFFFU;
  }
  const io::StylesTable& st = wb.styles();
  if (c->xf_index >= st.cell_xfs.size()) {
    return 0xFFFFFFFFU;  // dangling index -> style table did not round-trip
  }
  return st.cell_xfs[c->xf_index].num_fmt_id;
}

// Regression for the writer defect the cross-format check surfaced. Two halves,
// both required for a styled cell to survive `write_xlsb -> read_xlsb`:
//   1. the cell header must carry the 24-bit iStyleRef (was hardcoded to 0);
//   2. the workbook must declare the styles relationship so the reader can find
//      the (passthrough) styles.bin -- otherwise the index dangles against an
//      empty table.
// Asserting the *resolved* numFmtId (not just index equality) exercises both:
// D1 = yyyy/mm/dd (custom 179), D2 = #,##0.00 (built-in 4), D5 = 0.0% (custom
// 180). D3 keeps its non-default font xf.
TEST(XlsbWriteReadSymmetry, CellStyleAndNumberFormatSurviveRoundTrip) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsb"));
  ASSERT_FALSE(bytes.empty());
  auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << "read_xlsb failed: " << loaded.error().message;
  const Workbook& before = loaded.value().workbook;

  // Baseline: the fixture resolves the expected number formats.
  ASSERT_EQ(ResolvedNumFmtId(before, 0, 3), 179U);  // D1 yyyy/mm/dd
  ASSERT_EQ(ResolvedNumFmtId(before, 1, 3), 4U);    // D2 #,##0.00
  ASSERT_EQ(ResolvedNumFmtId(before, 4, 3), 180U);  // D5 0.0%
  const Cell* d3_before = before.sheet(0).cell_at(2, 3);
  ASSERT_NE(d3_before, nullptr);
  ASSERT_NE(d3_before->xf_index, 0U) << "fixture D3 should carry a non-default style";

  auto saved = io::xlsb::write_xlsb(before);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "reload failed: " << reloaded.error().message;
  const Workbook& after = reloaded.value().workbook;

  // Number formats still resolve after the round-trip (index + style table).
  EXPECT_EQ(ResolvedNumFmtId(after, 0, 3), 179U);
  EXPECT_EQ(ResolvedNumFmtId(after, 1, 3), 4U);
  EXPECT_EQ(ResolvedNumFmtId(after, 4, 3), 180U);
  // D3's font style index is preserved.
  const Cell* d3_after = after.sheet(0).cell_at(2, 3);
  ASSERT_NE(d3_after, nullptr);
  EXPECT_EQ(d3_after->xf_index, d3_before->xf_index);
}

}  // namespace
}  // namespace formulon
