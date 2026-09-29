//
// Unit tests for `formulon::io::read_sheet_data`. Each test feeds the
// reader a tiny synthetic `<worksheet>` document and asserts the
// resulting `Workbook` state. Recalc is invoked where the test wants to
// verify formulas evaluate end-to-end.

#include "io/sheet_reader.h"

#include <cstdint>
#include <deque>
#include <string>
#include <tuple>
#include <utility>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "styles.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

/// Returns the cached value stored at (row, col), or Blank when no cell
/// is present. Mirrors the helper used by workbook_recalc_test.cpp.
Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

/// Returns the formula text stored at (row, col), or empty when no cell.
std::string StoredFormula(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->formula_text;
  }
  return {};
}

TEST(SheetReader, EmptySheetDataNoCells) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<worksheet><sheetData/></worksheet>"));
  Workbook wb = Workbook::create();
  ASSERT_EQ(wb.sheet_count(), 1U);

  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  auto rs = read_sheet_data(doc, 0U, wb, ctx, text_storage);
  ASSERT_TRUE(static_cast<bool>(rs));
  EXPECT_EQ(wb.sheet(0).cell_count(), 0U);
  EXPECT_EQ(ctx.pending_sst_cells.size(), 0U);
}

TEST(SheetReader, SimpleLiteralsAndFormulaRecalc) {
  // A1=1, A2=2, A3==A1+A2 ; after recalc, A3=3.
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"
      "<row r=\"2\"><c r=\"A2\"><v>2</v></c></row>"
      "<row r=\"3\"><c r=\"A3\"><f>A1+A2</f></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  ASSERT_TRUE(StoredValue(wb, 0U, 0U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 0U).as_number(), 1.0);
  ASSERT_TRUE(StoredValue(wb, 0U, 1U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 1U, 0U).as_number(), 2.0);
  // Before recalc: formula stored, cached value is blank.
  EXPECT_EQ(StoredFormula(wb, 0U, 2U, 0U), "=A1+A2");
  EXPECT_TRUE(StoredValue(wb, 0U, 2U, 0U).is_blank());

  // After recalc: A3 = 3.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(StoredValue(wb, 0U, 2U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 2U, 0U).as_number(), 3.0);
}

TEST(SheetReader, DuplicateLiteralCellDropsTheEarlierFormulaDependencies) {
  // Literals bypass the workbook's dirty marking on load; a duplicate `<c>`
  // replacing a formula must still unregister that formula's reads.
  pugi::xml_document doc;
  ASSERT_TRUE(
      doc.load_string("<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><f>B1+SUM(C:C)</f></c><c r=\"A1\"><v>5</v></c>"
                      "<c r=\"C1\"><v>7</v></c></row>"
                      "</sheetData></worksheet>"));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  EXPECT_TRUE(StoredFormula(wb, 0U, 0U, 0U).empty());
  ASSERT_TRUE(StoredValue(wb, 0U, 0U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 0U).as_number(), 5.0);
  const eval::CellNodeId a1{0U, 0U, 0U};
  EXPECT_TRUE(wb.recalc_engine().dep_graph().dependencies_of(a1).empty());
  EXPECT_TRUE(wb.recalc_engine().compact_range_precedents_of(a1, wb).empty());
}

TEST(SheetReader, EmptyNumericValueElementIsBlank) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v/></c></row></sheetData></worksheet>"));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  EXPECT_TRUE(StoredValue(wb, 0U, 0U, 0U).is_blank());
}

TEST(SheetReader, FormulaCachedValueSurvivesLoadBeforeRecalc) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(
      "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><f>UNIMPLEMENTED(1)</f><v>42</v></c>"
      "<c r=\"B1\" t=\"str\"><f>UNIMPLEMENTED_TEXT()</f><v>Excel cache</v></c></row></sheetData></worksheet>"));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  EXPECT_EQ(StoredFormula(wb, 0U, 0U, 0U), "=UNIMPLEMENTED(1)");
  ASSERT_TRUE(StoredValue(wb, 0U, 0U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 0U).as_number(), 42.0);
  EXPECT_EQ(StoredFormula(wb, 0U, 0U, 1U), "=UNIMPLEMENTED_TEXT()");
  ASSERT_TRUE(StoredValue(wb, 0U, 0U, 1U).is_text());
  EXPECT_EQ(StoredValue(wb, 0U, 0U, 1U).as_text(), "Excel cache");
}

// A What-If data table's `<f t="dataTable" .../>` carries no body (its
// geometry lives in attributes this reader does not decode), so the
// cell falls back to its cached `<v>` like any formula-less cell -- but
// the drop must be counted, not silent.
TEST(SheetReader, DataTableFormulaFallsBackToCachedValueAndIsCounted) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(
      "<worksheet><sheetData><row r=\"1\">"
      "<c r=\"A1\"><f t=\"dataTable\" ref=\"A1:B2\" dt2D=\"1\" dtr=\"0\" r1=\"C1\" r2=\"C2\"/><v>42</v></c>"
      "</row></sheetData></worksheet>"));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ReadDiagnostics diagnostics;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage, &diagnostics)));

  EXPECT_EQ(StoredFormula(wb, 0U, 0U, 0U), "");
  ASSERT_TRUE(StoredValue(wb, 0U, 0U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 0U).as_number(), 42.0);
  EXPECT_EQ(diagnostics.skipped_feature_count, 1U);
}

// A null `diagnostics` (the default) must not crash -- callers that do
// not track read diagnostics still get the fallback-to-cached-value
// behaviour above.
TEST(SheetReader, DataTableFormulaWithNullDiagnosticsDoesNotCrash) {
  pugi::xml_document doc;
  ASSERT_TRUE(
      doc.load_string("<worksheet><sheetData><row r=\"1\">"
                      "<c r=\"A1\"><f t=\"dataTable\" ref=\"A1:B2\"/><v>7</v></c>"
                      "</row></sheetData></worksheet>"));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  ASSERT_TRUE(StoredValue(wb, 0U, 0U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 0U).as_number(), 7.0);
}

TEST(SheetReader, InlineStringCell) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Hello</t></is></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  Value v = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Hello");
}

TEST(SheetReader, ErrorCellRoundTrips) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"e\"><v>#DIV/0!</v></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  Value v = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(SheetReader, SharedFormulaSlaveShiftsRelativeReferences) {
  // si=0 master at E3 with formula "C3*D3"; slave at E4 with no body.
  // Excel applies the slave's row/column offset to relative references,
  // so E4 must become "C4*D4" and recalc against row 4 inputs.
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"3\"><c r=\"C3\"><v>10</v></c><c r=\"D3\"><v>20</v></c>"
      "<c r=\"E3\"><f t=\"shared\" si=\"0\" ref=\"E3:E4\">C3*D3</f></c></row>"
      "<row r=\"4\"><c r=\"C4\"><v>100</v></c><c r=\"D4\"><v>500</v></c>"
      "<c r=\"E4\"><f t=\"shared\" si=\"0\"/></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  EXPECT_EQ(StoredFormula(wb, 0U, 2U, 4U), "=C3*D3");
  EXPECT_EQ(StoredFormula(wb, 0U, 3U, 4U), "=C4*D4");

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(StoredValue(wb, 0U, 3U, 4U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 3U, 4U).as_number(), 50000.0);
}

TEST(SheetReader, SharedFormulaSlaveWithoutMasterErrors) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"7\"/></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  auto rs = read_sheet_data(doc, 0U, wb, ctx, text_storage);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(SheetReader, IgnoresArrayFormulaWhoseRefDoesNotStartAtItsAnchor) {
  // The ref's last column lies before C1. Previously this was recorded as a
  // spill beginning at C1, so `last_col - anchor_col + 1` underflowed during
  // registration and attempted a huge allocation.
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"C1\"><f t=\"array\" ref=\"A1:B2\">1</f></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  EXPECT_TRUE(ctx.array_anchors.empty());
  EXPECT_EQ(StoredFormula(wb, 0U, 0U, 2U), "=1");
}

TEST(SheetReader, RetainsSingleCellArrayFormulaAnchor) {
  pugi::xml_document doc;
  ASSERT_TRUE(
      doc.load_string("<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1\">IFS(TRUE,1)</f>"
                      "<v>1</v></c></row></sheetData></worksheet>"));
  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  // Spill registration is the caller's job, run after SST resolution --
  // see `SheetReadContext::array_anchors`.
  ASSERT_TRUE(static_cast<bool>(RegisterArraySpills(wb.sheet(0), ctx.array_anchors)));
  const SpillRegion* region = wb.sheet(0).spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 1U);
  EXPECT_EQ(region->cols, 1U);
}

TEST(SheetReader, RejectsFullGridArrayAnchorBeforeSpillWalk) {
  pugi::xml_document doc;
  ASSERT_TRUE(
      doc.load_string("<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:XFD1048576\">1</f>"
                      "</c></row></sheetData></worksheet>"));
  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;

  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  // The budget/validation check lives in `RegisterArraySpills`, which the
  // caller now runs separately after SST resolution.
  auto result = RegisterArraySpills(wb.sheet(0), ctx.array_anchors);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(0U, 0U), nullptr);
}

TEST(SheetReader, RejectsCumulativeArrayAnchorsOverSheetBudget) {
  pugi::xml_document doc;
  ASSERT_TRUE(
      doc.load_string("<worksheet><sheetData><row r=\"1\">"
                      "<c r=\"A1\"><f t=\"array\" ref=\"A1:A600000\">1</f></c>"
                      "<c r=\"B1\"><f t=\"array\" ref=\"B1:B600000\">1</f></c>"
                      "</row></sheetData></worksheet>"));
  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;

  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  auto result = RegisterArraySpills(wb.sheet(0), ctx.array_anchors);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_NE(result.error().context.find("format=ooxml"), std::string::npos);
  EXPECT_NE(result.error().context.find("anchor_row=0 anchor_col=1"), std::string::npos);
  EXPECT_NE(result.error().context.find("used=600000 requested=600000 ceiling=1048576"), std::string::npos);
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(0U, 0U), nullptr);
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(0U, 1U), nullptr);
}

TEST(SheetReader, PendingSstCellsCollected) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>2</v></c></row>"
      "<row r=\"2\"><c r=\"A2\"><v>42</v></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  ASSERT_EQ(ctx.pending_sst_cells.size(), 2U);
  EXPECT_EQ(std::get<0>(ctx.pending_sst_cells[0]), 0U);
  EXPECT_EQ(std::get<1>(ctx.pending_sst_cells[0]), 0U);
  EXPECT_EQ(std::get<2>(ctx.pending_sst_cells[0]), 0U);
  EXPECT_EQ(std::get<0>(ctx.pending_sst_cells[1]), 0U);
  EXPECT_EQ(std::get<1>(ctx.pending_sst_cells[1]), 1U);
  EXPECT_EQ(std::get<2>(ctx.pending_sst_cells[1]), 2U);

  // Placeholders are written as Text("").
  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_text());
  EXPECT_EQ(a1.as_text(), "");
  Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(b1.is_text());
  EXPECT_EQ(b1.as_text(), "");
  // The non-shared-string cell is unaffected.
  Value a2 = StoredValue(wb, 0U, 1U, 0U);
  ASSERT_TRUE(a2.is_number());
  EXPECT_DOUBLE_EQ(a2.as_number(), 42.0);
}

// Mirrors the real read pipeline's order (ooxml_reader.cpp resolves
// `ctx.pending_sst_cells` in its own step 6, then registers spills in
// step 6b): an SST-typed phantom cell inside a spill footprint must
// capture its fully-resolved text, not the unresolved `Text("")`
// placeholder `read_sheet_data` writes while SST resolution is still
// pending. Registering spills before resolving SST would freeze that
// placeholder into the region permanently.
TEST(SheetReader, SpillPhantomCapturesTheSstResolvedValueNotThePlaceholder) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:B1\">1</f><v>1</v></c>"
      "<c r=\"B1\" t=\"s\"><v>0</v></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  ASSERT_EQ(ctx.pending_sst_cells.size(), 1U);
  const std::vector<std::string_view> sst_entries = {"resolved"};
  for (const auto& triple : ctx.pending_sst_cells) {
    wb.sheet(0).set_cell_cached_value_borrowed(std::get<0>(triple), std::get<1>(triple),
                                               Value::text(sst_entries.at(std::get<2>(triple))));
  }
  ASSERT_TRUE(static_cast<bool>(RegisterArraySpills(wb.sheet(0), ctx.array_anchors)));

  const SpillRegion* region = wb.sheet(0).spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  ASSERT_EQ(region->cells.size(), 2U);
  ASSERT_TRUE(region->cells[1].is_text());
  EXPECT_EQ(region->cells[1].as_text(), "resolved");
}

TEST(SheetReader, RejectsOutOfRangeSheetIndex) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<worksheet><sheetData/></worksheet>"));
  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  auto rs = read_sheet_data(doc, /*sheet_index=*/5U, wb, ctx, text_storage);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(SheetReader, MalformedCellPropagates) {
  // <c r="A0"> — row 0 is invalid.
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A0\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  auto rs = read_sheet_data(doc, 0U, wb, ctx, text_storage);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

// `ST_PaneState` has three values and two of them freeze. A sheet saved
// with a movable split position spells its frozen pane `frozenSplit`;
// reading it as unfrozen would scroll the header rows away in a host UI.
TEST(SheetReader, FrozenSplitPaneIsFrozen) {
  const char* xml =
      "<worksheet><sheetViews><sheetView workbookViewId=\"0\">"
      "<pane xSplit=\"2\" ySplit=\"1\" topLeftCell=\"C2\" activePane=\"bottomRight\" state=\"frozenSplit\"/>"
      "</sheetView></sheetViews><sheetData/></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(read_sheet_view_and_layout(doc, 0U, wb)));
  EXPECT_EQ(wb.sheet(0).view().freeze_rows, 1U);
  EXPECT_EQ(wb.sheet(0).view().freeze_cols, 2U);
}

// The third `ST_PaneState` value is a plain split, which freezes nothing.
TEST(SheetReader, SplitPaneIsNotFrozen) {
  const char* xml =
      "<worksheet><sheetViews><sheetView workbookViewId=\"0\">"
      "<pane xSplit=\"2000\" ySplit=\"1000\" state=\"split\"/>"
      "</sheetView></sheetViews><sheetData/></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(read_sheet_view_and_layout(doc, 0U, wb)));
  EXPECT_EQ(wb.sheet(0).view().freeze_rows, 0U);
  EXPECT_EQ(wb.sheet(0).view().freeze_cols, 0U);
}

// A `<c s="...">` naming an xf the styles part does not define resolves
// to the default record, so every cell of a workbook that loaded can be
// looked up in `cell_xfs` and a save re-emits nothing dangling.
TEST(SheetReader, CellStyleIndexPastCellXfsFallsBackToDefault) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" s=\"1\"><v>1</v></c><c r=\"B1\" s=\"900\"><v>2</v></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  StylesTable styles;
  styles.cell_xfs.resize(2);  // valid indices are 0 and 1
  wb.set_styles(std::move(styles));

  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  const Sheet& sheet = wb.sheet(0);
  ASSERT_NE(sheet.cell_at(0U, 0U), nullptr);
  ASSERT_NE(sheet.cell_at(0U, 1U), nullptr);
  EXPECT_EQ(sheet.cell_at(0U, 0U)->xf_index, 1U);
  EXPECT_LT(sheet.cell_at(0U, 1U)->xf_index, wb.styles().cell_xfs.size());
  EXPECT_EQ(sheet.cell_at(0U, 1U)->xf_index, 0U);
}

// Two styled cells far apart force `RowCells::ensure()` to materialise
// every slot between them (`sheet.h`'s documented worst case). A tiny
// injected budget catches this without needing a real multi-hundred-MB
// fixture -- mirrors the finding's own PoC shape (styled cells at the
// row's two ends).
TEST(SheetReader, WideSparseStyledRowRejectedOverCellBudget) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" s=\"1\"><v>1</v></c><c r=\"XFD1\" s=\"1\"><v>2</v></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ReadDiagnostics diagnostics;
  auto rs = read_sheet_data(doc, 0U, wb, ctx, text_storage, &diagnostics, /*max_cell_bytes=*/1000U);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoFileTooLarge);
}

// Bare `<c>` elements with no value, style, or formula never call
// `RowCells::ensure()` (`ApplyParsedCell`'s own no-op path), so charging
// for the gap between them would reject a load that costs the sheet
// nothing at all.
TEST(SheetReader, TrulyEmptySparseCellsDoNotConsumeTheCellBudget) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"/><c r=\"XFD1\"/></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  auto rs = read_sheet_data(doc, 0U, wb, ctx, text_storage, /*diagnostics=*/nullptr, /*max_cell_bytes=*/1000U);
  ASSERT_TRUE(static_cast<bool>(rs)) << rs.error().message;
  EXPECT_EQ(wb.sheet(0).cell_count(), 0U);
}

}  // namespace
}  // namespace io
}  // namespace formulon
