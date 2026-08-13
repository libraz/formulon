//
// Read<->write symmetry tests for the worksheet part, built on
// tests/support/roundtrip_symmetry.h. A workbook is constructed with the
// worksheet-level features whose I/O has settled -- frozen panes / view flags,
// sheetProtection, autoFilter, hyperlink ranges, printOptions, headerFooter,
// and manual page breaks -- saved once to obtain real writer output, then run
// through `part_attributes_survive_save`: load -> save must keep every listed
// attribute (the reader must not drop what the writer emitted).
//
// Namespace-declaration completeness of the saved package is intentionally NOT
// asserted here (a separate lane owns that fix); these tests use only the
// attribute-set (`attributes_preserved`) granularity.

#include <cstdint>
#include <vector>

#include "gtest/gtest.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

namespace formulon {
namespace {

// Builds a workbook whose first sheet carries a broad slice of worksheet-level
// settings, and returns the saved xlsx bytes.
std::vector<std::uint8_t> BuildRichSheetXlsx() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);

  SheetView& v = s.mutable_view();
  v.zoom_scale = 85;
  v.freeze_rows = 1;
  v.freeze_cols = 2;
  v.show_grid_lines = false;
  v.tab_selected = true;

  SheetProtection& p = s.mutable_protection();
  p.enabled = true;
  p.sheet = true;
  p.objects = true;
  p.format_cells = false;        // non-default (schema default true)
  p.select_locked_cells = true;  // non-default (schema default false)

  s.mutable_hyperlinks().push_back(Hyperlink{0, 0, 1, 1, "https://example.com", "", "", "tip", ""});

  SheetPrintSettings& ps = s.mutable_print_settings();
  ps.print_options_xml = "<printOptions horizontalCentered=\"1\" gridLines=\"1\" headings=\"1\"/>";
  ps.header_footer_xml = "<headerFooter><oddHeader>Report</oddHeader><oddFooter>Foot</oddFooter></headerFooter>";
  ps.manual_row_breaks.push_back(ManualBreak{4, 0, 9, true});
  s.set_auto_filter_xml("<autoFilter ref=\"A1:J8\"/>");

  auto saved = wb.save();
  EXPECT_TRUE(static_cast<bool>(saved)) << "save failed: " << (saved ? "" : saved.error().message);
  if (!saved) {
    return {};
  }
  return std::move(saved.value());
}

TEST(SheetSymmetry, SheetViewAndPaneSurvive) {
  const std::vector<std::uint8_t> pkg = BuildRichSheetXlsx();
  ASSERT_FALSE(pkg.empty());
  const io::ByteSpan span = test::span_of(pkg);
  EXPECT_TRUE(test::part_attributes_survive_save(span, "xl/worksheets/sheet1.xml", "//sheetView",
                                                 {"showGridLines", "tabSelected", "zoomScale", "workbookViewId"}));
  EXPECT_TRUE(
      test::part_attributes_survive_save(span, "xl/worksheets/sheet1.xml", "//pane", {"xSplit", "ySplit", "state"}));
}

TEST(SheetSymmetry, SheetProtectionFlagsSurvive) {
  const std::vector<std::uint8_t> pkg = BuildRichSheetXlsx();
  ASSERT_FALSE(pkg.empty());
  // Both non-default OFF/ON flags and the enabling flags must round-trip.
  EXPECT_TRUE(test::part_attributes_survive_save(test::span_of(pkg), "xl/worksheets/sheet1.xml", "//sheetProtection",
                                                 {"sheet", "objects", "formatCells", "selectLockedCells"}));
}

TEST(SheetSymmetry, AutoFilterAndHyperlinkRangeSurvive) {
  const std::vector<std::uint8_t> pkg = BuildRichSheetXlsx();
  ASSERT_FALSE(pkg.empty());
  const io::ByteSpan span = test::span_of(pkg);
  EXPECT_TRUE(test::part_attributes_survive_save(span, "xl/worksheets/sheet1.xml", "//autoFilter", {"ref"}));
  // The multi-cell hyperlink `ref` range and its tooltip must survive.
  EXPECT_TRUE(test::part_attributes_survive_save(span, "xl/worksheets/sheet1.xml", "//hyperlink", {"ref", "tooltip"}));
}

TEST(SheetSymmetry, PrintOptionsHeaderFooterAndBreaksSurvive) {
  const std::vector<std::uint8_t> pkg = BuildRichSheetXlsx();
  ASSERT_FALSE(pkg.empty());
  const io::ByteSpan span = test::span_of(pkg);
  EXPECT_TRUE(test::part_attributes_survive_save(span, "xl/worksheets/sheet1.xml", "//printOptions",
                                                 {"horizontalCentered", "gridLines", "headings"}));
  EXPECT_TRUE(test::part_attributes_survive_save(span, "xl/worksheets/sheet1.xml", "//rowBreaks",
                                                 {"count", "manualBreakCount"}));
  EXPECT_TRUE(
      test::part_attributes_survive_save(span, "xl/worksheets/sheet1.xml", "//rowBreaks/brk", {"id", "max", "man"}));
}

TEST(SheetSymmetry, HeaderFooterElementSurvives) {
  // headerFooter carries element text rather than attributes; assert the
  // oddHeader / oddFooter elements survive the cycle.
  const std::vector<std::uint8_t> pkg = BuildRichSheetXlsx();
  ASSERT_FALSE(pkg.empty());
  std::vector<std::uint8_t> cycled;
  ASSERT_TRUE(test::load_save_cycle(test::span_of(pkg), &cycled));
  std::string after_xml;
  ASSERT_TRUE(test::extract_part(test::span_of(cycled), "xl/worksheets/sheet1.xml", &after_xml));
  pugi::xml_document after;
  ASSERT_TRUE(test::parse_xml(after_xml, &after));
  EXPECT_TRUE(after.select_node("//headerFooter/oddHeader")) << after_xml;
  EXPECT_TRUE(after.select_node("//headerFooter/oddFooter")) << after_xml;
}

}  // namespace
}  // namespace formulon
