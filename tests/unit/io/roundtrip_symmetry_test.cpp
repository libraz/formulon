//
// Read<->write symmetry tests built on tests/support/roundtrip_symmetry.h.
//
// These exercise the whole OOXML pipeline against real Excel-produced XML (the
// `xlsb_fidelity_base.xlsx` fixture) and against component reader/writer pairs,
// asserting that attributes a reader understands are not silently dropped by
// the writer. Scope is limited to areas whose I/O has settled: styles
// (num-format / font colour), workbook-level settings (workbookPr, bookViews),
// and conditional formatting (IconSet cfvo). Sheet / pivot / XLSB coverage is
// deferred until those readers/writers stop churning.

#include "support/roundtrip_symmetry.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/cf_reader.h"
#include "io/cf_writer.h"
#include "io/ooxml_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "styles.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

/// Stands in for the package's `<dxfs>` record count. The CF fixtures
/// below carry no `dxfId`, so any value large enough to be a no-op works.
constexpr std::size_t kDxfCount = 8;

std::string FixturePath(const char* name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
}

// ---------------------------------------------------------------------------
// Real Excel fixture: attributes must survive a full load -> save cycle.
// ---------------------------------------------------------------------------

TEST(RoundtripSymmetry, RealFixtureStylesNumFmtsSurvive) {
  const std::vector<std::uint8_t> pkg = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsx"));
  ASSERT_FALSE(pkg.empty());
  const io::ByteSpan span = test::span_of(pkg);
  // The fixture carries two custom number formats authored by Excel. The
  // writer must re-emit both with their original format codes.
  EXPECT_TRUE(test::part_attributes_survive_save(span, "xl/styles.xml", "//numFmt[@numFmtId='179']",
                                                 {"numFmtId", "formatCode"}));
  EXPECT_TRUE(test::part_attributes_survive_save(span, "xl/styles.xml", "//numFmt[@numFmtId='180']",
                                                 {"numFmtId", "formatCode"}));
}

TEST(RoundtripSymmetry, RealFixtureFontColorSurvives) {
  const std::vector<std::uint8_t> pkg = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsx"));
  ASSERT_FALSE(pkg.empty());
  // A red (FFFF0000) font colour authored by Excel. Selecting the node by its
  // attribute value means the assertion fails if the writer drops or rewrites
  // the colour, independent of font ordering.
  EXPECT_TRUE(
      test::part_attributes_survive_save(test::span_of(pkg), "xl/styles.xml", "//color[@rgb='FFFF0000']", {"rgb"}));
}

TEST(RoundtripSymmetry, RealFixtureWorkbookLevelSurvives) {
  const std::vector<std::uint8_t> pkg = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsx"));
  ASSERT_FALSE(pkg.empty());
  const io::ByteSpan span = test::span_of(pkg);
  // workbookPr (theme version) and the bookViews window / active-tab state.
  EXPECT_TRUE(test::part_attributes_survive_save(span, "xl/workbook.xml", "//workbookPr", {"defaultThemeVersion"}));
  EXPECT_TRUE(test::part_attributes_survive_save(span, "xl/workbook.xml", "//workbookView",
                                                 {"activeTab", "windowWidth", "windowHeight"}));
}

// ---------------------------------------------------------------------------
// Constructed workbook: our own writer output must round-trip idempotently
// (the reader must not drop an attribute the writer emitted).
// ---------------------------------------------------------------------------

TEST(RoundtripSymmetry, ConstructedCustomNumFmtIsIdempotent) {
  Workbook wb = Workbook::create();
  StylesTable styles;
  styles.fonts.emplace_back();
  styles.fills.emplace_back();
  styles.borders.emplace_back();
  styles.num_fmt_strings.emplace_back("0.000%");
  NumFmtRecord nf;
  nf.id = 200;
  nf.format_string_index = 0;
  styles.num_fmts.push_back(nf);
  styles.cell_xfs.emplace_back();
  CellXf xf;
  xf.num_fmt_id = 200;
  styles.cell_xfs.push_back(xf);
  wb.set_styles(std::move(styles));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_xf_index(0, 0, 0, 1)));

  const auto saved = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved)) << "save failed: " << saved.error().message;

  // Our serialised styles.xml, cycled once more, must keep the custom format.
  EXPECT_TRUE(test::part_attributes_survive_save(test::span_of(saved.value()), "xl/styles.xml",
                                                 "//numFmt[@numFmtId='200']", {"numFmtId", "formatCode"}));
}

// ---------------------------------------------------------------------------
// Conditional formatting: component reader/writer symmetry on real-shaped XML.
// ---------------------------------------------------------------------------

TEST(RoundtripSymmetry, IconSetCfvoSurvivesReadWrite) {
  // Real-shaped 3-icon rule with three percentage thresholds.
  constexpr const char* kBefore =
      "<worksheet><conditionalFormatting sqref=\"B2:B10\">"
      "<cfRule type=\"iconSet\" priority=\"1\">"
      "<iconSet iconSet=\"3Arrows\">"
      "<cfvo type=\"percent\" val=\"0\"/>"
      "<cfvo type=\"percent\" val=\"33\"/>"
      "<cfvo type=\"percent\" val=\"67\"/>"
      "</iconSet></cfRule></conditionalFormatting></worksheet>";
  pugi::xml_document before;
  ASSERT_TRUE(test::parse_xml(kBefore, &before));

  auto model = io::read_conditional_formats(before.child("worksheet"));
  ASSERT_TRUE(static_cast<bool>(model)) << "cf read failed: " << model.error().message;

  const std::string written =
      "<worksheet>" + io::write_conditional_formattings(model.value(), kDxfCount) + "</worksheet>";
  pugi::xml_document after;
  ASSERT_TRUE(test::parse_xml(written, &after));

  EXPECT_TRUE(test::attributes_preserved(before, after, "//conditionalFormatting", {"sqref"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//cfRule", {"type", "priority"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//iconSet", {"iconSet"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//iconSet/cfvo[1]", {"type", "val"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//iconSet/cfvo[2]", {"type", "val"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//iconSet/cfvo[3]", {"type", "val"}));
}

TEST(RoundtripSymmetry, DataBarCfvoSurvivesReadWrite) {
  constexpr const char* kBefore =
      "<worksheet><conditionalFormatting sqref=\"C1:C20\">"
      "<cfRule type=\"dataBar\" priority=\"2\">"
      "<dataBar>"
      "<cfvo type=\"min\"/>"
      "<cfvo type=\"max\"/>"
      "<color rgb=\"FF638EC6\"/>"
      "</dataBar></cfRule></conditionalFormatting></worksheet>";
  pugi::xml_document before;
  ASSERT_TRUE(test::parse_xml(kBefore, &before));

  auto model = io::read_conditional_formats(before.child("worksheet"));
  ASSERT_TRUE(static_cast<bool>(model)) << "cf read failed: " << model.error().message;

  const std::string written =
      "<worksheet>" + io::write_conditional_formattings(model.value(), kDxfCount) + "</worksheet>";
  pugi::xml_document after;
  ASSERT_TRUE(test::parse_xml(written, &after));

  EXPECT_TRUE(test::attributes_preserved(before, after, "//conditionalFormatting", {"sqref"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//dataBar/cfvo[1]", {"type"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//dataBar/cfvo[2]", {"type"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//dataBar/color", {"rgb"}));
}

// Style runs (blank formatted cells) cycle through save -> load -> save with
// the second sheet part equal to the first: no growth, no loss.
TEST(RoundtripSymmetry, BlankStyleRunsAreIdempotentAcrossTwoCycles) {
  Workbook wb = Workbook::create();
  wb.mutable_styles().cell_xfs.resize(3U);
  Sheet& sheet = wb.sheet(0);
  ColumnLayout column;
  column.first = 1U;
  column.last = 1U;
  column.has_style = true;
  column.style_xf = 1U;
  sheet.mutable_layout().columns.push_back(column);
  sheet.set_style_runs(0U, 0U, 5U, 2U);
  sheet.set_style_runs(3U, 0U, 3U, 0U);  // only B4 overrides the column
  sheet.set_cell_value(3U, 3U, Value::text("x"));

  const auto first = wb.save();
  ASSERT_TRUE(static_cast<bool>(first)) << first.error().message;
  auto reread = io::read_ooxml(test::span_of(first.value()));
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
  const auto second = reread.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(second)) << second.error().message;

  std::string first_sheet;
  std::string second_sheet;
  ASSERT_TRUE(test::extract_part(test::span_of(first.value()), "xl/worksheets/sheet1.xml", &first_sheet));
  ASSERT_TRUE(test::extract_part(test::span_of(second.value()), "xl/worksheets/sheet1.xml", &second_sheet));
  EXPECT_EQ(first_sheet, second_sheet);
  EXPECT_NE(first_sheet.find("<c r=\"B4\"/>"), std::string::npos);
  EXPECT_EQ(first_sheet.find("<c r=\"A4\""), std::string::npos);
  EXPECT_NE(first_sheet.find("<c r=\"E1\" s=\"2\"/>"), std::string::npos);
  EXPECT_EQ(reread.value().workbook.sheet(0).cell_count(), 5U + 1U + 1U);
}

}  // namespace
}  // namespace formulon
