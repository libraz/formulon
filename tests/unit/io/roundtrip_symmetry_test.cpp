// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

#include <cstdint>
#include <string>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/cf_reader.h"
#include "io/cf_writer.h"
#include "pugixml.hpp"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

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
  io::StylesTable styles;
  styles.fonts.emplace_back();
  styles.fills.emplace_back();
  styles.borders.emplace_back();
  styles.num_fmt_strings.emplace_back("0.000%");
  io::NumFmtRecord nf;
  nf.id = 200;
  nf.format_string_index = 0;
  styles.num_fmts.push_back(nf);
  styles.cell_xfs.emplace_back();
  io::CellXf xf;
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
      "<iconSet iconSet=\"3TrafficLights1\">"
      "<cfvo type=\"percent\" val=\"0\"/>"
      "<cfvo type=\"percent\" val=\"33\"/>"
      "<cfvo type=\"percent\" val=\"67\"/>"
      "</iconSet></cfRule></conditionalFormatting></worksheet>";
  pugi::xml_document before;
  ASSERT_TRUE(test::parse_xml(kBefore, &before));

  auto model = io::read_conditional_formats(before.child("worksheet"));
  ASSERT_TRUE(static_cast<bool>(model)) << "cf read failed: " << model.error().message;

  const std::string written = "<worksheet>" + io::write_conditional_formattings(model.value()) + "</worksheet>";
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

  const std::string written = "<worksheet>" + io::write_conditional_formattings(model.value()) + "</worksheet>";
  pugi::xml_document after;
  ASSERT_TRUE(test::parse_xml(written, &after));

  EXPECT_TRUE(test::attributes_preserved(before, after, "//conditionalFormatting", {"sqref"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//dataBar/cfvo[1]", {"type"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//dataBar/cfvo[2]", {"type"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//dataBar/color", {"rgb"}));
}

}  // namespace
}  // namespace formulon
