#include "io/ooxml/drawing_guides.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "drawing/shape_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "passthrough_part.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr const char* kOpen =
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">";

// A drawing part holding one shape of preset `prst` with `guides` inside its `a:avLst`.
PassthroughPart Drawing(const std::string& prst, const std::string& guides) {
  const std::string xml =
      std::string(kOpen) +
      "<xdr:absoluteAnchor><xdr:pos x=\"0\" y=\"0\"/><xdr:ext cx=\"9525\" cy=\"9525\"/><xdr:sp>"
      "<xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"Shape 1\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>"
      "<a:prstGeom prst=\"" +
      prst + "\"><a:avLst>" + guides +
      "</a:avLst></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>";
  return PassthroughPart("xl/drawings/drawing1.xml", std::string(io::kCtDrawing),
                         std::vector<std::uint8_t>(xml.begin(), xml.end()));
}

std::string Gd(const char* name, const char* fmla) {
  return std::string("<a:gd name=\"") + name + "\" fmla=\"" + fmla + "\"/>";
}

std::string Text(const PassthroughPart& part) {
  return std::string(part.bytes.begin(), part.bytes.end());
}

TEST(DrawingGuides, APartWithNothingToDropIsWrittenAsItIs) {
  std::size_t refused = 0;
  const std::string defined = Gd("adj", "val 30000");
  EXPECT_FALSE(drawing_guides::without_undefined_guides(Drawing("roundRect", defined), io::kCtDrawing, &refused));
  EXPECT_FALSE(drawing_guides::without_undefined_guides(Drawing("rect", ""), io::kCtDrawing, &refused));
  // A formula guide the preset defines stays as written.
  EXPECT_FALSE(drawing_guides::without_undefined_guides(Drawing("roundRect", Gd("adj", "*/ 60000 1 2")), io::kCtDrawing,
                                                        &refused));
  // A name that is not a preset is not judged; another part type is never parsed.
  EXPECT_FALSE(
      drawing_guides::without_undefined_guides(Drawing("notAShape", Gd("adj9", "val 1")), io::kCtDrawing, &refused));
  EXPECT_FALSE(
      drawing_guides::without_undefined_guides(Drawing("roundRect", Gd("adj9", "val 1")), "application/xml", &refused));
  EXPECT_EQ(refused, 0U);
}

TEST(DrawingGuides, UndefinedAndRepeatedGuidesAreLeftOut) {
  std::size_t refused = 0;
  const std::optional<PassthroughPart> fixed = drawing_guides::without_undefined_guides(
      Drawing("roundRect", Gd("adj", "val 30000") + Gd("adj9", "val 1") + Gd("adj", "val 5")), io::kCtDrawing,
      &refused);
  ASSERT_TRUE(fixed.has_value());
  EXPECT_EQ(refused, 2U);
  EXPECT_EQ(fixed->path, "xl/drawings/drawing1.xml");
  const std::string text = Text(*fixed);
  EXPECT_NE(text.find("<a:avLst><a:gd name=\"adj\" fmla=\"val 30000\"/></a:avLst>"), std::string::npos) << text;
}

TEST(DrawingGuides, AGuideOfAPresetWithoutAnyIsLeftOutUncounted) {
  std::size_t refused = 0;
  const std::optional<PassthroughPart> fixed =
      drawing_guides::without_undefined_guides(Drawing("rect", Gd("adj", "val 20000")), io::kCtDrawing, &refused);
  ASSERT_TRUE(fixed.has_value());
  EXPECT_EQ(refused, 0U);
  EXPECT_NE(Text(*fixed).find("<a:avLst/>"), std::string::npos) << Text(*fixed);
}

// A workbook whose only drawing holds one roundRect stating `guides`.
Workbook WithGuides(const std::string& guides) {
  Workbook wb = Workbook::create();
  ShapeSpec spec;
  spec.fields = shape_field::kAnchor | shape_field::kPrst;
  spec.prst = "roundRect";
  spec.anchor_kind = AnchorKind::kOneCell;
  spec.from = AnchorPoint{1, 1, 0, 0};
  spec.cx = 914400;
  spec.cy = 457200;
  EXPECT_TRUE(static_cast<bool>(insert_shape(wb, 0, spec)));
  const std::string path = wb.sheet(0).drawing_rel_target();
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      std::string text(part.bytes.begin(), part.bytes.end());
      const std::string empty = "<a:avLst/>";
      text.replace(text.find(empty), empty.size(), "<a:avLst>" + guides + "</a:avLst>");
      EXPECT_TRUE(
          static_cast<bool>(wb.replace_passthrough_part(path, std::vector<std::uint8_t>(text.begin(), text.end()))));
    }
  }
  return wb;
}

std::string DrawingText(const Workbook& wb) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == wb.sheet(0).drawing_rel_target()) {
      return std::string(part.bytes.begin(), part.bytes.end());
    }
  }
  return {};
}

TEST(DrawingGuides, BothSavesLeaveOutWhatExcelRefusesAndCountIt) {
  const Workbook clean = WithGuides(Gd("adj", "val 30000"));
  const Workbook bad = WithGuides(Gd("adj", "val 30000") + Gd("adj9", "val 1"));
  const auto xlsx_clean = io::write_ooxml_with_result(clean);
  const auto xlsx_bad = io::write_ooxml_with_result(bad);
  ASSERT_TRUE(static_cast<bool>(xlsx_clean) && static_cast<bool>(xlsx_bad));
  EXPECT_EQ(xlsx_bad.value().diagnostics.deferred_feature_count,
            xlsx_clean.value().diagnostics.deferred_feature_count + 1U);
  const auto xlsb_clean = io::xlsb::write_xlsb_with_result(clean);
  const auto xlsb_bad = io::xlsb::write_xlsb_with_result(bad);
  ASSERT_TRUE(static_cast<bool>(xlsb_clean) && static_cast<bool>(xlsb_bad));
  EXPECT_EQ(xlsb_bad.value().diagnostics.deferred_feature_count,
            xlsb_clean.value().diagnostics.deferred_feature_count + 1U);

  const auto from_xlsx = io::read_ooxml(io::ByteSpan{xlsx_bad.value().bytes.data(), xlsx_bad.value().bytes.size()});
  const auto from_xlsb =
      io::xlsb::read_xlsb(io::ByteSpan{xlsb_bad.value().bytes.data(), xlsb_bad.value().bytes.size()});
  ASSERT_TRUE(static_cast<bool>(from_xlsx) && static_cast<bool>(from_xlsb));
  for (const Workbook* again : {&from_xlsx.value().workbook, &from_xlsb.value().workbook}) {
    const std::string text = DrawingText(*again);
    EXPECT_EQ(text.find("adj9"), std::string::npos) << text;
    EXPECT_NE(text.find("<a:gd name=\"adj\" fmla=\"val 30000\"/>"), std::string::npos) << text;
  }
}

TEST(DrawingGuides, ACleanDrawingIsSavedByteForByte) {
  const Workbook clean = WithGuides(Gd("adj", "val 30000"));
  const std::string before = DrawingText(clean);
  const auto xlsx = io::write_ooxml(clean);
  const auto xlsb = io::xlsb::write_xlsb(clean);
  ASSERT_TRUE(static_cast<bool>(xlsx) && static_cast<bool>(xlsb));
  const auto from_xlsx = io::read_ooxml(io::ByteSpan{xlsx.value().data(), xlsx.value().size()});
  const auto from_xlsb = io::xlsb::read_xlsb(io::ByteSpan{xlsb.value().data(), xlsb.value().size()});
  ASSERT_TRUE(static_cast<bool>(from_xlsx) && static_cast<bool>(from_xlsb));
  EXPECT_EQ(DrawingText(from_xlsx.value().workbook), before);
  EXPECT_EQ(DrawingText(from_xlsb.value().workbook), before);
}

// A read lists only the guides the save keeps.
TEST(DrawingGuides, AReadListsOnlyTheGuidesThePresetDefines) {
  const Workbook bad = WithGuides(Gd("adj9", "val 1") + Gd("adj", "val 30000") + Gd("adj", "val 5"));
  auto shapes = list_shapes(bad, 0);
  ASSERT_TRUE(static_cast<bool>(shapes));
  ASSERT_EQ(shapes.value().size(), 1U);
  EXPECT_EQ(shapes.value()[0].adjust_values, std::vector<ShapeAdjust>({{"adj", 30000}}));
}

}  // namespace
}  // namespace formulon
