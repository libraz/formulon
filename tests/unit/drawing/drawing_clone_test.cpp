#include "drawing/drawing_clone.h"

#include <cstdint>
#include <string>
#include <vector>

#include "drawing/drawing_edit.h"
#include "gtest/gtest.h"
#include "parser/ref_transforms.h"
#include "passthrough_part.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kCtDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";
constexpr const char* kCtChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
constexpr const char* kRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
constexpr const char* kDrawingPath = "xl/drawings/drawing1.xml";
constexpr const char* kChartPath = "xl/charts/chart1.xml";

Bytes ToBytes(const std::string& s) {
  return Bytes(s.begin(), s.end());
}

Bytes Png() {
  Bytes b = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R', 0, 0, 0, 8, 0, 0, 0, 8};
  b.insert(b.end(), {8, 6, 0, 0, 0, 0x1F, 0x15, 0xC4, 0x89});
  return b;
}

std::string Marker(const char* tag, int col, int row) {
  return std::string("<xdr:") + tag + "><xdr:col>" + std::to_string(col) + "</xdr:col><xdr:colOff>0</xdr:colOff>" +
         "<xdr:row>" + std::to_string(row) + "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:" + tag + ">";
}

std::string PictureAnchor() {
  return "<xdr:twoCellAnchor>" + Marker("from", 0, 0) + Marker("to", 2, 2) +
         "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"2\" name=\"Pic\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill>"
         "<a:blip r:embed=\"rId1\"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>";
}

std::string ChartAnchor() {
  return "<xdr:twoCellAnchor>" + Marker("from", 4, 0) + Marker("to", 9, 10) +
         "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"3\" name=\"Chart 3\"/>"
         "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
         "</xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
         "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId2\"/>"
         "</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>";
}

std::string ShapeAnchor() {
  return "<xdr:oneCellAnchor>" + Marker("from", 8, 1) +
         "<xdr:ext cx=\"914400\" cy=\"457200\"/><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"4\" "
         "name=\"Oval\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>";
}

std::string ChartSpace() {
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
         "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:plotArea>"
         "<c:barChart><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:val><c:numRef><c:f>Sheet1!$B$1:$B$3</c:f>"
         "</c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>";
}

// Sheet "Sheet1" with a picture, a chart over Sheet1!$B$1:$B$3 and, optionally, a shape.
Workbook Fixture(bool with_shape) {
  const std::string drawing =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
      PictureAnchor() + ChartAnchor() + (with_shape ? ShapeAnchor() : std::string()) + "</xdr:wsDr>";
  const std::string rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" Type=\"" +
      std::string(kRelNs) + "image\" Target=\"../media/image1.png\"/><Relationship Id=\"rId2\" Type=\"" + kRelNs +
      "chart\" Target=\"../charts/chart1.xml\"/></Relationships>";
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kDrawingPath, kCtDrawing, ToBytes(drawing)))));
  EXPECT_TRUE(static_cast<bool>(
      wb.add_passthrough_part(PassthroughPart("xl/drawings/_rels/drawing1.xml.rels", "", ToBytes(rels)))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/image1.png", "", Png()))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kChartPath, kCtChart, ToBytes(ChartSpace())))));
  wb.set_default_content_types({{"png", "image/png"}});
  wb.sheet(0).set_drawing_rel_target(kDrawingPath);
  return wb;
}

std::string PartText(const Workbook& wb, const std::string& path) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      return std::string(part.bytes.begin(), part.bytes.end());
    }
  }
  return {};
}

std::string ChartText(const Workbook& wb, std::size_t sheet) {
  auto refs = list_chart_parts(wb, sheet);
  EXPECT_TRUE(static_cast<bool>(refs));
  if (!refs || refs.value().empty()) {
    return {};
  }
  return PartText(wb, refs.value().front().part_path);
}

Expected<void, Error> Copy(Workbook& wb, std::size_t target) {
  auto plan = plan_drawing_copy(wb, 0);
  if (!plan) {
    return plan.error();
  }
  const parser::SheetCopyTransform transform("Sheet1", "Sheet1 (2)", {});
  return apply_drawing_copy(wb, target, plan.value(), transform);
}

TEST(DrawingClone, CopiesPictureAndChart) {
  Workbook wb = Fixture(false);
  const std::size_t copy = wb.add_sheet("Sheet1 (2)");
  ASSERT_TRUE(static_cast<bool>(Copy(wb, copy)));
  auto objects = list_drawing_objects(wb, copy);
  ASSERT_TRUE(static_cast<bool>(objects));
  ASSERT_EQ(objects.value().size(), 2U);
  EXPECT_EQ(objects.value()[0].kind, DrawingObjectKind::kPicture);
  EXPECT_EQ(objects.value()[1].kind, DrawingObjectKind::kChart);
  auto source = list_drawing_objects(wb, 0);
  ASSERT_TRUE(static_cast<bool>(source));
  EXPECT_EQ(source.value().size(), 2U);
}

TEST(DrawingClone, RewritesCopiedChartOnly) {
  Workbook wb = Fixture(false);
  const std::size_t copy = wb.add_sheet("Sheet1 (2)");
  ASSERT_TRUE(static_cast<bool>(Copy(wb, copy)));
  EXPECT_NE(ChartText(wb, copy).find("<c:f>'Sheet1 (2)'!$B$1:$B$3</c:f>"), std::string::npos) << ChartText(wb, copy);
  EXPECT_EQ(PartText(wb, kChartPath), ChartSpace());
  EXPECT_NE(ChartText(wb, 0).find("<c:f>Sheet1!$B$1:$B$3</c:f>"), std::string::npos);
}

TEST(DrawingClone, RefusesShape) {
  Workbook wb = Fixture(true);
  const std::size_t parts = wb.passthrough_parts().size();
  auto plan = plan_drawing_copy(wb, 0);
  ASSERT_FALSE(static_cast<bool>(plan));
  EXPECT_EQ(plan.error().code, FormulonErrorCode::kNotImplemented);
  EXPECT_EQ(wb.passthrough_parts().size(), parts);
  EXPECT_EQ(PartText(wb, kChartPath), ChartSpace());
}

TEST(DrawingClone, EmptyPlanForNoDrawing) {
  Workbook wb = Workbook::create();
  auto plan = plan_drawing_copy(wb, 0);
  ASSERT_TRUE(static_cast<bool>(plan));
  EXPECT_TRUE(plan.value().snapshots.empty());
  const std::size_t copy = wb.add_sheet("Sheet1 (2)");
  const parser::SheetCopyTransform transform("Sheet1", "Sheet1 (2)", {});
  EXPECT_TRUE(static_cast<bool>(apply_drawing_copy(wb, copy, plan.value(), transform)));
  EXPECT_TRUE(wb.sheet(copy).drawing_rel_target().empty());
}

}  // namespace
}  // namespace formulon
