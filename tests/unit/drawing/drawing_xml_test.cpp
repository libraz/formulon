#include "drawing/drawing_xml.h"

#include <cstdint>
#include <string>
#include <vector>

#include "drawing/drawing_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "passthrough_part.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

Bytes ToBytes(const std::string& s) {
  return Bytes(s.begin(), s.end());
}

// A picture (move but don't size), a one-cell shape, an absolute chart, a
// group, a connector and a chartex frame wrapped in mc:AlternateContent.
constexpr const char* kDrawing =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
    "<xdr:twoCellAnchor editAs=\"oneCell\">"
    "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>38100</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>50800</xdr:rowOff>"
    "</xdr:from>"
    "<xdr:to><xdr:col>1</xdr:col><xdr:colOff>419100</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>127000</xdr:rowOff>"
    "</xdr:to>"
    "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"2\" name=\"Picture 1\" descr=\"Logo &amp; mark\"/>"
    "<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>"
    "<xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
    "r:embed=\"rId2\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"
    "<xdr:spPr><a:xfrm><a:off x=\"647700\" y=\"965200\"/><a:ext cx=\"381000\" cy=\"762000\"/></a:xfrm>"
    "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"
    "<xdr:oneCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row>"
    "<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"914400\" cy=\"457200\"/>"
    "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Rectangle 2\"/><xdr:cNvSpPr/></xdr:nvSpPr>"
    "<xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/>"
    "</xdr:oneCellAnchor>"
    "<xdr:absoluteAnchor><xdr:pos x=\"100\" y=\"200\"/><xdr:ext cx=\"4572000\" cy=\"2743200\"/>"
    "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"4\" name=\"Chart 3\"/>"
    "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
    "</xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
    "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"/>"
    "</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor>"
    "<xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>"
    "<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row>"
    "<xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"7\" name=\"Group 6\"/>"
    "<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x=\"1\" y=\"2\"/><a:ext cx=\"600\" cy=\"700\"/>"
    "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"600\" cy=\"700\"/></a:xfrm></xdr:grpSpPr></xdr:grpSp><xdr:clientData/>"
    "</xdr:twoCellAnchor>"
    "<xdr:twoCellAnchor editAs=\"absolute\"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>"
    "<xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff>"
    "<xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:cxnSp macro=\"\"><xdr:nvCxnSpPr>"
    "<xdr:cNvPr id=\"8\" name=\"Straight Connector 7\"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr/></xdr:cxnSp>"
    "<xdr:clientData/></xdr:twoCellAnchor>"
    "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
    "<mc:Choice xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" Requires=\"cx1\">"
    "<xdr:twoCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row>"
    "<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff>"
    "<xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr>"
    "<xdr:cNvPr id=\"9\" name=\"Chart 8\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/>"
    "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"/></a:graphic>"
    "</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></mc:Choice>"
    "<mc:Fallback><xdr:twoCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row>"
    "<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff>"
    "<xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr>"
    "<xdr:cNvPr id=\"0\" name=\"\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/>"
    "</xdr:twoCellAnchor></mc:Fallback></mc:AlternateContent>"
    "</xdr:wsDr>";

constexpr const char* kRels =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" "
    "Target=\"../charts/chart1.xml\"/>"
    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
    "Target=\"../media/image1.png\"/>"
    "<Relationship Id=\"rId3\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" "
    "Target=\"https://example.com/a\" TargetMode=\"External\"/>"
    "</Relationships>";

std::vector<DrawingObject> ReadObjects(const std::string& xml, const std::vector<DrawingRel>& rels) {
  pugi::xml_document doc;
  EXPECT_TRUE(static_cast<bool>(parse_drawing_part(ToBytes(xml), doc)));
  return read_drawing_objects(doc, rels);
}

TEST(DrawingXml, ReadsEveryObjectKindAndAnchor) {
  auto rels = parse_part_rels(ToBytes(kRels), "xl/drawings");
  ASSERT_TRUE(static_cast<bool>(rels));
  const std::vector<DrawingObject> objects = ReadObjects(kDrawing, rels.value());
  ASSERT_EQ(objects.size(), 6U);

  const DrawingObject& pic = objects[0];
  EXPECT_EQ(pic.kind, DrawingObjectKind::kPicture);
  EXPECT_EQ(pic.anchor_kind, AnchorKind::kTwoCell);
  EXPECT_EQ(pic.edit_as, EditAs::kOneCell);
  EXPECT_EQ(pic.object_id, 2U);
  EXPECT_EQ(pic.name, "Picture 1");
  EXPECT_EQ(pic.descr, "Logo & mark");
  EXPECT_EQ(pic.from.col, 1U);
  EXPECT_EQ(pic.from.col_off, 38100);
  EXPECT_EQ(pic.from.row, 4U);
  EXPECT_EQ(pic.from.row_off, 50800);
  EXPECT_EQ(pic.to.col_off, 419100);
  EXPECT_EQ(pic.to.row, 7U);
  EXPECT_EQ(pic.to.row_off, 127000);
  EXPECT_EQ(pic.cx, 381000);
  EXPECT_EQ(pic.cy, 762000);
  EXPECT_EQ(pic.image_rel_id, "rId2");
  EXPECT_EQ(pic.media_path, "xl/media/image1.png");

  const DrawingObject& shape = objects[1];
  EXPECT_EQ(shape.kind, DrawingObjectKind::kShape);
  EXPECT_EQ(shape.anchor_kind, AnchorKind::kOneCell);
  EXPECT_EQ(shape.edit_as, EditAs::kOneCell);
  EXPECT_EQ(shape.from.col, 3U);
  EXPECT_EQ(shape.cx, 914400);
  EXPECT_EQ(shape.cy, 457200);
  EXPECT_TRUE(shape.media_path.empty());

  const DrawingObject& chart = objects[2];
  EXPECT_EQ(chart.kind, DrawingObjectKind::kChart);
  EXPECT_EQ(chart.anchor_kind, AnchorKind::kAbsolute);
  EXPECT_EQ(chart.edit_as, EditAs::kAbsolute);
  EXPECT_EQ(chart.from.col_off, 100);
  EXPECT_EQ(chart.from.row_off, 200);
  EXPECT_EQ(chart.cx, 4572000);
  EXPECT_EQ(chart.name, "Chart 3");

  EXPECT_EQ(objects[3].kind, DrawingObjectKind::kGroup);
  EXPECT_EQ(objects[3].edit_as, EditAs::kTwoCell);
  EXPECT_EQ(objects[3].object_id, 7U);
  EXPECT_EQ(objects[3].cx, 600);
  EXPECT_EQ(objects[3].cy, 700);

  EXPECT_EQ(objects[4].kind, DrawingObjectKind::kConnector);
  EXPECT_EQ(objects[4].edit_as, EditAs::kAbsolute);

  // The chartex frame is reported once, from its mc:Choice branch.
  EXPECT_EQ(objects[5].kind, DrawingObjectKind::kChart);
  EXPECT_EQ(objects[5].object_id, 9U);
  EXPECT_EQ(objects[5].to.row, 12U);
}

TEST(DrawingXml, FallbackAnchorsAreIncludedOnlyOnRequest) {
  pugi::xml_document doc;
  ASSERT_TRUE(static_cast<bool>(parse_drawing_part(ToBytes(kDrawing), doc)));
  EXPECT_EQ(drawing_anchors(doc.document_element(), false).size(), 6U);
  EXPECT_EQ(drawing_anchors(doc.document_element(), true).size(), 7U);
}

TEST(DrawingXml, MatchesElementsByLocalName) {
  const std::string xml =
      "<wsDr xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
      "xmlns:d=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><oneCellAnchor><from><col>2</col>"
      "<colOff>5</colOff><row>3</row><rowOff>6</rowOff></from><ext cx=\"10\" cy=\"20\"/><pic><nvPicPr>"
      "<cNvPr id=\"11\" name=\"P\"/></nvPicPr><blipFill><d:blip xmlns:rel=\"http://schemas.openxmlformats.org/"
      "officeDocument/2006/relationships\" rel:embed=\"rId9\"/></blipFill></pic><clientData/></oneCellAnchor></wsDr>";
  const std::vector<DrawingObject> objects = ReadObjects(xml, {{"rId9", "image", "xl/media/x.gif", false}});
  ASSERT_EQ(objects.size(), 1U);
  EXPECT_EQ(objects[0].kind, DrawingObjectKind::kPicture);
  EXPECT_EQ(objects[0].object_id, 11U);
  EXPECT_EQ(objects[0].from.row, 3U);
  EXPECT_EQ(objects[0].from.col_off, 5);
  EXPECT_EQ(objects[0].image_rel_id, "rId9");
  EXPECT_EQ(objects[0].media_path, "xl/media/x.gif");
}

TEST(DrawingXml, UnparseablePartIsRefused) {
  pugi::xml_document doc;
  auto broken = parse_drawing_part(ToBytes("<xdr:wsDr><xdr:twoCellAnchor>"), doc);
  ASSERT_FALSE(static_cast<bool>(broken));
  EXPECT_EQ(broken.error().code, FormulonErrorCode::kIoDrawingUnparseable);
  auto wrong_root = parse_drawing_part(ToBytes("<c:chartSpace xmlns:c=\"urn:x\"/>"), doc);
  ASSERT_FALSE(static_cast<bool>(wrong_root));
  EXPECT_EQ(wrong_root.error().code, FormulonErrorCode::kIoDrawingUnparseable);
  auto bad_rels = parse_part_rels(ToBytes("<Relationships><Relationship"), "xl/drawings");
  ASSERT_FALSE(static_cast<bool>(bad_rels));
  EXPECT_EQ(bad_rels.error().code, FormulonErrorCode::kIoDrawingUnparseable);
}

TEST(DrawingXml, RelsResolveAgainstOwnerDirectory) {
  auto rels = parse_part_rels(ToBytes(kRels), "xl/drawings");
  ASSERT_TRUE(static_cast<bool>(rels));
  ASSERT_EQ(rels.value().size(), 3U);
  EXPECT_EQ(rels.value()[0].target, "xl/charts/chart1.xml");
  EXPECT_FALSE(rels.value()[0].external);
  EXPECT_EQ(rels.value()[2].target, "https://example.com/a");
  EXPECT_TRUE(rels.value()[2].external);
}

TEST(DrawingXml, RelativeTargets) {
  EXPECT_EQ(relative_target("xl/drawings", "xl/media/image1.png"), "../media/image1.png");
  EXPECT_EQ(relative_target("xl/drawings", "xl/drawings/drawing2.xml"), "drawing2.xml");
  EXPECT_EQ(relative_target("xl/worksheets", "xl/drawings/drawing1.xml"), "../drawings/drawing1.xml");
}

TEST(DrawingXml, AnchorPointWriteReadRoundTrip) {
  pugi::xml_document doc;
  ASSERT_TRUE(static_cast<bool>(parse_drawing_part(ToBytes(kDrawing), doc)));
  const pugi::xml_node anchor = drawing_anchors(doc.document_element(), false)[0];
  const AnchorPoint written{12U, 3U, 9000000000LL, 4};
  write_anchor_point(child_local(anchor, "to"), written);
  const AnchorPoint read = read_anchor_point(child_local(anchor, "to"));
  EXPECT_EQ(read.row, 12U);
  EXPECT_EQ(read.col, 3U);
  EXPECT_EQ(read.row_off, 9000000000LL);
  EXPECT_EQ(read.col_off, 4);
}

TEST(DrawingXml, WorkbookRoundTripKeepsDrawingBytes) {
  Workbook wb = Workbook::create();
  const Bytes png = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13,
                     'I',  'H', 'D', 'R', 0,    0,    0,    40,   0, 0, 0, 80};
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(
      "xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml", ToBytes(kDrawing)))));
  ASSERT_TRUE(static_cast<bool>(
      wb.add_passthrough_part(PassthroughPart("xl/drawings/_rels/drawing1.xml.rels", "", ToBytes(kRels)))));
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/image1.png", "", png))));
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(
      PassthroughPart("xl/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                      ToBytes("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>")))));
  wb.set_default_content_types({{"png", "image/png"}});
  wb.sheet(0).set_drawing_rel_target("xl/drawings/drawing1.xml");

  auto bytes = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes));
  auto read = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  ASSERT_TRUE(static_cast<bool>(read));
  const Workbook& reloaded = read.value().workbook;
  EXPECT_EQ(reloaded.sheet(0).drawing_rel_target(), "xl/drawings/drawing1.xml");
  for (const char* path : {"xl/drawings/drawing1.xml", "xl/media/image1.png", "xl/charts/chart1.xml"}) {
    const PassthroughPart* before = nullptr;
    const PassthroughPart* after = nullptr;
    for (const PassthroughPart& p : wb.passthrough_parts()) {
      before = p.path == path ? &p : before;
    }
    for (const PassthroughPart& p : reloaded.passthrough_parts()) {
      after = p.path == path ? &p : after;
    }
    ASSERT_NE(after, nullptr) << path;
    EXPECT_EQ(after->bytes, before->bytes) << path;
  }
  auto listed = list_drawing_objects(reloaded, 0);
  ASSERT_TRUE(static_cast<bool>(listed));
  ASSERT_EQ(listed.value().size(), 6U);
  EXPECT_EQ(listed.value()[0].media_path, "xl/media/image1.png");
  EXPECT_EQ(listed.value()[0].image_format, ImageFormat::kPng);
}

}  // namespace
}  // namespace formulon
