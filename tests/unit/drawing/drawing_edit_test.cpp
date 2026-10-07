#include "drawing/drawing_edit.h"

#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "unknown_relationship.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kCtDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";
constexpr const char* kRelDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
constexpr const char* kDrawingPath = "xl/drawings/drawing1.xml";
constexpr const char* kRelsPath = "xl/drawings/_rels/drawing1.xml.rels";
constexpr const char* kWsDrOpen =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">";
constexpr const char* kRelsOpen =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
constexpr const char* kImageRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

Bytes ToBytes(const std::string& s) {
  return Bytes(s.begin(), s.end());
}

Bytes Png(std::uint32_t w, std::uint32_t h) {
  Bytes b = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  for (std::uint32_t v : {w, h}) {
    for (int shift = 24; shift >= 0; shift -= 8) {
      b.push_back(static_cast<std::uint8_t>(v >> shift));
    }
  }
  b.insert(b.end(), {8, 6, 0, 0, 0, 0x1F, 0x15, 0xC4, 0x89});
  return b;
}

Bytes Jpeg(std::uint16_t w, std::uint16_t h) {
  return {0xFF,
          0xD8,
          0xFF,
          0xC0,
          0x00,
          0x11,
          0x08,
          static_cast<std::uint8_t>(h >> 8),
          static_cast<std::uint8_t>(h),
          static_cast<std::uint8_t>(w >> 8),
          static_cast<std::uint8_t>(w),
          0x03,
          0xFF,
          0xD9};
}

std::string Marker(const char* tag, std::uint32_t col, std::int64_t col_off, std::uint32_t row, std::int64_t row_off) {
  return std::string("<xdr:") + tag + "><xdr:col>" + std::to_string(col) + "</xdr:col><xdr:colOff>" +
         std::to_string(col_off) + "</xdr:colOff><xdr:row>" + std::to_string(row) + "</xdr:row><xdr:rowOff>" +
         std::to_string(row_off) + "</xdr:rowOff></xdr:" + tag + ">";
}

std::string Pic(std::uint32_t id, const char* rid) {
  return "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" + std::to_string(id) + "\" name=\"Picture " + std::to_string(id - 1) +
         "\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill>"
         "<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"" +
         rid +
         "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>"
         "<a:ext cx=\"381000\" cy=\"762000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>"
         "</xdr:pic><xdr:clientData/>";
}

std::string Rel(const char* id, const char* type, const char* target) {
  return std::string("<Relationship Id=\"") + id + "\" Type=\"" + type + "\" Target=\"" + target + "\"/>";
}

const PassthroughPart* Part(const Workbook& wb, const std::string& path) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      return &part;
    }
  }
  return nullptr;
}

Workbook RoundTrip(const Workbook& wb) {
  auto bytes = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes));
  auto read = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  EXPECT_TRUE(static_cast<bool>(read));
  return std::move(read.value().workbook);
}

std::vector<DrawingObject> List(const Workbook& wb) {
  auto objects = list_drawing_objects(wb, 0);
  EXPECT_TRUE(static_cast<bool>(objects));
  return objects ? objects.value() : std::vector<DrawingObject>{};
}

// A workbook whose first sheet uses 18 pt rows and carries `drawing` with
// `rels`, plus the given media parts.
Workbook WithDrawing(const std::string& drawing, const std::string& rels, const std::vector<std::string>& media) {
  Workbook wb = Workbook::create();
  SheetFormatDefaults& defaults = wb.sheet(0).mutable_format_defaults();
  defaults.default_row_height = 18.0;
  defaults.has_default_row_height = true;
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kDrawingPath, kCtDrawing, ToBytes(drawing)))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kRelsPath, "", ToBytes(rels)))));
  for (const std::string& path : media) {
    EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(path, "", Png(40, 80)))));
  }
  wb.set_default_content_types({{"png", "image/png"}});
  wb.sheet(0).set_drawing_rel_target(kDrawingPath);
  return wb;
}

// Three 30 x 60 pt pictures at B5, D5 and F5 on 18 pt rows: move and size
// with cells, move only, and free floating.
Workbook ThreePlacements() {
  const std::string drawing = std::string(kWsDrOpen) + "<xdr:twoCellAnchor>" + Marker("from", 1, 38100, 4, 50800) +
                              Marker("to", 1, 419100, 7, 127000) + Pic(2, "rId1") +
                              "</xdr:twoCellAnchor><xdr:twoCellAnchor editAs=\"oneCell\">" +
                              Marker("from", 3, 38100, 4, 50800) + Marker("to", 3, 419100, 7, 127000) + Pic(3, "rId2") +
                              "</xdr:twoCellAnchor><xdr:twoCellAnchor editAs=\"absolute\">" +
                              Marker("from", 5, 38100, 4, 50800) + Marker("to", 5, 419100, 7, 127000) + Pic(4, "rId3") +
                              "</xdr:twoCellAnchor></xdr:wsDr>";
  const std::string rels = std::string(kRelsOpen) + Rel("rId1", kImageRel, "../media/image1.png") +
                           Rel("rId2", kImageRel, "../media/image2.png") +
                           Rel("rId3", kImageRel, "../media/image3.png") + "</Relationships>";
  return WithDrawing(drawing, rels, {"xl/media/image1.png", "xl/media/image2.png", "xl/media/image3.png"});
}

void ExpectAnchor(const DrawingObject& obj, const AnchorPoint& from, const AnchorPoint& to) {
  EXPECT_EQ(obj.from.row, from.row);
  EXPECT_EQ(obj.from.row_off, from.row_off);
  EXPECT_EQ(obj.from.col, from.col);
  EXPECT_EQ(obj.from.col_off, from.col_off);
  EXPECT_EQ(obj.to.row, to.row);
  EXPECT_EQ(obj.to.row_off, to.row_off);
  EXPECT_EQ(obj.to.col, to.col);
  EXPECT_EQ(obj.to.col_off, to.col_off);
}

TEST(DrawingEdit, InsertTwoRemoveOneKeepsOthers) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(21))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 1, "=A1*2")));
  ASSERT_TRUE(static_cast<bool>(wb.add_merge(0, MergeRange{2, 0, 3, 1})));

  const Bytes png = Png(40, 80);
  ImageInsertOptions first;
  first.name = "First";
  first.descr = "first image";
  first.row = 1;
  first.col = 1;
  auto id1 = insert_image(wb, 0, png.data(), png.size(), first);
  ASSERT_TRUE(static_cast<bool>(id1));

  const Bytes jpeg = Jpeg(100, 50);
  ImageInsertOptions second;
  second.descr = "second <image>";
  second.anchor_kind = AnchorKind::kTwoCell;
  second.edit_as = EditAs::kOneCell;
  second.row = 10;
  second.col = 2;
  second.width_emu = 952500;
  second.height_emu = 476250;
  auto id2 = insert_image(wb, 0, jpeg.data(), jpeg.size(), second);
  ASSERT_TRUE(static_cast<bool>(id2));
  EXPECT_EQ(id1.value(), 2U);
  EXPECT_EQ(id2.value(), 3U);

  Workbook saved = RoundTrip(wb);
  std::vector<DrawingObject> objects = List(saved);
  ASSERT_EQ(objects.size(), 2U);
  EXPECT_EQ(objects[0].anchor_kind, AnchorKind::kOneCell);
  EXPECT_EQ(objects[0].name, "First");
  EXPECT_EQ(objects[0].descr, "first image");
  EXPECT_EQ(objects[0].cx, 40 * 9525);
  EXPECT_EQ(objects[0].cy, 80 * 9525);
  EXPECT_EQ(objects[0].image_format, ImageFormat::kPng);
  EXPECT_EQ(objects[1].anchor_kind, AnchorKind::kTwoCell);
  EXPECT_EQ(objects[1].edit_as, EditAs::kOneCell);
  EXPECT_EQ(objects[1].image_format, ImageFormat::kJpeg);
  EXPECT_EQ(objects[1].descr, "second <image>");
  ASSERT_NE(Part(saved, objects[0].media_path), nullptr);
  EXPECT_EQ(Part(saved, objects[0].media_path)->bytes, png);
  const std::string jpeg_path = objects[1].media_path;
  EXPECT_EQ(Part(saved, jpeg_path)->bytes, jpeg);

  const std::string png_path = objects[0].media_path;
  ASSERT_TRUE(static_cast<bool>(remove_image(saved, 0, id1.value())));
  objects = List(saved);
  ASSERT_EQ(objects.size(), 1U);
  EXPECT_EQ(objects[0].object_id, id2.value());
  EXPECT_EQ(Part(saved, png_path), nullptr);

  Workbook again = RoundTrip(saved);
  objects = List(again);
  ASSERT_EQ(objects.size(), 1U);
  EXPECT_EQ(objects[0].descr, "second <image>");
  EXPECT_EQ(Part(again, jpeg_path)->bytes, jpeg);
  EXPECT_EQ(Part(again, png_path), nullptr);
  EXPECT_EQ(again.sheet(0).cell_at(0, 1)->formula_text, "=A1*2");
  ASSERT_EQ(again.sheet(0).merges().size(), 1U);
  EXPECT_EQ(again.sheet(0).merges()[0].last_row, 3U);
}

TEST(DrawingEdit, CoexistsWithChart) {
  const std::string drawing =
      std::string(kWsDrOpen) + "<xdr:twoCellAnchor>" + Marker("from", 0, 0, 0, 0) + Marker("to", 6, 0, 14, 0) +
      "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 1\"/>"
      "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
      "</xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
      "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"/>"
      "</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"
      "<xdr:oneCellAnchor>" +
      Marker("from", 8, 0, 1, 0) +
      "<xdr:ext cx=\"914400\" cy=\"457200\"/><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr>"
      "<xdr:cNvPr id=\"3\" name=\"Oval 2\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"ellipse\">"
      "<a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>";
  const std::string rels =
      std::string(kRelsOpen) +
      Rel("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "../charts/chart1.xml") +
      "</Relationships>";
  Workbook wb = WithDrawing(drawing, rels, {});
  const Bytes chart = ToBytes("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>");
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(
      "xl/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", chart))));

  const Bytes png = Png(10, 10);
  auto id = insert_image(wb, 0, png.data(), png.size(), ImageInsertOptions{});
  ASSERT_TRUE(static_cast<bool>(id));
  EXPECT_EQ(id.value(), 4U);

  Workbook saved = RoundTrip(wb);
  std::vector<DrawingObject> objects = List(saved);
  ASSERT_EQ(objects.size(), 3U);
  EXPECT_EQ(objects[0].kind, DrawingObjectKind::kChart);
  EXPECT_EQ(objects[0].to.col, 6U);
  EXPECT_EQ(objects[0].to.row, 14U);
  EXPECT_EQ(objects[1].kind, DrawingObjectKind::kShape);
  EXPECT_EQ(objects[1].name, "Oval 2");
  EXPECT_EQ(objects[2].kind, DrawingObjectKind::kPicture);
  EXPECT_EQ(objects[2].image_rel_id, "rId2");
  ASSERT_NE(Part(saved, "xl/charts/chart1.xml"), nullptr);
  EXPECT_EQ(Part(saved, "xl/charts/chart1.xml")->bytes, chart);
  const std::string saved_rels(Part(saved, kRelsPath)->bytes.begin(), Part(saved, kRelsPath)->bytes.end());
  EXPECT_NE(saved_rels.find("Target=\"../charts/chart1.xml\""), std::string::npos);

  // Removing the picture leaves the chart, its relationship and its part.
  ASSERT_TRUE(static_cast<bool>(remove_image(saved, 0, id.value())));
  objects = List(saved);
  ASSERT_EQ(objects.size(), 2U);
  EXPECT_EQ(objects[0].kind, DrawingObjectKind::kChart);
  EXPECT_NE(Part(saved, "xl/charts/chart1.xml"), nullptr);
  const std::string after_rels(Part(saved, kRelsPath)->bytes.begin(), Part(saved, kRelsPath)->bytes.end());
  EXPECT_NE(after_rels.find("rId1"), std::string::npos);
  EXPECT_EQ(after_rels.find("rId2"), std::string::npos);
  // The chart is not a picture, so it cannot be removed as one.
  auto refused = remove_image(saved, 0, 2);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(DrawingEdit, TwoCellDeletedWithRows) {
  // Rows 5..8 hold all three pictures and are deleted.
  Workbook wb = ThreePlacements();
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 4, 4)));
  const std::vector<DrawingObject> objects = List(wb);
  ASSERT_EQ(objects.size(), 2U);
  EXPECT_EQ(objects[0].object_id, 3U);
  ExpectAnchor(objects[0], {4, 3, 0, 38100}, {7, 3, 76200, 419100});
  EXPECT_EQ(objects[1].object_id, 4U);
  ExpectAnchor(objects[1], {4, 5, 50800, 38100}, {7, 5, 127000, 419100});
  // The removed picture's relationship and media go with it.
  EXPECT_EQ(Part(wb, "xl/media/image1.png"), nullptr);
  EXPECT_NE(Part(wb, "xl/media/image2.png"), nullptr);
  const std::string rels(Part(wb, kRelsPath)->bytes.begin(), Part(wb, kRelsPath)->bytes.end());
  EXPECT_EQ(rels.find("image1.png"), std::string::npos);
  EXPECT_NE(rels.find("image3.png"), std::string::npos);
}

struct AnchorCase {
  const char* name;
  std::uint32_t index;
  std::uint32_t count;
  bool is_delete;
  bool row_axis;
  bool move_and_size_removed;
  AnchorPoint move_and_size_from, move_and_size_to;
  AnchorPoint move_from, move_to;
};

TEST(DrawingEdit, AnchorsFollowMeasuredRowColumnRules) {
  // {row, col, row_off, col_off}; Excel's anchors after each edit.
  const AnchorCase cases[] = {
      {"delete rows 3-4 above",
       2,
       2,
       true,
       true,
       false,
       {2, 1, 50800, 38100},
       {5, 1, 127000, 419100},
       {2, 3, 50800, 38100},
       {5, 3, 127000, 419100}},
      {"delete rows 6-7 inside",
       5,
       2,
       true,
       true,
       false,
       {4, 1, 50800, 38100},
       {5, 1, 127000, 419100},
       {4, 3, 50800, 38100},
       {7, 3, 127000, 419100}},
      {"delete rows 7-10 bottom overlap",
       6,
       4,
       true,
       true,
       false,
       {4, 1, 50800, 38100},
       {6, 1, 0, 419100},
       {4, 3, 50800, 38100},
       {7, 3, 127000, 419100}},
      {"delete rows 3-6 top overlap",
       2,
       4,
       true,
       true,
       false,
       {2, 1, 0, 38100},
       {3, 1, 127000, 419100},
       {2, 3, 0, 38100},
       {5, 3, 76200, 419100}},
      {"delete row 5 first row",
       4,
       1,
       true,
       true,
       false,
       {4, 1, 0, 38100},
       {6, 1, 127000, 419100},
       {4, 3, 0, 38100},
       {7, 3, 76200, 419100}},
      {"delete row 8 last row",
       7,
       1,
       true,
       true,
       false,
       {4, 1, 50800, 38100},
       {7, 1, 0, 419100},
       {4, 3, 50800, 38100},
       {7, 3, 127000, 419100}},
      {"insert rows 6-7 inside",
       5,
       2,
       false,
       true,
       false,
       {4, 1, 50800, 38100},
       {9, 1, 127000, 419100},
       {4, 3, 50800, 38100},
       {7, 3, 127000, 419100}},
      {"delete column C",
       2,
       1,
       true,
       false,
       false,
       {4, 1, 50800, 38100},
       {7, 1, 127000, 419100},
       {4, 2, 50800, 38100},
       {7, 2, 127000, 419100}},
      {"delete column D",
       3,
       1,
       true,
       false,
       false,
       {4, 1, 50800, 38100},
       {7, 1, 127000, 419100},
       {4, 3, 50800, 0},
       {7, 3, 127000, 381000}},
      {"delete columns B-D", 1, 3, true, false, true, {}, {}, {4, 1, 50800, 0}, {7, 1, 127000, 381000}},
  };
  for (const AnchorCase& c : cases) {
    SCOPED_TRACE(c.name);
    Workbook wb = ThreePlacements();
    shift_drawing_anchors(wb, 0, c.index, c.count, c.is_delete, c.row_axis);
    const std::vector<DrawingObject> objects = List(wb);
    ASSERT_EQ(objects.size(), c.move_and_size_removed ? 2U : 3U);
    std::size_t i = 0;
    if (!c.move_and_size_removed) {
      ExpectAnchor(objects[i++], c.move_and_size_from, c.move_and_size_to);
    }
    ExpectAnchor(objects[i++], c.move_from, c.move_to);
    ExpectAnchor(objects[i], {4, 5, 50800, 38100}, {7, 5, 127000, 419100});
  }
}

TEST(DrawingEdit, LargeInsertClampsAnchorsWithoutWrapping) {
  const std::uint32_t huge_count = std::numeric_limits<std::uint32_t>::max();

  Workbook rows = ThreePlacements();
  shift_drawing_anchors(rows, 0, 0, huge_count, false, true);
  std::vector<DrawingObject> objects = List(rows);
  ASSERT_EQ(objects.size(), 3U);
  EXPECT_EQ(objects[0].from.row, Sheet::kMaxRows - 1U);
  EXPECT_EQ(objects[0].to.row, Sheet::kMaxRows - 1U);
  EXPECT_EQ(objects[1].from.row, Sheet::kMaxRows - 1U);
  EXPECT_EQ(objects[1].to.row, Sheet::kMaxRows - 1U);
  EXPECT_EQ(objects[2].from.row, 4U);
  EXPECT_EQ(objects[2].to.row, 7U);

  Workbook columns = ThreePlacements();
  shift_drawing_anchors(columns, 0, 0, huge_count, false, false);
  objects = List(columns);
  ASSERT_EQ(objects.size(), 3U);
  EXPECT_EQ(objects[0].from.col, Sheet::kMaxCols - 1U);
  EXPECT_EQ(objects[0].to.col, Sheet::kMaxCols - 1U);
  EXPECT_EQ(objects[1].from.col, Sheet::kMaxCols - 1U);
  EXPECT_EQ(objects[1].to.col, Sheet::kMaxCols - 1U);
  EXPECT_EQ(objects[2].from.col, 5U);
  EXPECT_EQ(objects[2].to.col, 5U);

  // A regular insertion still shifts anchors by its exact count.
  Workbook ordinary = ThreePlacements();
  shift_drawing_anchors(ordinary, 0, 5, 2, false, true);
  objects = List(ordinary);
  ASSERT_EQ(objects.size(), 3U);
  EXPECT_EQ(objects[0].from.row, 4U);
  EXPECT_EQ(objects[0].to.row, 9U);
  EXPECT_EQ(objects[1].from.row, 4U);
  EXPECT_EQ(objects[1].to.row, 7U);
}

TEST(DrawingEdit, OneCellAndAbsoluteAnchorsOnRowDelete) {
  const std::string drawing = std::string(kWsDrOpen) + "<xdr:oneCellAnchor>" + Marker("from", 1, 38100, 4, 50800) +
                              "<xdr:ext cx=\"381000\" cy=\"762000\"/>" + Pic(2, "rId1") +
                              "</xdr:oneCellAnchor><xdr:absoluteAnchor><xdr:pos x=\"5000\" y=\"6000\"/>"
                              "<xdr:ext cx=\"381000\" cy=\"762000\"/>" +
                              Pic(3, "rId1") + "</xdr:absoluteAnchor></xdr:wsDr>";
  const std::string rels = std::string(kRelsOpen) + Rel("rId1", kImageRel, "../media/image1.png") + "</Relationships>";

  Workbook covering = WithDrawing(drawing, rels, {"xl/media/image1.png"});
  shift_drawing_anchors(covering, 0, 4, 4, true, true);
  std::vector<DrawingObject> objects = List(covering);
  ASSERT_EQ(objects.size(), 2U);
  EXPECT_EQ(objects[0].from.row, 4U);
  EXPECT_EQ(objects[0].from.row_off, 0);
  EXPECT_EQ(objects[0].from.col_off, 38100);
  EXPECT_EQ(objects[0].cy, 762000);
  EXPECT_EQ(objects[1].from.col_off, 5000);
  EXPECT_EQ(objects[1].from.row_off, 6000);

  Workbook above = WithDrawing(drawing, rels, {"xl/media/image1.png"});
  shift_drawing_anchors(above, 0, 0, 2, true, true);
  objects = List(above);
  ASSERT_EQ(objects.size(), 2U);
  EXPECT_EQ(objects[0].from.row, 2U);
  EXPECT_EQ(objects[0].from.row_off, 50800);
}

TEST(DrawingEdit, EditAfterEveryObjectKeepsBytes) {
  Workbook wb = ThreePlacements();
  const Bytes before = Part(wb, kDrawingPath)->bytes;
  shift_drawing_anchors(wb, 0, 20, 3, true, true);
  shift_drawing_anchors(wb, 0, 10, 2, false, false);
  EXPECT_EQ(Part(wb, kDrawingPath)->bytes, before);
}

TEST(DrawingEdit, RefusesXlsbRetainedDrawing) {
  Workbook wb = Workbook::create();
  const Bytes drawing = ToBytes(std::string(kWsDrOpen) + "</xdr:wsDr>");
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kDrawingPath, kCtDrawing, drawing))));
  wb.sheet(0).set_unknown_relationships({UnknownRelationship{"rId1", kRelDrawing, kDrawingPath, false}});
  const Bytes png = Png(4, 4);
  auto inserted = insert_image(wb, 0, png.data(), png.size(), ImageInsertOptions{});
  ASSERT_FALSE(static_cast<bool>(inserted));
  EXPECT_EQ(inserted.error().code, FormulonErrorCode::kIoDrawingUnparseable);
  auto removed = remove_image(wb, 0, 2);
  ASSERT_FALSE(static_cast<bool>(removed));
  EXPECT_EQ(removed.error().code, FormulonErrorCode::kIoDrawingUnparseable);
  EXPECT_EQ(wb.passthrough_parts().size(), 1U);
  EXPECT_EQ(Part(wb, kDrawingPath)->bytes, drawing);
  EXPECT_TRUE(wb.sheet(0).drawing_rel_target().empty());
}

TEST(DrawingEdit, UnparseableDrawingRefusesEdits) {
  Workbook wb =
      WithDrawing(std::string(kWsDrOpen) + "<xdr:twoCellAnchor>", std::string(kRelsOpen) + "</Relationships>", {});
  const Bytes before = Part(wb, kDrawingPath)->bytes;
  const Bytes png = Png(4, 4);
  auto inserted = insert_image(wb, 0, png.data(), png.size(), ImageInsertOptions{});
  ASSERT_FALSE(static_cast<bool>(inserted));
  EXPECT_EQ(inserted.error().code, FormulonErrorCode::kIoDrawingUnparseable);
  auto listed = list_drawing_objects(wb, 0);
  ASSERT_FALSE(static_cast<bool>(listed));
  EXPECT_EQ(listed.error().code, FormulonErrorCode::kIoDrawingUnparseable);
  shift_drawing_anchors(wb, 0, 0, 1, true, true);
  EXPECT_EQ(Part(wb, kDrawingPath)->bytes, before);
  EXPECT_EQ(wb.passthrough_parts().size(), 2U);
}

TEST(DrawingEdit, InsertSaveReloadKeepsBytes) {
  Workbook wb = Workbook::create();
  const Bytes png = Png(16, 8);
  ImageInsertOptions options;
  options.descr = "alt";
  options.row = 3;
  options.col = 2;
  options.row_off = 1000;
  options.col_off = 2000;
  ASSERT_TRUE(static_cast<bool>(insert_image(wb, 0, png.data(), png.size(), options)));
  EXPECT_EQ(wb.sheet(0).drawing_rel_target(), kDrawingPath);
  const Bytes drawing = Part(wb, kDrawingPath)->bytes;
  const Bytes rels = Part(wb, kRelsPath)->bytes;

  Workbook reloaded = RoundTrip(wb);
  ASSERT_NE(Part(reloaded, kDrawingPath), nullptr);
  EXPECT_EQ(Part(reloaded, kDrawingPath)->bytes, drawing);
  EXPECT_EQ(Part(reloaded, kDrawingPath)->content_type, kCtDrawing);
  EXPECT_EQ(Part(reloaded, kRelsPath)->bytes, rels);
  EXPECT_EQ(Part(reloaded, "xl/media/image1.png")->bytes, png);
  bool png_default = false;
  for (const DefaultContentType& def : reloaded.default_content_types()) {
    png_default = png_default || (def.extension == "png" && def.content_type == "image/png");
  }
  EXPECT_TRUE(png_default);
  const std::vector<DrawingObject> objects = List(reloaded);
  ASSERT_EQ(objects.size(), 1U);
  EXPECT_EQ(objects[0].from.row, 3U);
  EXPECT_EQ(objects[0].from.row_off, 1000);
  EXPECT_EQ(objects[0].from.col, 2U);
  EXPECT_EQ(objects[0].from.col_off, 2000);
  EXPECT_EQ(objects[0].cx, 16 * 9525);
  EXPECT_EQ(objects[0].cy, 8 * 9525);
  EXPECT_EQ(objects[0].descr, "alt");
}

TEST(DrawingEdit, TwoCellInsertDerivesToFromSize) {
  // A 30 x 60 pt picture at B5 + (3 pt, 4 pt) on 18 pt rows ends 10 pt into
  // row 8, as Excel anchors it.
  Workbook wb = WithDrawing(std::string(kWsDrOpen) + "</xdr:wsDr>", std::string(kRelsOpen) + "</Relationships>", {});
  const Bytes png = Png(4, 4);
  ImageInsertOptions options;
  options.anchor_kind = AnchorKind::kTwoCell;
  options.row = 4;
  options.col = 1;
  options.row_off = 50800;
  options.col_off = 38100;
  options.width_emu = 381000;
  options.height_emu = 762000;
  ASSERT_TRUE(static_cast<bool>(insert_image(wb, 0, png.data(), png.size(), options)));
  const std::vector<DrawingObject> objects = List(wb);
  ASSERT_EQ(objects.size(), 1U);
  EXPECT_EQ(objects[0].anchor_kind, AnchorKind::kTwoCell);
  EXPECT_EQ(objects[0].edit_as, EditAs::kTwoCell);
  ExpectAnchor(objects[0], {4, 1, 50800, 38100}, {7, 1, 127000, 419100});
  const std::string xml(Part(wb, kDrawingPath)->bytes.begin(), Part(wb, kDrawingPath)->bytes.end());
  EXPECT_EQ(xml.find("editAs"), std::string::npos);
}

TEST(DrawingEdit, IdsStayUniqueAcrossNestedObjects) {
  const std::string drawing =
      std::string(kWsDrOpen) + "<xdr:twoCellAnchor>" + Marker("from", 0, 0, 0, 0) + Marker("to", 2, 0, 2, 0) +
      "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"2\" name=\"Group 1\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>"
      "<xdr:grpSpPr/><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"5\" name=\"Shape 4\"/>"
      "<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>";
  Workbook wb = WithDrawing(drawing, std::string(kRelsOpen) + "</Relationships>", {});
  const Bytes jpeg = Jpeg(8, 8);
  auto id = insert_image(wb, 0, jpeg.data(), jpeg.size(), ImageInsertOptions{});
  ASSERT_TRUE(static_cast<bool>(id));
  EXPECT_EQ(id.value(), 6U);
  EXPECT_NE(Part(wb, "xl/media/image1.jpeg"), nullptr);
  bool jpeg_default = false;
  for (const DefaultContentType& def : wb.default_content_types()) {
    jpeg_default = jpeg_default || (def.extension == "jpeg" && def.content_type == "image/jpeg");
  }
  EXPECT_TRUE(jpeg_default);
}

TEST(DrawingEdit, RejectsInvalidInsertions) {
  Workbook wb = Workbook::create();
  const Bytes png = Png(4, 4);
  const auto code = [&](const ImageInsertOptions& options, const Bytes& bytes) {
    auto result = insert_image(wb, 0, bytes.data(), bytes.size(), options);
    return result ? FormulonErrorCode::kOk : result.error().code;
  };
  ImageInsertOptions absolute;
  absolute.anchor_kind = AnchorKind::kAbsolute;
  EXPECT_EQ(code(absolute, png), FormulonErrorCode::kInvalidArgument);
  ImageInsertOptions negative;
  negative.row_off = -1;
  EXPECT_EQ(code(negative, png), FormulonErrorCode::kInvalidArgument);
  ImageInsertOptions huge;
  huge.width_emu = std::int64_t{1} << 31;
  EXPECT_EQ(code(huge, png), FormulonErrorCode::kInvalidArgument);
  ImageInsertOptions off_sheet;
  off_sheet.row = Sheet::kMaxRows;
  EXPECT_EQ(code(off_sheet, png), FormulonErrorCode::kInvalidArgument);
  ImageInsertOptions past_end;
  past_end.row = Sheet::kMaxRows - 1;
  past_end.height_emu = 762000;
  EXPECT_EQ(code(past_end, png), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(code(ImageInsertOptions{}, ToBytes("not an image")), FormulonErrorCode::kIoImageUnsupported);
  Bytes oversized = Png(4, 4);
  oversized.resize(kMaxImageBytes + 1);
  EXPECT_EQ(code(ImageInsertOptions{}, oversized), FormulonErrorCode::kInvalidArgument);
  auto bad_sheet = insert_image(wb, 3, png.data(), png.size(), ImageInsertOptions{});
  ASSERT_FALSE(static_cast<bool>(bad_sheet));
  EXPECT_EQ(bad_sheet.error().code, FormulonErrorCode::kInvalidArgument);
  // Nothing was added by the refused insertions.
  EXPECT_TRUE(wb.passthrough_parts().empty());
  EXPECT_TRUE(wb.sheet(0).drawing_rel_target().empty());
}

// ---------------------------------------------------------------------------
// set_image_anchor / set_image_z_order

constexpr const char* kNsMc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

struct StringWriter : pugi::xml_writer {
  std::string out;
  void write(const void* data, std::size_t size) override { out.append(static_cast<const char*>(data), size); }
};

std::string Raw(const pugi::xml_node& node) {
  StringWriter writer;
  node.print(writer, "", pugi::format_raw);
  return writer.out;
}

std::vector<Bytes> AllBytes(const Workbook& wb) {
  std::vector<Bytes> out;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    out.push_back(part.bytes);
  }
  return out;
}

void ParseDrawing(const Workbook& wb, pugi::xml_document& doc) {
  ASSERT_NE(Part(wb, kDrawingPath), nullptr);
  ASSERT_TRUE(static_cast<bool>(parse_drawing_part(Part(wb, kDrawingPath)->bytes, doc)));
}

pugi::xml_node PicXfrm(const pugi::xml_node& anchor) {
  return child_local(child_local(anchor_content(anchor), "spPr"), "xfrm");
}

// The anchor with every marker, `xdr:ext`/`xdr:pos` and `a:off`/`a:ext`
// removed: what `set_image_anchor` must leave untouched.
std::string Skeleton(const pugi::xml_node& top) {
  pugi::xml_document copy;
  copy.append_copy(top);
  for (pugi::xml_node anchor : drawing_anchors(copy, true)) {
    for (const char* name : {"from", "to", "ext", "pos"}) {
      anchor.remove_child(child_local(anchor, name));
    }
    pugi::xml_node xfrm = PicXfrm(anchor);
    xfrm.remove_child(child_local(xfrm, "off"));
    xfrm.remove_child(child_local(xfrm, "ext"));
  }
  return Raw(copy.document_element());
}

// Moves a marker `extent` EMU along one axis of sheet 0, whose lines all have
// the default size `sheet_offset_emu` gives line 0.
void Advance(const Workbook& wb, bool row_axis, std::uint32_t& line, std::int64_t& off, std::int64_t extent) {
  const std::int64_t size = sheet_offset_emu(wb, 0, row_axis, 1);
  off += extent;
  while (off >= size) {
    off -= size;
    ++line;
  }
}

struct Box {
  AnchorPoint from, to;
  std::int64_t box_w = 0, box_h = 0;
  std::int64_t off_x = 0, off_y = 0;
};

// The anchor box and `a:off` for a `cx` x `cy` picture rotated `deg` degrees
// placed at `a`'s in-cell marker, computed from line offsets alone.
Box ExpectedBox(const Workbook& wb, const ImageAnchor& a, std::int64_t cx, std::int64_t cy, int deg) {
  const bool swap = (deg >= 45 && deg < 135) || (deg >= 225 && deg < 315);
  Box box;
  box.box_w = swap ? cy : cx;
  box.box_h = swap ? cx : cy;
  box.from = AnchorPoint{a.row, a.col, a.row_off, a.col_off};
  box.to = box.from;
  Advance(wb, false, box.to.col, box.to.col_off, box.box_w);
  Advance(wb, true, box.to.row, box.to.row_off, box.box_h);
  box.off_x = sheet_offset_emu(wb, 0, false, a.col) + a.col_off + (box.box_w - cx) / 2;
  box.off_y = sheet_offset_emu(wb, 0, true, a.row) + a.row_off + (box.box_h - cy) / 2;
  return box;
}

void ExpectPoint(const AnchorPoint& got, const AnchorPoint& want) {
  EXPECT_EQ(got.row, want.row);
  EXPECT_EQ(got.row_off, want.row_off);
  EXPECT_EQ(got.col, want.col);
  EXPECT_EQ(got.col_off, want.col_off);
}

// Checks one anchor element against `box` for a `cx` x `cy` picture.
void ExpectPlaced(const pugi::xml_node& anchor, AnchorKind kind, const Box& box, std::int64_t cx, std::int64_t cy) {
  const DrawingObject obj = read_drawing_object(anchor, {});
  EXPECT_EQ(obj.anchor_kind, kind);
  ExpectPoint(obj.from, box.from);
  if (kind == AnchorKind::kTwoCell) {
    ExpectPoint(obj.to, box.to);
    EXPECT_FALSE(child_local(anchor, "ext"));
  } else {
    EXPECT_FALSE(child_local(anchor, "to"));
    EXPECT_EQ(child_local(anchor, "ext").attribute("cx").as_llong(), box.box_w);
    EXPECT_EQ(child_local(anchor, "ext").attribute("cy").as_llong(), box.box_h);
  }
  EXPECT_FALSE(child_local(anchor, "pos"));
  EXPECT_EQ(local_name(child_local(anchor, "")), "from");
  const pugi::xml_node xfrm = PicXfrm(anchor);
  EXPECT_EQ(child_local(xfrm, "off").attribute("x").as_llong(), box.off_x);
  EXPECT_EQ(child_local(xfrm, "off").attribute("y").as_llong(), box.off_y);
  EXPECT_EQ(child_local(xfrm, "ext").attribute("cx").as_llong(), cx);
  EXPECT_EQ(child_local(xfrm, "ext").attribute("cy").as_llong(), cy);
}

// A picture carrying the measured Excel extras (creationId, useLocalDpi, an
// empty picLocks) plus a crop and a horizontal flip.
std::string RichPic(std::uint32_t id, const std::string& rot_attr) {
  return "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" + std::to_string(id) +
         "\" name=\"PicA\"><a:extLst><a:ext uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\"><a16:creationId "
         "xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" "
         "id=\"{916FAF09-56A5-F8BE-D4D0-73AC940BFF01}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvPicPr><a:picLocks/>"
         "</xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip "
         "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId1\">"
         "<a:extLst><a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\"><a14:useLocalDpi "
         "xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/></a:ext></a:extLst>"
         "</a:blip><a:srcRect l=\"1000\" t=\"2000\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"
         "<xdr:spPr><a:xfrm" +
         rot_attr +
         " flipH=\"1\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"381000\" cy=\"762000\"/></a:xfrm><a:prstGeom "
         "prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>";
}

// `RichPic` under a `kind` ("one", "two" or "absolute") anchor.
std::string RichAnchor(const std::string& kind, std::uint32_t id, int rot_deg, const std::string& edit_as_attr = "") {
  const std::string rot = rot_deg == 0 ? "" : " rot=\"" + std::to_string(std::int64_t{rot_deg} * 60000) + "\"";
  if (kind == "one") {
    return "<xdr:oneCellAnchor>" + Marker("from", 1, 38100, 4, 50800) + "<xdr:ext cx=\"381000\" cy=\"762000\"/>" +
           RichPic(id, rot) + "</xdr:oneCellAnchor>";
  }
  if (kind == "absolute") {
    return "<xdr:absoluteAnchor><xdr:pos x=\"5000\" y=\"6000\"/><xdr:ext cx=\"381000\" cy=\"762000\"/>" +
           RichPic(id, rot) + "</xdr:absoluteAnchor>";
  }
  return "<xdr:twoCellAnchor" + edit_as_attr + ">" + Marker("from", 1, 38100, 4, 50800) +
         Marker("to", 1, 419100, 7, 127000) + RichPic(id, rot) + "</xdr:twoCellAnchor>";
}

std::string Wrapped(const std::string& anchor) {
  return std::string("<mc:AlternateContent xmlns:mc=\"") + kNsMc +
         "\"><mc:Choice xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" Requires=\"a14\">" +
         anchor + "</mc:Choice><mc:Fallback>" + anchor + "</mc:Fallback></mc:AlternateContent>";
}

Workbook OnePicture(const std::string& top) {
  return WithDrawing(std::string(kWsDrOpen) + top + "</xdr:wsDr>",
                     std::string(kRelsOpen) + Rel("rId1", kImageRel, "../media/image1.png") + "</Relationships>",
                     {"xl/media/image1.png"});
}

ImageAnchor Placement(AnchorKind kind, EditAs edit_as, std::int64_t width, std::int64_t height) {
  ImageAnchor a;
  a.anchor_kind = kind;
  a.edit_as = edit_as;
  a.row = 3;
  a.col = 2;
  a.row_off = 1000;
  a.col_off = 2000;
  a.width_emu = width;
  a.height_emu = height;
  return a;
}

ImageAnchor FromListed(const DrawingObject& obj) {
  ImageAnchor a;
  a.anchor_kind = obj.anchor_kind;
  a.edit_as = obj.edit_as;
  a.row = obj.from.row;
  a.col = obj.from.col;
  a.row_off = obj.from.row_off;
  a.col_off = obj.from.col_off;
  a.width_emu = obj.cx;
  a.height_emu = obj.cy;
  return a;
}

FormulonErrorCode Code(const Expected<void, Error>& result) {
  return result ? FormulonErrorCode::kOk : result.error().code;
}

TEST(DrawingAnchorUpdate, SheetOffsetFollowsLineSizes) {
  Workbook wb = ThreePlacements();
  EXPECT_EQ(sheet_offset_emu(wb, 0, true, 0), 0);
  EXPECT_EQ(sheet_offset_emu(wb, 0, true, 6), 6 * 228600);  // 18 pt rows.
  // Columns sum unrounded point widths, so the total drifts from 4 x one width.
  const std::int64_t col = sheet_offset_emu(wb, 0, false, 1);
  EXPECT_GT(col, 0);
  EXPECT_NEAR(static_cast<double>(sheet_offset_emu(wb, 0, false, 4)), 4.0 * static_cast<double>(col), 2.0);
}

TEST(DrawingAnchorUpdate, ListedPlacementWritesNothing) {
  Workbook wb = WithDrawing(std::string(kWsDrOpen) + "</xdr:wsDr>", std::string(kRelsOpen) + "</Relationships>", {});
  const Bytes png = Png(40, 80);
  ImageInsertOptions one;
  one.row = 2;
  one.col = 1;
  one.row_off = 1000;
  one.col_off = 2000;
  ASSERT_TRUE(static_cast<bool>(insert_image(wb, 0, png.data(), png.size(), one)));
  ImageInsertOptions two = one;
  two.anchor_kind = AnchorKind::kTwoCell;
  two.edit_as = EditAs::kOneCell;
  ASSERT_TRUE(static_cast<bool>(insert_image(wb, 0, png.data(), png.size(), two)));
  ImageInsertOptions plain = two;
  plain.edit_as = EditAs::kTwoCell;
  ASSERT_TRUE(static_cast<bool>(insert_image(wb, 0, png.data(), png.size(), plain)));
  const std::vector<Bytes> before = AllBytes(wb);
  for (const DrawingObject& obj : List(wb)) {
    ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, obj.object_id, FromListed(obj))));
    ImageAnchor kept_size = FromListed(obj);
    kept_size.width_emu = 0;
    kept_size.height_emu = 0;
    ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, obj.object_id, kept_size)));
  }
  EXPECT_EQ(AllBytes(wb), before);
}

TEST(DrawingAnchorUpdate, MoveAndResizeRewriteMarkersAndXfrm) {
  Workbook wb = WithDrawing(std::string(kWsDrOpen) + "</xdr:wsDr>", std::string(kRelsOpen) + "</Relationships>", {});
  const Bytes png = Png(40, 80);
  ImageInsertOptions options;
  options.anchor_kind = AnchorKind::kTwoCell;
  auto id = insert_image(wb, 0, png.data(), png.size(), options);
  ASSERT_TRUE(static_cast<bool>(id));
  const ImageAnchor a = Placement(AnchorKind::kTwoCell, EditAs::kTwoCell, 3048000, 762000);
  ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, id.value(), a)));
  const Box box = ExpectedBox(wb, a, 3048000, 762000, 0);
  EXPECT_EQ(box.off_y, 3 * 228600 + 1000);
  pugi::xml_document doc;
  ParseDrawing(wb, doc);
  ExpectPlaced(drawing_anchors(doc.document_element(), false)[0], AnchorKind::kTwoCell, box, 3048000, 762000);
  const std::vector<DrawingObject> objects = List(wb);
  ASSERT_EQ(objects.size(), 1U);
  EXPECT_EQ(objects[0].object_id, id.value());
  EXPECT_EQ(objects[0].cx, 3048000);
  EXPECT_EQ(objects[0].cy, 762000);

  // A zero size keeps the current one; only the marker moves.
  ImageAnchor moved = a;
  moved.row = 9;
  moved.width_emu = 0;
  moved.height_emu = 0;
  ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, id.value(), moved)));
  pugi::xml_document again;
  ParseDrawing(wb, again);
  ExpectPlaced(drawing_anchors(again.document_element(), false)[0], AnchorKind::kTwoCell,
               ExpectedBox(wb, moved, 3048000, 762000, 0), 3048000, 762000);
}

TEST(DrawingAnchorUpdate, MeasuredExcelPictureKeepsEverythingElse) {
  // Excel's own picture XML (measured): editAs="oneCell", creationId,
  // useLocalDpi and an empty picLocks.
  const std::string excel =
      "<xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>317500</xdr:colOff>"
      "<xdr:row>2</xdr:row><xdr:rowOff>127000</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col>"
      "<xdr:colOff>444500</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>127000</xdr:rowOff></xdr:to><xdr:pic>"
      "<xdr:nvPicPr><xdr:cNvPr id=\"3\" name=\"PicA\"><a:extLst><a:ext uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\">"
      "<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" "
      "id=\"{916FAF09-56A5-F8BE-D4D0-73AC940BFF01}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvPicPr><a:picLocks/>"
      "</xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId1\"><a:extLst>"
      "<a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\"><a14:useLocalDpi "
      "xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/></a:ext></a:extLst></a:blip>"
      "<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x=\"1270000\" y=\"635000\"/>"
      "<a:ext cx=\"2032000\" cy=\"1016000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>"
      "</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>";
  Workbook wb = OnePicture(excel);
  pugi::xml_document before;
  ParseDrawing(wb, before);
  const std::string skeleton = Skeleton(before.document_element().first_child());

  const ImageAnchor a = Placement(AnchorKind::kTwoCell, EditAs::kOneCell, 3048000, 762000);
  ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, 3, a)));
  pugi::xml_document after;
  ParseDrawing(wb, after);
  const pugi::xml_node anchor = after.document_element().first_child();
  EXPECT_EQ(Skeleton(anchor), skeleton);
  EXPECT_STREQ(anchor.attribute("editAs").value(), "oneCell");
  ExpectPlaced(anchor, AnchorKind::kTwoCell, ExpectedBox(wb, a, 3048000, 762000, 0), 3048000, 762000);
}

TEST(DrawingAnchorUpdate, RotatedPictureSwapsTheBox) {
  // An odd width/height difference exercises the truncation toward zero.
  struct Case {
    int deg;
    bool swapped;
  };
  for (const Case& c : {Case{90, true}, Case{30, false}, Case{45, true}, Case{135, false}, Case{-90, true},
                        Case{180, false}, Case{314, true}, Case{315, false}}) {
    SCOPED_TRACE(c.deg);
    Workbook wb = OnePicture(RichAnchor("two", 2, c.deg));
    pugi::xml_document before;
    ParseDrawing(wb, before);
    const std::string skeleton = Skeleton(before.document_element().first_child());
    const ImageAnchor a = Placement(AnchorKind::kTwoCell, EditAs::kTwoCell, 381000, 762001);
    ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, 2, a)));
    const Box box = ExpectedBox(wb, a, 381000, 762001, c.deg < 0 ? c.deg + 360 : c.deg);
    EXPECT_EQ(box.box_w, c.swapped ? 762001 : 381000);
    if (c.swapped) {
      EXPECT_EQ(box.off_x - (sheet_offset_emu(wb, 0, false, 2) + 2000), 190500);
      EXPECT_EQ(box.off_y - (3 * 228600 + 1000), -190500);
    }
    pugi::xml_document after;
    ParseDrawing(wb, after);
    const pugi::xml_node anchor = after.document_element().first_child();
    ExpectPlaced(anchor, AnchorKind::kTwoCell, box, 381000, 762001);
    EXPECT_EQ(Skeleton(anchor), skeleton);
    // The listed size is the unrotated one, so the list round-trips.
    const std::vector<DrawingObject> objects = List(wb);
    ASSERT_EQ(objects.size(), 1U);
    EXPECT_EQ(objects[0].cx, 381000);
    EXPECT_EQ(objects[0].cy, 762001);
    const std::vector<Bytes> placed = AllBytes(wb);
    ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, 2, FromListed(objects[0]))));
    EXPECT_EQ(AllBytes(wb), placed);
  }
}

TEST(DrawingAnchorUpdate, ConvertsBetweenAnchorKinds) {
  Workbook one = OnePicture(RichAnchor("one", 2, 0));
  const ImageAnchor to_two = Placement(AnchorKind::kTwoCell, EditAs::kTwoCell, 0, 0);
  ASSERT_TRUE(static_cast<bool>(set_image_anchor(one, 0, 2, to_two)));
  pugi::xml_document doc;
  ParseDrawing(one, doc);
  pugi::xml_node anchor = doc.document_element().first_child();
  EXPECT_STREQ(anchor.name(), "xdr:twoCellAnchor");
  EXPECT_FALSE(anchor.attribute("editAs"));
  ExpectPlaced(anchor, AnchorKind::kTwoCell, ExpectedBox(one, to_two, 381000, 762000, 0), 381000, 762000);
  EXPECT_EQ(local_name(child_local(anchor, "to").next_sibling()), "pic");

  const ImageAnchor back = Placement(AnchorKind::kOneCell, EditAs::kTwoCell, 0, 0);
  ASSERT_TRUE(static_cast<bool>(set_image_anchor(one, 0, 2, back)));
  pugi::xml_document doc2;
  ParseDrawing(one, doc2);
  anchor = doc2.document_element().first_child();
  EXPECT_STREQ(anchor.name(), "xdr:oneCellAnchor");
  ExpectPlaced(anchor, AnchorKind::kOneCell, ExpectedBox(one, back, 381000, 762000, 0), 381000, 762000);
  EXPECT_EQ(local_name(child_local(anchor, "ext").next_sibling()), "pic");
  EXPECT_EQ(local_name(anchor.last_child()), "clientData");

  Workbook absolute = OnePicture(RichAnchor("absolute", 2, 0));
  const ImageAnchor pinned = Placement(AnchorKind::kTwoCell, EditAs::kAbsolute, 0, 0);
  ASSERT_TRUE(static_cast<bool>(set_image_anchor(absolute, 0, 2, pinned)));
  pugi::xml_document doc3;
  ParseDrawing(absolute, doc3);
  anchor = doc3.document_element().first_child();
  EXPECT_STREQ(anchor.name(), "xdr:twoCellAnchor");
  EXPECT_STREQ(anchor.attribute("editAs").value(), "absolute");
  ExpectPlaced(anchor, AnchorKind::kTwoCell, ExpectedBox(absolute, pinned, 381000, 762000, 0), 381000, 762000);
}

TEST(DrawingAnchorUpdate, EditAsChangesOnlyWithItsValue) {
  struct Case {
    const char* attr;
    EditAs input;
    const char* expected;  // nullptr: no attribute.
  };
  for (const Case& c : {Case{" editAs=\"twoCell\"", EditAs::kTwoCell, "twoCell"}, Case{"", EditAs::kTwoCell, nullptr},
                        Case{" editAs=\"oneCell\"", EditAs::kTwoCell, nullptr}, Case{"", EditAs::kAbsolute, "absolute"},
                        Case{" editAs=\"oneCell\"", EditAs::kOneCell, "oneCell"}}) {
    SCOPED_TRACE(std::string(c.attr) + " -> " + std::to_string(static_cast<int>(c.input)));
    Workbook wb = OnePicture(RichAnchor("two", 2, 0, c.attr));
    ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, 2, Placement(AnchorKind::kTwoCell, c.input, 0, 0))));
    pugi::xml_document doc;
    ParseDrawing(wb, doc);
    const pugi::xml_attribute attr = doc.document_element().first_child().attribute("editAs");
    if (c.expected == nullptr) {
      EXPECT_FALSE(attr);
    } else {
      EXPECT_STREQ(attr.value(), c.expected);
    }
    EXPECT_EQ(List(wb)[0].edit_as, c.input);
  }
}

TEST(DrawingAnchorUpdate, AlternateContentMovesBothBranches) {
  Workbook wb = OnePicture(Wrapped(RichAnchor("two", 2, 90)));
  const ImageAnchor a = Placement(AnchorKind::kOneCell, EditAs::kTwoCell, 500000, 0);
  ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, 2, a)));
  pugi::xml_document doc;
  ParseDrawing(wb, doc);
  const std::vector<pugi::xml_node> anchors = drawing_anchors(doc.document_element(), true);
  ASSERT_EQ(anchors.size(), 2U);
  const Box box = ExpectedBox(wb, a, 500000, 762000, 90);
  for (const pugi::xml_node& anchor : anchors) {
    ExpectPlaced(anchor, AnchorKind::kOneCell, box, 500000, 762000);
  }
  EXPECT_EQ(local_name(doc.document_element().first_child()), "AlternateContent");
}

TEST(DrawingAnchorUpdate, RefusesBadPlacements) {
  Workbook wb = ThreePlacements();
  const Bytes before = Part(wb, kDrawingPath)->bytes;
  const ImageAnchor ok = Placement(AnchorKind::kTwoCell, EditAs::kTwoCell, 0, 0);
  ImageAnchor absolute = ok;
  absolute.anchor_kind = AnchorKind::kAbsolute;
  EXPECT_EQ(Code(set_image_anchor(wb, 0, 2, absolute)), FormulonErrorCode::kInvalidArgument);
  ImageAnchor negative = ok;
  negative.col_off = -1;
  EXPECT_EQ(Code(set_image_anchor(wb, 0, 2, negative)), FormulonErrorCode::kInvalidArgument);
  ImageAnchor huge = ok;
  huge.height_emu = std::int64_t{1} << 31;
  EXPECT_EQ(Code(set_image_anchor(wb, 0, 2, huge)), FormulonErrorCode::kInvalidArgument);
  ImageAnchor off_sheet = ok;
  off_sheet.col = Sheet::kMaxCols;
  EXPECT_EQ(Code(set_image_anchor(wb, 0, 2, off_sheet)), FormulonErrorCode::kInvalidArgument);
  ImageAnchor past_end = ok;
  past_end.row = Sheet::kMaxRows - 1;
  EXPECT_EQ(Code(set_image_anchor(wb, 0, 2, past_end)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(set_image_anchor(wb, 0, 9, ok)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(set_image_anchor(wb, 5, 2, ok)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Part(wb, kDrawingPath)->bytes, before);

  // Only pictures move: the chart with id 2 is refused.
  const std::string chart =
      std::string(kWsDrOpen) + "<xdr:twoCellAnchor>" + Marker("from", 0, 0, 0, 0) + Marker("to", 6, 0, 14, 0) +
      "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 1\"/>"
      "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData "
      "uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/></a:graphic></xdr:graphicFrame>"
      "<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>";
  Workbook charted = WithDrawing(chart, std::string(kRelsOpen) + "</Relationships>", {});
  EXPECT_EQ(Code(set_image_anchor(charted, 0, 2, ok)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(set_image_z_order(charted, 0, 2, 0)), FormulonErrorCode::kInvalidArgument);

  Workbook xlsb = Workbook::create();
  const Bytes drawing = ToBytes(std::string(kWsDrOpen) + RichAnchor("two", 2, 0) + "</xdr:wsDr>");
  ASSERT_TRUE(static_cast<bool>(xlsb.add_passthrough_part(PassthroughPart(kDrawingPath, kCtDrawing, drawing))));
  xlsb.sheet(0).set_unknown_relationships({UnknownRelationship{"rId1", kRelDrawing, kDrawingPath, false}});
  EXPECT_EQ(Code(set_image_anchor(xlsb, 0, 2, ok)), FormulonErrorCode::kIoDrawingUnparseable);
  EXPECT_EQ(Code(set_image_z_order(xlsb, 0, 2, 0)), FormulonErrorCode::kIoDrawingUnparseable);
  EXPECT_EQ(Part(xlsb, kDrawingPath)->bytes, drawing);
}

// Pairwise rows (coverwise, strength 2, seed 1) over input kind x current kind
// x edit_as x rotation x AlternateContent x size; see the T-01 model.
struct PairRow {
  AnchorKind input;
  const char* current;
  EditAs edit_as;
  int deg;
  bool wrapped;
  bool sized;
};

TEST(DrawingAnchorUpdate, PairwiseCombinations) {
  constexpr AnchorKind kOne = AnchorKind::kOneCell;
  constexpr AnchorKind kTwo = AnchorKind::kTwoCell;
  const PairRow rows[] = {
      {kOne, "absolute", EditAs::kTwoCell, 90, false, true},  {kTwo, "two", EditAs::kAbsolute, 180, true, true},
      {kTwo, "one", EditAs::kOneCell, 90, false, false},      {kOne, "two", EditAs::kOneCell, 180, true, false},
      {kTwo, "absolute", EditAs::kTwoCell, 0, true, false},   {kOne, "one", EditAs::kAbsolute, 0, false, true},
      {kOne, "absolute", EditAs::kAbsolute, 90, true, false}, {kTwo, "two", EditAs::kTwoCell, 0, false, false},
      {kTwo, "one", EditAs::kTwoCell, 180, false, true},      {kTwo, "one", EditAs::kOneCell, 0, true, true},
      {kTwo, "absolute", EditAs::kOneCell, 0, false, true},   {kTwo, "absolute", EditAs::kOneCell, 180, false, false},
      {kOne, "two", EditAs::kTwoCell, 90, true, false},
  };
  for (const PairRow& row : rows) {
    SCOPED_TRACE(std::string(row.current) + " -> " + (row.input == kOne ? "one" : "two") + " editAs " +
                 std::to_string(static_cast<int>(row.edit_as)) + " rot " + std::to_string(row.deg) +
                 (row.wrapped ? " wrapped" : "") + (row.sized ? " sized" : ""));
    const std::string anchor = RichAnchor(row.current, 2, row.deg);
    Workbook wb = OnePicture(row.wrapped ? Wrapped(anchor) : anchor);
    pugi::xml_document before;
    ParseDrawing(wb, before);
    const std::int64_t cx = row.sized ? 500000 : 381000;
    const std::int64_t cy = row.sized ? 250001 : 762000;
    const ImageAnchor a = Placement(row.input, row.edit_as, row.sized ? cx : 0, row.sized ? cy : 0);
    ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, 2, a)));

    const std::vector<DrawingObject> objects = List(wb);
    ASSERT_EQ(objects.size(), 1U);
    EXPECT_EQ(objects[0].object_id, 2U);
    EXPECT_EQ(objects[0].name, "PicA");
    EXPECT_EQ(objects[0].edit_as, row.input == kTwo ? row.edit_as : EditAs::kOneCell);
    EXPECT_EQ(objects[0].media_path, "xl/media/image1.png");
    pugi::xml_document doc;
    ParseDrawing(wb, doc);
    const std::vector<pugi::xml_node> anchors = drawing_anchors(doc.document_element(), true);
    ASSERT_EQ(anchors.size(), row.wrapped ? 2U : 1U);
    const Box box = ExpectedBox(wb, a, cx, cy, row.deg);
    for (const pugi::xml_node& node : anchors) {
      ExpectPlaced(node, row.input, box, cx, cy);
      const pugi::xml_node xfrm = PicXfrm(node);
      EXPECT_EQ(xfrm.attribute("rot").as_llong(), std::int64_t{row.deg} * 60000);
      EXPECT_STREQ(xfrm.attribute("flipH").value(), "1");
      const bool attr_expected = row.input == kTwo && row.edit_as != EditAs::kTwoCell;
      EXPECT_EQ(static_cast<bool>(node.attribute("editAs")), attr_expected);
    }
    // Everything below the anchor element is the picture as it was.
    const pugi::xml_node old_pic = anchor_content(drawing_anchors(before.document_element(), false)[0]);
    EXPECT_NE(Skeleton(anchors[0]).find(Raw(child_local(old_pic, "nvPicPr"))), std::string::npos);
    EXPECT_NE(Raw(anchor_content(anchors[0])).find(Raw(child_local(old_pic, "blipFill"))), std::string::npos);
  }
}

// Picture 2, shape 3 and picture 4 wrapped in mc:AlternateContent.
Workbook ZOrderFixture() {
  const std::string shape = "<xdr:oneCellAnchor>" + Marker("from", 8, 0, 1, 0) +
                            "<xdr:ext cx=\"914400\" cy=\"457200\"/><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr>"
                            "<xdr:cNvPr id=\"3\" name=\"Oval 2\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp>"
                            "<xdr:clientData/></xdr:oneCellAnchor>";
  return OnePicture(RichAnchor("two", 2, 0) + shape + Wrapped(RichAnchor("one", 4, 0)));
}

std::vector<std::uint32_t> Ids(const Workbook& wb) {
  std::vector<std::uint32_t> ids;
  for (const DrawingObject& obj : List(wb)) {
    ids.push_back(obj.object_id);
  }
  return ids;
}

TEST(DrawingZOrder, MovesTopLevelElements) {
  using Ids3 = std::vector<std::uint32_t>;
  struct Case {
    std::uint32_t id;
    std::uint32_t index;
    Ids3 expected;
  };
  for (const Case& c :
       {Case{4, 0, Ids3{4, 2, 3}}, Case{2, 2, Ids3{3, 4, 2}}, Case{2, 1, Ids3{3, 2, 4}}, Case{4, 1, Ids3{2, 4, 3}}}) {
    SCOPED_TRACE(std::to_string(c.id) + " -> " + std::to_string(c.index));
    Workbook wb = ZOrderFixture();
    ASSERT_TRUE(static_cast<bool>(set_image_z_order(wb, 0, c.id, c.index)));
    EXPECT_EQ(Ids(wb), c.expected);
    pugi::xml_document doc;
    ParseDrawing(wb, doc);
    std::size_t tops = 0;
    for (pugi::xml_node top = doc.document_element().first_child(); top; top = top.next_sibling()) {
      ++tops;
    }
    EXPECT_EQ(tops, 3U);
    // The wrapper travels whole, so its fallback stays with it.
    EXPECT_EQ(drawing_anchors(doc.document_element(), true).size(), 4U);
  }
}

TEST(DrawingZOrder, UnchangedOrderWritesNothing) {
  Workbook wb = ZOrderFixture();
  const std::vector<Bytes> before = AllBytes(wb);
  ASSERT_TRUE(static_cast<bool>(set_image_z_order(wb, 0, 2, 0)));
  ASSERT_TRUE(static_cast<bool>(set_image_z_order(wb, 0, 4, 2)));
  EXPECT_EQ(AllBytes(wb), before);
}

TEST(DrawingZOrder, RefusesOutOfRangeAndNonPictures) {
  Workbook wb = ZOrderFixture();
  const std::vector<Bytes> before = AllBytes(wb);
  EXPECT_EQ(Code(set_image_z_order(wb, 0, 2, 3)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(set_image_z_order(wb, 0, 3, 0)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(set_image_z_order(wb, 0, 7, 0)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(set_image_z_order(wb, 2, 2, 0)), FormulonErrorCode::kInvalidArgument);
  Workbook empty = Workbook::create();
  EXPECT_EQ(Code(set_image_z_order(empty, 0, 2, 0)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(AllBytes(wb), before);
}

}  // namespace
}  // namespace formulon
