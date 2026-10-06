#include "drawing/drawing_edit.h"

#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "passthrough_part.h"
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

}  // namespace
}  // namespace formulon
