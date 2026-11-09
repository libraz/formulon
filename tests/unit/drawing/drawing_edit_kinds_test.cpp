#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "drawing/drawing_edit.h"
#include "drawing/shape_edit.h"
#include "drawing/slicer_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "table.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kCtDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";
constexpr const char* kImageRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
constexpr const char* kHyperlinkRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
constexpr const char* kDrawingPath = "xl/drawings/drawing1.xml";
constexpr const char* kRelsPath = "xl/drawings/_rels/drawing1.xml.rels";
constexpr const char* kLink = "https://example.com/";
constexpr const char* kWsDrOpen =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";
constexpr const char* kRelsOpen =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";

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

std::string Marker(const char* tag, std::uint32_t col, std::uint32_t row) {
  return std::string("<xdr:") + tag + "><xdr:col>" + std::to_string(col) +
         "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" + std::to_string(row) +
         "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:" + tag + ">";
}

std::string Anchor(std::uint32_t row, const std::string& content) {
  return "<xdr:twoCellAnchor>" + Marker("from", 1, row) + Marker("to", 3, row + 3) + content +
         "<xdr:clientData/></xdr:twoCellAnchor>";
}

std::string Pic(std::uint32_t id, const char* rid) {
  return "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" + std::to_string(id) + "\" name=\"Picture " + std::to_string(id - 1) +
         "\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"" + rid +
         "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>"
         "<a:ext cx=\"381000\" cy=\"762000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>"
         "</xdr:pic>";
}

// A rectangle; `cnvpr_body` goes inside its `cNvPr`, `fill` inside its `spPr`.
std::string Sp(std::uint32_t id, const char* name, const std::string& cnvpr_body, const std::string& fill) {
  return "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"" + std::to_string(id) + "\" name=\"" + name +
         "\">" + cnvpr_body +
         "</xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>"
         "<a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>" +
         fill + "</xdr:spPr></xdr:sp>";
}

std::string Frame(std::uint32_t id, const char* uri, const std::string& data) {
  return "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" + std::to_string(id) +
         "\" name=\"Frame " + std::to_string(id) +
         "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" "
         "cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"" +
         uri + "\">" + data + "</a:graphicData></a:graphic></xdr:graphicFrame>";
}

std::string Rel(const char* id, const char* type, const char* target, bool external = false) {
  return std::string("<Relationship Id=\"") + id + "\" Type=\"" + type + "\" Target=\"" + target + "\"" +
         (external ? " TargetMode=\"External\"" : "") + "/>";
}

const PassthroughPart* Part(const Workbook& wb, const std::string& path) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      return &part;
    }
  }
  return nullptr;
}

std::vector<Bytes> AllBytes(const Workbook& wb) {
  std::vector<Bytes> out;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    out.push_back(part.bytes);
  }
  return out;
}

Workbook RoundTrip(const Workbook& wb) {
  auto bytes = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes));
  auto read = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  EXPECT_TRUE(static_cast<bool>(read));
  return std::move(read.value().workbook);
}

std::vector<DrawingObject> List(const Workbook& wb, std::size_t sheet = 0) {
  auto objects = list_drawing_objects(wb, sheet);
  EXPECT_TRUE(static_cast<bool>(objects));
  return objects ? objects.value() : std::vector<DrawingObject>{};
}

std::vector<std::uint32_t> Ids(const Workbook& wb, std::size_t sheet = 0) {
  std::vector<std::uint32_t> ids;
  for (const DrawingObject& obj : List(wb, sheet)) {
    ids.push_back(obj.object_id);
  }
  return ids;
}

std::vector<std::string> Names(const Workbook& wb, std::size_t sheet = 0) {
  std::vector<std::string> names;
  for (const DrawingObject& obj : List(wb, sheet)) {
    names.push_back(obj.name);
  }
  return names;
}

template <typename T>
FormulonErrorCode Code(const Expected<T, Error>& result) {
  return result ? FormulonErrorCode::kOk : result.error().code;
}

void ParseDrawing(const Workbook& wb, std::size_t sheet, pugi::xml_document& doc) {
  const PassthroughPart* part = Part(wb, wb.sheet(sheet).drawing_rel_target());
  ASSERT_NE(part, nullptr);
  ASSERT_TRUE(static_cast<bool>(parse_drawing_part(part->bytes, doc)));
}

std::vector<pugi::xml_node> Tops(const pugi::xml_document& doc) {
  std::vector<pugi::xml_node> tops;
  for (pugi::xml_node top = doc.document_element().first_child(); top; top = top.next_sibling()) {
    if (top.type() == pugi::node_element) {
      tops.push_back(top);
    }
  }
  return tops;
}

// Picture 2, shape 3 filled with image2.png, and shape 4 carrying an external hyperlink.
Workbook ShapeFixture() {
  const std::string drawing =
      std::string(kWsDrOpen) + Anchor(0, Pic(2, "rId1")) +
      Anchor(5, Sp(3, "Rectangle 2", "",
                   "<a:blipFill><a:blip r:embed=\"rId2\"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>")) +
      Anchor(10, Sp(4, "Rectangle 3", "<a:hlinkClick r:id=\"rId3\"/>", "<a:noFill/>")) + "</xdr:wsDr>";
  const std::string rels = std::string(kRelsOpen) + Rel("rId1", kImageRel, "../media/image1.png") +
                           Rel("rId2", kImageRel, "../media/image2.png") + Rel("rId3", kHyperlinkRel, kLink, true) +
                           "</Relationships>";
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kDrawingPath, kCtDrawing, ToBytes(drawing)))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kRelsPath, "", ToBytes(rels)))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/image1.png", "", Png(40, 80)))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/image2.png", "", Png(10, 10)))));
  wb.set_default_content_types({{"png", "image/png"}});
  wb.sheet(0).set_drawing_rel_target(kDrawingPath);
  return wb;
}

// The relationship the `attr` attribute of the first `element` under sheet
// `sheet`'s object `id` names.
io::ooxml::DrawingRel NamedRel(const Workbook& wb, std::size_t sheet, std::uint32_t id, std::string_view element,
                               const char* attr) {
  const std::string path = wb.sheet(sheet).drawing_rel_target();
  const PassthroughPart* rels_part = Part(wb, "xl/drawings/_rels/" + path.substr(path.rfind('/') + 1) + ".rels");
  if (rels_part == nullptr) {
    ADD_FAILURE() << "no drawing rels on sheet " << sheet;
    return {};
  }
  auto rels = io::ooxml::parse_part_rels(rels_part->bytes, "xl/drawings");
  EXPECT_TRUE(static_cast<bool>(rels));
  pugi::xml_document doc;
  ParseDrawing(wb, sheet, doc);
  for (const pugi::xml_node& anchor : drawing_anchors(doc.document_element(), false)) {
    if (read_drawing_object(anchor, {}).object_id != id) {
      continue;
    }
    const pugi::xml_node node = anchor.find_node(
        [element](const pugi::xml_node& n) { return n.type() == pugi::node_element && local_name(n) == element; });
    for (const io::ooxml::DrawingRel& rel : rels.value()) {
      if (rel.id == node.attribute(attr).value()) {
        return rel;
      }
    }
  }
  ADD_FAILURE() << "no relationship for object " << id;
  return {};
}

TEST(DrawingEditKinds, ShapeRenameAndZOrderSurviveSaveAndReload) {
  Workbook wb = ShapeFixture();
  ASSERT_TRUE(static_cast<bool>(set_drawing_name(wb, 0, 3, "Photo frame")));
  ASSERT_TRUE(static_cast<bool>(set_drawing_name(wb, 0, 4, "Link")));
  ASSERT_TRUE(static_cast<bool>(set_drawing_z_order(wb, 0, 4, 0)));
  ASSERT_TRUE(static_cast<bool>(set_drawing_z_order(wb, 0, 3, 1)));
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{4, 3, 2}));
  EXPECT_EQ(Names(wb), (std::vector<std::string>{"Link", "Photo frame", "Picture 1"}));
  // An unchanged order or name writes nothing.
  const std::vector<Bytes> before = AllBytes(wb);
  ASSERT_TRUE(static_cast<bool>(set_drawing_z_order(wb, 0, 3, 1)));
  ASSERT_TRUE(static_cast<bool>(set_drawing_name(wb, 0, 4, "Link")));
  EXPECT_EQ(AllBytes(wb), before);

  const Workbook reloaded = RoundTrip(wb);
  EXPECT_EQ(Ids(reloaded), (std::vector<std::uint32_t>{4, 3, 2}));
  EXPECT_EQ(Names(reloaded), (std::vector<std::string>{"Link", "Photo frame", "Picture 1"}));
  const io::ooxml::DrawingRel link = NamedRel(reloaded, 0, 4, "hlinkClick", "r:id");
  EXPECT_EQ(link.type, kHyperlinkRel);
  EXPECT_TRUE(link.external);
  EXPECT_EQ(link.target, kLink);
  const io::ooxml::DrawingRel fill = NamedRel(reloaded, 0, 3, "blip", "r:embed");
  EXPECT_EQ(fill.target, "xl/media/image2.png");
  ASSERT_NE(Part(reloaded, fill.target), nullptr);
  EXPECT_EQ(Part(reloaded, fill.target)->bytes, Png(10, 10));
}

TEST(DrawingEditKinds, ShapeStillRefusesAnchor) {
  Workbook wb = ShapeFixture();
  const std::vector<Bytes> before = AllBytes(wb);
  DrawingAnchor anchor;
  anchor.row = 20;
  EXPECT_EQ(Code(set_drawing_anchor(wb, 0, 3, anchor)), FormulonErrorCode::kInvalidArgument);
  // An id no object holds is unknown rather than of the wrong kind.
  EXPECT_EQ(Code(set_drawing_anchor(wb, 0, 99, anchor)), FormulonErrorCode::kNotFound);
  EXPECT_EQ(AllBytes(wb), before);
}

TEST(DrawingEditKinds, RemoveDrawingRemovesAShape) {
  Workbook wb = ShapeFixture();
  const std::vector<Bytes> before = AllBytes(wb);
  EXPECT_EQ(Code(remove_drawing(wb, 0, 99)), FormulonErrorCode::kNotFound);
  EXPECT_EQ(AllBytes(wb), before);
  ASSERT_TRUE(static_cast<bool>(remove_drawing(wb, 0, 3)));
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{2, 4}));
  EXPECT_EQ(Code(remove_drawing(wb, 0, 3)), FormulonErrorCode::kNotFound);
}

TEST(DrawingEditKinds, ShapeSnapshotCarriesItsRelationshipsToAnotherSheet) {
  Workbook wb = ShapeFixture();
  const std::size_t other = wb.add_sheet("Other");
  for (const std::uint32_t id : {3U, 4U}) {
    auto snap = snapshot_drawing(wb, 0, id);
    ASSERT_TRUE(static_cast<bool>(snap));
    auto restored = restore_drawing(wb, other, snap.value().data(), snap.value().size(), 0);
    ASSERT_TRUE(static_cast<bool>(restored));
    EXPECT_EQ(restored.value(), id);
  }
  const auto check = [other](const Workbook& book) {
    const std::vector<DrawingObject> objects = List(book, other);
    ASSERT_EQ(objects.size(), 2U);
    EXPECT_EQ(objects[0].kind, DrawingObjectKind::kShape);
    EXPECT_EQ(objects[0].name, "Rectangle 2");
    EXPECT_EQ(objects[1].kind, DrawingObjectKind::kShape);
    EXPECT_EQ(objects[1].name, "Rectangle 3");
    const io::ooxml::DrawingRel fill = NamedRel(book, other, 3, "blip", "r:embed");
    EXPECT_EQ(fill.type, kImageRel);
    EXPECT_FALSE(fill.external);
    EXPECT_EQ(fill.target, "xl/media/image2.png");
    ASSERT_NE(Part(book, fill.target), nullptr);
    EXPECT_EQ(Part(book, fill.target)->bytes, Png(10, 10));
    const io::ooxml::DrawingRel link = NamedRel(book, other, 4, "hlinkClick", "r:id");
    EXPECT_EQ(link.type, kHyperlinkRel);
    EXPECT_TRUE(link.external);
    EXPECT_EQ(link.target, kLink);
  };
  check(wb);
  // The identical media part is shared, not copied.
  EXPECT_EQ(Part(wb, "xl/media/image3.png"), nullptr);
  check(RoundTrip(wb));
  EXPECT_EQ(Ids(wb, 0), (std::vector<std::uint32_t>{2, 3, 4}));
}

// Sheet 0's drawing relationships as "id|type|target|external", in rels order.
std::vector<std::string> Rels(const Workbook& wb) {
  auto rels = io::ooxml::parse_part_rels(Part(wb, kRelsPath)->bytes, "xl/drawings");
  EXPECT_TRUE(static_cast<bool>(rels));
  std::vector<std::string> out;
  for (const io::ooxml::DrawingRel& rel : rels ? rels.value() : std::vector<io::ooxml::DrawingRel>{}) {
    out.push_back(rel.id + "|" + rel.type + "|" + rel.target + "|" + (rel.external ? "1" : "0"));
  }
  return out;
}

TEST(DrawingEditKinds, ShapeRestoreReplacesInPlaceOrAddsACopy) {
  Workbook wb = ShapeFixture();
  auto snap = snapshot_drawing(wb, 0, 3);
  ASSERT_TRUE(static_cast<bool>(snap));
  const std::vector<std::string> rels_before = Rels(wb);
  ASSERT_TRUE(static_cast<bool>(set_drawing_name(wb, 0, 3, "Renamed")));
  ASSERT_TRUE(static_cast<bool>(set_drawing_z_order(wb, 0, 3, 2)));
  ASSERT_EQ(Ids(wb), (std::vector<std::uint32_t>{2, 4, 3}));

  auto restored = restore_drawing(wb, 0, snap.value().data(), snap.value().size(), 0);
  ASSERT_TRUE(static_cast<bool>(restored));
  EXPECT_EQ(restored.value(), 3U);
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{2, 3, 4}));
  EXPECT_EQ(Names(wb), (std::vector<std::string>{"Picture 1", "Rectangle 2", "Rectangle 3"}));
  EXPECT_EQ(Rels(wb), rels_before);
  EXPECT_EQ(NamedRel(wb, 0, 3, "blip", "r:embed").target, "xl/media/image2.png");

  auto link = snapshot_drawing(wb, 0, 4);
  ASSERT_TRUE(static_cast<bool>(link));
  auto copy = restore_drawing(wb, 0, link.value().data(), link.value().size(), kDrawingRestoreNewId);
  ASSERT_TRUE(static_cast<bool>(copy));
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{2, 3, 4, copy.value()}));
  EXPECT_EQ(List(wb)[3].kind, DrawingObjectKind::kShape);
  // The copy names the same hyperlink relationship.
  EXPECT_EQ(Rels(wb), rels_before);
  EXPECT_EQ(NamedRel(wb, 0, copy.value(), "hlinkClick", "r:id").target, kLink);
  const Workbook reloaded = RoundTrip(wb);
  EXPECT_EQ(Ids(reloaded), (std::vector<std::uint32_t>{2, 3, 4, copy.value()}));
}

// Sheet 0 with table Sales over A1:B3, a picture, a slicer on Product and a shape, back to front.
struct SlicerBook {
  Workbook wb = Workbook::create();
  std::uint32_t picture = 0;
  std::uint32_t slicer = 0;
  std::uint32_t shape = 0;
};

SlicerBook MakeSlicerBook() {
  SlicerBook book;
  Workbook& wb = book.wb;
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::text("Product"))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 1, Value::text("Amount"))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::text("Widget"))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 1, Value::number(10.0))));
  TableMetadata table;
  table.sheet_index = 0;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:B3";
  for (const char* name : {"Product", "Amount"}) {
    TableColumn column;
    column.name = name;
    table.columns.push_back(std::move(column));
  }
  EXPECT_TRUE(static_cast<bool>(wb.add_table(std::move(table))));
  const Bytes png = Png(40, 80);
  auto picture = insert_image(wb, 0, png.data(), png.size(), ImageInsertOptions{});
  EXPECT_TRUE(static_cast<bool>(picture));
  SlicerSpec slicer;
  slicer.fields = slicer_field::kAnchor;
  slicer.from = AnchorPoint{3, 4, 0, 0};
  slicer.to = AnchorPoint{13, 6, 0, 0};
  auto frame = add_slicer(wb, 0, 0, 0, slicer);
  EXPECT_TRUE(static_cast<bool>(frame));
  ShapeSpec shape;
  shape.fields = shape_field::kAnchor;
  shape.anchor_kind = AnchorKind::kOneCell;
  shape.from = AnchorPoint{1, 8, 0, 0};
  shape.cx = 914400;
  shape.cy = 457200;
  auto box = insert_shape(wb, 0, shape);
  EXPECT_TRUE(static_cast<bool>(box));
  book.picture = picture ? picture.value() : 0;
  book.slicer = frame ? frame.value() : 0;
  book.shape = box ? box.value() : 0;
  return book;
}

// The slicer frame is a nested `mc:AlternateContent` whose `sle15` choice holds the frame and whose fallback a shape.
void ExpectSlicerFrame(const pugi::xml_node& top, std::uint32_t id) {
  ASSERT_TRUE(is_slicer_frame(top));
  EXPECT_EQ(anchor_cnvpr(top).attribute("id").as_uint(), id);
  const pugi::xml_node wrapper = child_local(top, "AlternateContent");
  EXPECT_STREQ(child_local(wrapper, "Choice").attribute("Requires").value(), "sle15");
  EXPECT_EQ(local_name(child_local(child_local(wrapper, "Choice"), "")), "graphicFrame");
  EXPECT_EQ(local_name(child_local(child_local(wrapper, "Fallback"), "")), "sp");
}

TEST(DrawingEditKinds, SlicerZOrderMovesTheWholeFrame) {
  SlicerBook book = MakeSlicerBook();
  Workbook& wb = book.wb;
  ASSERT_EQ(Ids(wb), (std::vector<std::uint32_t>{book.picture, book.slicer, book.shape}));
  ASSERT_TRUE(static_cast<bool>(set_drawing_z_order(wb, 0, book.slicer, 2)));
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{book.picture, book.shape, book.slicer}));
  // A shape moves past the slicer too, by the same index space.
  ASSERT_TRUE(static_cast<bool>(set_drawing_z_order(wb, 0, book.shape, 0)));
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{book.shape, book.picture, book.slicer}));

  const auto check = [&book](const Workbook& source) {
    EXPECT_EQ(Ids(source), (std::vector<std::uint32_t>{book.shape, book.picture, book.slicer}));
    pugi::xml_document doc;
    ParseDrawing(source, 0, doc);
    const std::vector<pugi::xml_node> tops = Tops(doc);
    ASSERT_EQ(tops.size(), 3U);
    ExpectSlicerFrame(tops[2], book.slicer);
    auto slicers = list_slicers(source, 0);
    ASSERT_TRUE(static_cast<bool>(slicers));
    ASSERT_EQ(slicers.value().size(), 1U);
    EXPECT_EQ(slicers.value()[0].name, "Product");
    EXPECT_EQ(slicers.value()[0].object.object_id, book.slicer);
  };
  check(wb);
  Workbook reloaded = RoundTrip(wb);
  check(reloaded);
  ASSERT_TRUE(static_cast<bool>(set_drawing_z_order(reloaded, 0, book.slicer, 0)));
  EXPECT_EQ(Ids(reloaded), (std::vector<std::uint32_t>{book.slicer, book.shape, book.picture}));
  pugi::xml_document doc;
  ParseDrawing(reloaded, 0, doc);
  ExpectSlicerFrame(Tops(doc)[0], book.slicer);
}

TEST(DrawingEditKinds, RemoveDrawingRemovesASlicerWithItsFrame) {
  SlicerBook book = MakeSlicerBook();
  Workbook& wb = book.wb;
  ASSERT_TRUE(static_cast<bool>(remove_drawing(wb, 0, book.slicer)));
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{book.picture, book.shape}));
  EXPECT_TRUE(list_slicers(wb, 0).value().empty());
  EXPECT_EQ(Code(remove_drawing(wb, 0, book.slicer)), FormulonErrorCode::kNotFound);
}

// The listed object with `id`, or a default one.
DrawingObject ObjectWithId(const Workbook& wb, std::uint32_t id) {
  for (const DrawingObject& obj : List(wb)) {
    if (obj.object_id == id) {
      return obj;
    }
  }
  ADD_FAILURE() << "no object " << id;
  return DrawingObject{};
}

TEST(DrawingEditKinds, UpdateDrawingWritesHiddenAndDescriptionForEveryKind) {
  SlicerBook book = MakeSlicerBook();
  Workbook& wb = book.wb;
  for (const std::uint32_t id : {book.picture, book.slicer, book.shape}) {
    DrawingCommonPatch patch;
    patch.hidden = true;
    patch.description = "Object " + std::to_string(id);
    ASSERT_TRUE(static_cast<bool>(update_drawing(wb, 0, id, patch))) << id;
  }
  DrawingCommonPatch rename;
  rename.name = "Logo";
  ASSERT_TRUE(static_cast<bool>(update_drawing(wb, 0, book.picture, rename)));
  rename.name = "Box";
  ASSERT_TRUE(static_cast<bool>(update_drawing(wb, 0, book.shape, rename)));

  const Workbook reloaded = RoundTrip(wb);
  for (const std::uint32_t id : {book.picture, book.slicer, book.shape}) {
    const DrawingObject obj = ObjectWithId(reloaded, id);
    EXPECT_TRUE(obj.hidden) << id;
    EXPECT_EQ(obj.descr, "Object " + std::to_string(id));
  }
  EXPECT_EQ(ObjectWithId(reloaded, book.picture).name, "Logo");
  EXPECT_EQ(ObjectWithId(reloaded, book.shape).name, "Box");
  // Both branches of the slicer's `mc:AlternateContent` carry the flags.
  pugi::xml_document doc;
  ParseDrawing(reloaded, 0, doc);
  std::size_t flagged = 0;
  doc.document_element().find_node([&](const pugi::xml_node& node) {
    if (node.type() == pugi::node_element && local_name(node) == "cNvPr" &&
        node.attribute("id").as_uint() == book.slicer) {
      EXPECT_STREQ(node.attribute("hidden").value(), "1");
      ++flagged;
    }
    return false;
  });
  EXPECT_EQ(flagged, 2U);
}

TEST(DrawingEditKinds, UpdateDrawingLeavesUnsetFieldsAndRefusesAsAWhole) {
  SlicerBook book = MakeSlicerBook();
  Workbook& wb = book.wb;
  DrawingCommonPatch describe;
  describe.description = "Sales logo";
  ASSERT_TRUE(static_cast<bool>(update_drawing(wb, 0, book.picture, describe)));
  const DrawingObject before = ObjectWithId(wb, book.picture);
  EXPECT_EQ(before.descr, "Sales logo");
  EXPECT_FALSE(before.hidden);

  // Only the start row moves; name, description and the rest of the anchor stay.
  DrawingCommonPatch move;
  move.anchor.from_row = 20;
  ASSERT_TRUE(static_cast<bool>(update_drawing(wb, 0, book.picture, move)));
  const DrawingObject moved = ObjectWithId(wb, book.picture);
  EXPECT_EQ(moved.from.row, 20U);
  EXPECT_EQ(moved.from.col, before.from.col);
  EXPECT_EQ(moved.anchor_kind, before.anchor_kind);
  EXPECT_EQ(moved.cx, before.cx);
  EXPECT_EQ(moved.name, before.name);
  EXPECT_EQ(moved.descr, "Sales logo");

  DrawingCommonPatch clear;
  clear.description = "";
  ASSERT_TRUE(static_cast<bool>(update_drawing(wb, 0, book.picture, clear)));
  EXPECT_TRUE(ObjectWithId(wb, book.picture).descr.empty());

  const std::vector<Bytes> unchanged = AllBytes(wb);
  DrawingCommonPatch bad;
  bad.hidden = true;
  bad.anchor.to_row = 30;  // The picture is one-cell.
  EXPECT_EQ(Code(update_drawing(wb, 0, book.picture, bad)), FormulonErrorCode::kInvalidArgument);
  DrawingCommonPatch empty_name;
  empty_name.name = "";
  EXPECT_EQ(Code(update_drawing(wb, 0, book.shape, empty_name)), FormulonErrorCode::kInvalidArgument);
  DrawingCommonPatch slicer_name;
  slicer_name.name = "Renamed";
  EXPECT_EQ(Code(update_drawing(wb, 0, book.slicer, slicer_name)), FormulonErrorCode::kPreconditionFailed);
  DrawingCommonPatch one_cell_slicer;
  one_cell_slicer.anchor.kind = AnchorKind::kOneCell;
  EXPECT_EQ(Code(update_drawing(wb, 0, book.slicer, one_cell_slicer)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(update_drawing(wb, 0, 99, describe)), FormulonErrorCode::kNotFound);
  EXPECT_EQ(AllBytes(wb), unchanged);
}

TEST(DrawingEditKinds, UpdateDrawingMergesShapeAndSlicerAnchors) {
  SlicerBook book = MakeSlicerBook();
  Workbook& wb = book.wb;
  DrawingCommonPatch shape;
  shape.anchor.from_col = 10;
  shape.anchor.width_emu = 457200;
  ASSERT_TRUE(static_cast<bool>(update_drawing(wb, 0, book.shape, shape)));
  const DrawingObject box = ObjectWithId(wb, book.shape);
  EXPECT_EQ(box.from.col, 10U);
  EXPECT_EQ(box.from.row, 1U);
  EXPECT_EQ(box.cx, 457200);
  EXPECT_EQ(box.cy, 457200);

  DrawingCommonPatch slicer;
  slicer.anchor.to_row = 20;
  slicer.hidden = true;
  ASSERT_TRUE(static_cast<bool>(update_drawing(wb, 0, book.slicer, slicer)));
  auto slicers = list_slicers(wb, 0);
  ASSERT_TRUE(static_cast<bool>(slicers));
  ASSERT_EQ(slicers.value().size(), 1U);
  EXPECT_EQ(slicers.value()[0].object.from.row, 3U);
  EXPECT_EQ(slicers.value()[0].object.to.row, 20U);
  EXPECT_EQ(slicers.value()[0].object.to.col, 6U);
  EXPECT_TRUE(slicers.value()[0].object.hidden);
}

TEST(DrawingEditKinds, SlicerRefusesAnchorRenameAndSnapshot) {
  SlicerBook book = MakeSlicerBook();
  Workbook& wb = book.wb;
  const std::vector<Bytes> before = AllBytes(wb);
  DrawingAnchor anchor;
  anchor.row = 20;
  EXPECT_EQ(Code(set_drawing_anchor(wb, 0, book.slicer, anchor)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(set_drawing_name(wb, 0, book.slicer, "Renamed")), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(snapshot_drawing(wb, 0, book.slicer)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(AllBytes(wb), before);
  EXPECT_EQ(list_slicers(wb, 0).value().size(), 1U);
}

constexpr const char* kSlicerUri = "http://schemas.microsoft.com/office/drawing/2010/slicer";

// Picture 2, slicer 3 wrapped whole in a top-level `mc:AlternateContent`, and a diagram frame 4.
Workbook WrappedSlicerFixture() {
  const std::string slicer =
      Frame(3, kSlicerUri, std::string("<sle:slicer xmlns:sle=\"") + kSlicerUri + "\" name=\"Product\"/>");
  const std::string wrapped =
      "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice "
      "xmlns:sle15=\"http://schemas.microsoft.com/office/drawing/2012/slicer\" Requires=\"sle15\">" +
      Anchor(5, slicer) + "</mc:Choice><mc:Fallback>" + Anchor(5, Sp(3, "Product", "", "")) +
      "</mc:Fallback></mc:AlternateContent>";
  const std::string diagram = Anchor(10, Frame(4, "http://schemas.openxmlformats.org/drawingml/2006/diagram", ""));
  const std::string drawing = std::string(kWsDrOpen) + Anchor(0, Pic(2, "rId1")) + wrapped + diagram + "</xdr:wsDr>";
  const std::string rels = std::string(kRelsOpen) + Rel("rId1", kImageRel, "../media/image1.png") + "</Relationships>";
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kDrawingPath, kCtDrawing, ToBytes(drawing)))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kRelsPath, "", ToBytes(rels)))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/image1.png", "", Png(40, 80)))));
  wb.set_default_content_types({{"png", "image/png"}});
  wb.sheet(0).set_drawing_rel_target(kDrawingPath);
  return wb;
}

TEST(DrawingEditKinds, WrappedSlicerMovesWithItsFallback) {
  Workbook wb = WrappedSlicerFixture();
  ASSERT_EQ(Ids(wb), (std::vector<std::uint32_t>{2, 3, 4}));
  ASSERT_TRUE(static_cast<bool>(set_drawing_z_order(wb, 0, 3, 2)));
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{2, 4, 3}));
  ASSERT_TRUE(static_cast<bool>(set_drawing_z_order(wb, 0, 3, 0)));
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{3, 2, 4}));
  pugi::xml_document doc;
  ParseDrawing(wb, 0, doc);
  const std::vector<pugi::xml_node> tops = Tops(doc);
  ASSERT_EQ(tops.size(), 3U);
  EXPECT_EQ(local_name(tops[0]), "AlternateContent");
  EXPECT_STREQ(child_local(tops[0], "Choice").attribute("Requires").value(), "sle15");
  EXPECT_TRUE(is_slicer_frame(child_local(child_local(tops[0], "Choice"), "twoCellAnchor")));
  EXPECT_EQ(local_name(anchor_content(child_local(child_local(tops[0], "Fallback"), "twoCellAnchor"))), "sp");
  EXPECT_EQ(drawing_anchors(doc.document_element(), true).size(), 4U);

  // Only the slicer uri makes a graphic frame orderable or keeps it from the other edits.
  const std::vector<Bytes> before = AllBytes(wb);
  EXPECT_EQ(Code(set_drawing_z_order(wb, 0, 4, 0)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(set_drawing_name(wb, 0, 3, "Renamed")), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(snapshot_drawing(wb, 0, 3)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Code(remove_drawing(wb, 0, 3)), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(AllBytes(wb), before);

  const Workbook reloaded = RoundTrip(wb);
  EXPECT_EQ(Ids(reloaded), (std::vector<std::uint32_t>{3, 2, 4}));
}

// An unset paint removes the shape's own fill and outline fill elements (an emptied `a:ln` goes too, one
// keeping its width stays), so the style's paint applies; the result survives a save and reload.
TEST(DrawingEditKinds, ShapeUnsetPaintRemovesItsOwnFillAndOutline) {
  const std::string outline = "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></a:ln>";
  const std::string drawing =
      std::string(kWsDrOpen) +
      Anchor(0, Sp(2, "Filled", "", "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>" + outline)) +
      Anchor(5, Sp(3, "Bare outline", "", "<a:noFill/><a:ln><a:noFill/></a:ln>")) + "</xdr:wsDr>";
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kDrawingPath, kCtDrawing, ToBytes(drawing)))));
  wb.sheet(0).set_drawing_rel_target(kDrawingPath);

  ShapeSpec unset;
  unset.fields = shape_field::kFill | shape_field::kLine;
  unset.fill = ShapePaint::kUnset;
  unset.line = ShapePaint::kUnset;
  ASSERT_TRUE(static_cast<bool>(update_shape(wb, 0, 2, unset)));
  ASSERT_TRUE(static_cast<bool>(update_shape(wb, 0, 3, unset)));

  const Workbook reloaded = RoundTrip(wb);
  auto shapes = list_shapes(reloaded, 0);
  ASSERT_TRUE(static_cast<bool>(shapes));
  ASSERT_EQ(shapes.value().size(), 2U);
  for (const ShapeInfo& shape : shapes.value()) {
    EXPECT_EQ(shape.fill, ShapePaint::kUnset) << shape.object.object_id;
    EXPECT_EQ(shape.line, ShapePaint::kUnset) << shape.object.object_id;
  }
  EXPECT_EQ(shapes.value()[0].line_width_emu, 12700);

  pugi::xml_document doc;
  ParseDrawing(reloaded, 0, doc);
  const std::vector<pugi::xml_node> tops = Tops(doc);
  ASSERT_EQ(tops.size(), 2U);
  const pugi::xml_node filled = child_local(anchor_content(tops[0]), "spPr");
  EXPECT_FALSE(child_local(filled, "solidFill"));
  ASSERT_TRUE(child_local(filled, "ln"));
  EXPECT_FALSE(child_local(child_local(filled, "ln"), "solidFill"));
  const pugi::xml_node bare = child_local(anchor_content(tops[1]), "spPr");
  EXPECT_FALSE(child_local(bare, "noFill"));
  EXPECT_FALSE(child_local(bare, "ln"));
  EXPECT_TRUE(child_local(bare, "prstGeom"));
}

}  // namespace
}  // namespace formulon
