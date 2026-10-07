#include <algorithm>
#include <cstdint>
#include <string>
#include <vector>

#include "drawing/drawing_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "unknown_relationship.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kCtDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";
constexpr const char* kRelDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
constexpr const char* kImageRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
constexpr const char* kDrawingPath = "xl/drawings/drawing1.xml";
constexpr const char* kRelsPath = "xl/drawings/_rels/drawing1.xml.rels";

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

std::string Marker(const char* tag, std::uint32_t col, std::int64_t col_off, std::uint32_t row, std::int64_t row_off) {
  return std::string("<xdr:") + tag + "><xdr:col>" + std::to_string(col) + "</xdr:col><xdr:colOff>" +
         std::to_string(col_off) + "</xdr:colOff><xdr:row>" + std::to_string(row) + "</xdr:row><xdr:rowOff>" +
         std::to_string(row_off) + "</xdr:rowOff></xdr:" + tag + ">";
}

std::string Rel(const char* id, const char* target) {
  return std::string("<Relationship Id=\"") + id + "\" Type=\"" + kImageRel + "\" Target=\"" + target + "\"/>";
}

const PassthroughPart* Part(const Workbook& wb, const std::string& path) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      return &part;
    }
  }
  return nullptr;
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

// Picture `id` as listed on sheet 0, or an object with id 0 when absent.
DrawingObject Find(const Workbook& wb, std::uint32_t id) {
  for (const DrawingObject& obj : List(wb)) {
    if (obj.object_id == id) {
      return obj;
    }
  }
  return DrawingObject{};
}

struct StringWriter : pugi::xml_writer {
  std::string out;
  void write(const void* data, std::size_t size) override { out.append(static_cast<const char*>(data), size); }
};

// The top-level element holding picture `id` in sheet `sheet`'s drawing, as
// the restore contract compares it: relationship ids read as the type and
// target they name, and the element's own namespace declarations dropped (a
// restore onto a root lacking one must carry it).
std::string TopXml(const Workbook& wb, std::size_t sheet, std::uint32_t id) {
  const std::string path = wb.sheet(sheet).drawing_rel_target();
  const PassthroughPart* part = Part(wb, path);
  const PassthroughPart* rels_part = Part(wb, "xl/drawings/_rels/" + path.substr(path.rfind('/') + 1) + ".rels");
  if (part == nullptr || rels_part == nullptr) {
    ADD_FAILURE() << "no drawing on sheet " << sheet;
    return {};
  }
  auto rels = io::ooxml::parse_part_rels(rels_part->bytes, "xl/drawings");
  EXPECT_TRUE(static_cast<bool>(rels));
  pugi::xml_document doc;
  EXPECT_TRUE(static_cast<bool>(parse_drawing_part(part->bytes, doc)));
  const pugi::xml_node root = doc.document_element();
  for (const pugi::xml_node& anchor : drawing_anchors(root, false)) {
    const DrawingObject obj = read_drawing_object(anchor, {});
    if (obj.kind != DrawingObjectKind::kPicture || obj.object_id != id) {
      continue;
    }
    pugi::xml_node top = anchor;
    while (top.parent() != root) {
      top = top.parent();
    }
    pugi::xml_document copy;
    pugi::xml_node out = copy.append_copy(top);
    for (pugi::xml_attribute attr = out.first_attribute(); attr;) {
      const pugi::xml_attribute next = attr.next_attribute();
      if (std::string(attr.name()).rfind("xmlns", 0) == 0) {
        out.remove_attribute(attr);
      }
      attr = next;
    }
    std::vector<pugi::xml_node> stack = {out};
    while (!stack.empty()) {
      pugi::xml_node node = stack.back();
      stack.pop_back();
      for (pugi::xml_attribute attr = node.first_attribute(); attr; attr = attr.next_attribute()) {
        const std::string name = attr.name();
        if (name.find(':') == std::string::npos || name.rfind("xmlns:", 0) == 0) {
          continue;
        }
        for (const io::ooxml::DrawingRel& rel : rels.value()) {
          if (rel.id == attr.value()) {
            attr.set_value(("rel:" + rel.type + "|" + rel.target).c_str());
            break;
          }
        }
      }
      for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
        if (child.type() == pugi::node_element) {
          stack.push_back(child);
        }
      }
    }
    StringWriter writer;
    out.print(writer, "", pugi::format_raw);
    return writer.out;
  }
  ADD_FAILURE() << "no picture " << id;
  return {};
}

pugi::xml_node RestoredTop(const pugi::xml_document& doc, std::uint32_t id) {
  const pugi::xml_node root = doc.document_element();
  for (const pugi::xml_node& anchor : drawing_anchors(root, false)) {
    if (read_drawing_object(anchor, {}).object_id == id) {
      pugi::xml_node top = anchor;
      while (top.parent() != root) {
        top = top.parent();
      }
      return top;
    }
  }
  return {};
}

// Picture 2 (rotated, cropped, with a creationId), shape 3 and picture 4. The
// root declares `r`, so the pictures' `r:embed` relies on it.
Workbook Fixture() {
  const std::string pic2 =
      "<xdr:twoCellAnchor editAs=\"oneCell\">" + Marker("from", 1, 317500, 2, 127000) +
      Marker("to", 3, 444500, 6, 127000) +
      "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"2\" name=\"PicA\"><a:extLst><a:ext "
      "uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\"><a16:creationId "
      "xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" "
      "id=\"{916FAF09-56A5-F8BE-D4D0-73AC940BFF01}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvPicPr><a:picLocks/>"
      "</xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rId1\"/><a:srcRect l=\"1000\" r=\"500\"/>"
      "<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm rot=\"5400000\" flipV=\"1\">"
      "<a:off x=\"1270000\" y=\"635000\"/><a:ext cx=\"2032000\" cy=\"1016000\"/></a:xfrm><a:prstGeom prst=\"rect\">"
      "<a:avLst/></a:prstGeom></xdr:spPr><unknown:x xmlns:unknown=\"urn:unknown\" keep=\"1\"/></xdr:pic>"
      "<xdr:clientData fLocksWithSheet=\"0\"/></xdr:twoCellAnchor>";
  const std::string shape3 =
      "<xdr:oneCellAnchor>" + Marker("from", 8, 0, 1, 0) +
      "<xdr:ext cx=\"914400\" cy=\"457200\"/><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"3\" "
      "name=\"Oval 2\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>";
  const std::string pic4 =
      "<xdr:oneCellAnchor>" + Marker("from", 4, 0, 10, 0) +
      "<xdr:ext cx=\"95250\" cy=\"95250\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"4\" name=\"Small\"/>"
      "<xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rId2\"/></xdr:blipFill><xdr:spPr><a:xfrm>"
      "<a:off x=\"0\" y=\"0\"/><a:ext cx=\"95250\" cy=\"95250\"/></a:xfrm></xdr:spPr></xdr:pic><xdr:clientData/>"
      "</xdr:oneCellAnchor>";
  const std::string drawing =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
      pic2 + shape3 + pic4 + "</xdr:wsDr>";
  const std::string rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
      Rel("rId1", "../media/image1.png") + Rel("rId2", "../media/image2.png") + "</Relationships>";
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kDrawingPath, kCtDrawing, ToBytes(drawing)))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kRelsPath, "", ToBytes(rels)))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/image1.png", "", Png(40, 80)))));
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/image2.png", "", Png(10, 10)))));
  wb.set_default_content_types({{"png", "image/png"}});
  wb.sheet(0).set_drawing_rel_target(kDrawingPath);
  return wb;
}

Bytes Snapshot(const Workbook& wb, std::uint32_t id) {
  auto bytes = snapshot_image(wb, 0, id);
  EXPECT_TRUE(static_cast<bool>(bytes));
  return bytes ? bytes.value() : Bytes{};
}

Expected<std::uint32_t, Error> Restore(Workbook& wb, std::size_t sheet, const Bytes& bytes, std::uint32_t flags = 0) {
  return restore_image(wb, sheet, bytes.data(), bytes.size(), flags);
}

Workbook RoundTrip(const Workbook& wb) {
  auto bytes = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes));
  auto read = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  EXPECT_TRUE(static_cast<bool>(read));
  return std::move(read.value().workbook);
}

TEST(DrawingSnapshot, DeleteThenRestoreReturnsTheSamePicture) {
  for (const std::uint32_t id : {2U, 4U}) {
    SCOPED_TRACE(id);
    Workbook wb = Fixture();
    const std::string before = TopXml(wb, 0, id);
    const std::string media = Find(wb, id).media_path;
    const Bytes media_bytes = Part(wb, media)->bytes;
    const Bytes snap = Snapshot(wb, id);
    ASSERT_TRUE(static_cast<bool>(remove_image(wb, 0, id)));
    ASSERT_EQ(Part(wb, media), nullptr);

    auto restored = Restore(wb, 0, snap);
    ASSERT_TRUE(static_cast<bool>(restored));
    EXPECT_EQ(restored.value(), id);
    EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{2, 3, 4}));
    EXPECT_EQ(TopXml(wb, 0, id), before);
    const DrawingObject obj = Find(wb, id);
    ASSERT_EQ(obj.object_id, id);
    EXPECT_EQ(obj.media_path, media);
    EXPECT_EQ(Part(wb, media)->bytes, media_bytes);
    // The root declares `r`, so the restored element does not repeat it.
    pugi::xml_document doc;
    ASSERT_TRUE(static_cast<bool>(parse_drawing_part(Part(wb, kDrawingPath)->bytes, doc)));
    EXPECT_FALSE(RestoredTop(doc, id).attribute("xmlns:r"));
  }
}

TEST(DrawingSnapshot, MoveAndReorderThenRestore) {
  Workbook wb = Fixture();
  const std::string before = TopXml(wb, 0, 4);
  const DrawingObject listed = Find(wb, 4);
  const Bytes snap = Snapshot(wb, 4);
  ImageAnchor moved;
  moved.anchor_kind = AnchorKind::kTwoCell;
  moved.row = 20;
  moved.col = 6;
  moved.width_emu = 400000;
  ASSERT_TRUE(static_cast<bool>(set_image_anchor(wb, 0, 4, moved)));
  ASSERT_TRUE(static_cast<bool>(set_image_z_order(wb, 0, 4, 0)));
  ASSERT_EQ(Ids(wb), (std::vector<std::uint32_t>{4, 2, 3}));

  auto restored = Restore(wb, 0, snap);
  ASSERT_TRUE(static_cast<bool>(restored));
  EXPECT_EQ(restored.value(), 4U);
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{2, 3, 4}));
  EXPECT_EQ(TopXml(wb, 0, 4), before);
  const DrawingObject obj = Find(wb, 4);
  ASSERT_EQ(obj.object_id, 4U);
  EXPECT_EQ(obj.anchor_kind, listed.anchor_kind);
  EXPECT_EQ(obj.from.row, listed.from.row);
  EXPECT_EQ(obj.from.col, listed.from.col);
  EXPECT_EQ(obj.cx, listed.cx);
  EXPECT_EQ(Part(wb, "xl/media/image2.png")->bytes, Png(10, 10));

  // Restoring the bottom picture after it moved to the top.
  const std::string bottom = TopXml(wb, 0, 2);
  const Bytes snap2 = Snapshot(wb, 2);
  ASSERT_TRUE(static_cast<bool>(set_image_z_order(wb, 0, 2, 2)));
  ASSERT_TRUE(static_cast<bool>(Restore(wb, 0, snap2)));
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{2, 3, 4}));
  EXPECT_EQ(TopXml(wb, 0, 2), bottom);
}

TEST(DrawingSnapshot, NewIdAddsACopyOnTop) {
  Workbook wb = Fixture();
  const Bytes rels_before = Part(wb, kRelsPath)->bytes;
  const std::size_t parts_before = wb.passthrough_parts().size();
  const Bytes snap = Snapshot(wb, 2);
  auto copy = Restore(wb, 0, snap, kImageRestoreNewId);
  ASSERT_TRUE(static_cast<bool>(copy));
  EXPECT_EQ(copy.value(), 5U);
  EXPECT_EQ(Ids(wb), (std::vector<std::uint32_t>{2, 3, 4, 5}));
  const std::vector<DrawingObject> objects = List(wb);
  EXPECT_EQ(objects[3].name, "PicA");
  EXPECT_EQ(objects[3].media_path, objects[0].media_path);
  // The media and its relationship are shared, not duplicated.
  EXPECT_EQ(wb.passthrough_parts().size(), parts_before);
  EXPECT_EQ(Part(wb, kRelsPath)->bytes, rels_before);
  const std::string drawing(Part(wb, kDrawingPath)->bytes.begin(), Part(wb, kDrawingPath)->bytes.end());
  EXPECT_EQ(drawing.find("creationId"), drawing.rfind("creationId"));
  pugi::xml_document doc;
  ASSERT_TRUE(static_cast<bool>(parse_drawing_part(Part(wb, kDrawingPath)->bytes, doc)));
  const pugi::xml_node cnvpr = anchor_cnvpr(drawing_anchors(doc.document_element(), false)[3]);
  EXPECT_EQ(cnvpr.attribute("id").as_uint(), 5U);
  EXPECT_FALSE(child_local(cnvpr, "extLst"));
  const pugi::xml_node original = anchor_cnvpr(drawing_anchors(doc.document_element(), false)[0]);
  EXPECT_TRUE(child_local(original, "extLst"));
}

TEST(DrawingSnapshot, RestoresOntoAnotherSheet) {
  Workbook wb = Fixture();
  const std::size_t other = wb.add_sheet("Other");
  const std::string before = TopXml(wb, 0, 2);
  const Bytes snap = Snapshot(wb, 2);
  auto restored = Restore(wb, other, snap);
  ASSERT_TRUE(static_cast<bool>(restored));
  EXPECT_EQ(restored.value(), 2U);
  const std::string path = wb.sheet(other).drawing_rel_target();
  EXPECT_EQ(path, "xl/drawings/drawing2.xml");
  ASSERT_NE(Part(wb, path), nullptr);
  EXPECT_EQ(Part(wb, path)->content_type, kCtDrawing);
  EXPECT_EQ(TopXml(wb, other, 2), before);
  const std::vector<DrawingObject> objects = List(wb, other);
  ASSERT_EQ(objects.size(), 1U);
  EXPECT_EQ(objects[0].media_path, "xl/media/image1.png");
  EXPECT_EQ(objects[0].image_format, ImageFormat::kPng);
  // The fresh root does not declare `r`, so the element carries it.
  pugi::xml_document doc;
  ASSERT_TRUE(static_cast<bool>(parse_drawing_part(Part(wb, path)->bytes, doc)));
  EXPECT_STREQ(RestoredTop(doc, 2).attribute("xmlns:r").value(),
               "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

  Workbook reloaded = RoundTrip(wb);
  const std::vector<DrawingObject> again = List(reloaded, other);
  ASSERT_EQ(again.size(), 1U);
  EXPECT_EQ(again[0].object_id, 2U);
  EXPECT_EQ(Part(reloaded, again[0].media_path)->bytes, Png(40, 80));
  EXPECT_EQ(Ids(reloaded, 0), (std::vector<std::uint32_t>{2, 3, 4}));
}

TEST(DrawingSnapshot, DifferentBytesAtThePathTakeAnotherPath) {
  Workbook wb = Fixture();
  const Bytes snap = Snapshot(wb, 2);
  ASSERT_TRUE(static_cast<bool>(remove_image(wb, 0, 2)));
  const Bytes imposter = Png(5, 5);
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/image1.png", "", imposter))));
  ASSERT_TRUE(static_cast<bool>(Restore(wb, 0, snap)));
  const DrawingObject obj = Find(wb, 2);
  ASSERT_EQ(obj.object_id, 2U);
  ASSERT_EQ(obj.media_path, "xl/media/image3.png");
  EXPECT_EQ(Part(wb, obj.media_path)->bytes, Png(40, 80));
  EXPECT_EQ(Part(wb, "xl/media/image1.png")->bytes, imposter);
  const std::string rels(Part(wb, kRelsPath)->bytes.begin(), Part(wb, kRelsPath)->bytes.end());
  EXPECT_NE(rels.find("Target=\"../media/image3.png\""), std::string::npos);
}

TEST(DrawingSnapshot, SaveReloadKeepsTheList) {
  Workbook wb = Fixture();
  const Bytes snap = Snapshot(wb, 2);
  ASSERT_TRUE(static_cast<bool>(remove_image(wb, 0, 2)));
  ASSERT_TRUE(static_cast<bool>(Restore(wb, 0, snap)));
  ASSERT_TRUE(static_cast<bool>(Restore(wb, 0, snap, kImageRestoreNewId)));
  const std::vector<DrawingObject> before = List(wb);
  const std::vector<DrawingObject> after = List(RoundTrip(wb));
  ASSERT_EQ(after.size(), before.size());
  for (std::size_t i = 0; i < before.size(); ++i) {
    SCOPED_TRACE(i);
    EXPECT_EQ(after[i].object_id, before[i].object_id);
    EXPECT_EQ(after[i].kind, before[i].kind);
    EXPECT_EQ(after[i].anchor_kind, before[i].anchor_kind);
    EXPECT_EQ(after[i].edit_as, before[i].edit_as);
    EXPECT_EQ(after[i].from.row, before[i].from.row);
    EXPECT_EQ(after[i].from.col_off, before[i].from.col_off);
    EXPECT_EQ(after[i].to.col, before[i].to.col);
    EXPECT_EQ(after[i].cx, before[i].cx);
    EXPECT_EQ(after[i].cy, before[i].cy);
    EXPECT_EQ(after[i].name, before[i].name);
    EXPECT_EQ(after[i].media_path, before[i].media_path);
  }
}

void PutU32(Bytes& out, std::uint32_t v) {
  for (int shift = 0; shift < 32; shift += 8) {
    out.push_back(static_cast<std::uint8_t>(v >> shift));
  }
}

Bytes Handmade(std::uint32_t version, const std::string& xml) {
  Bytes out = {'F', 'M', 'I', 'S'};
  PutU32(out, version);
  PutU32(out, 0);
  PutU32(out, static_cast<std::uint32_t>(xml.size()));
  out.insert(out.end(), xml.begin(), xml.end());
  PutU32(out, 0);
  return out;
}

TEST(DrawingSnapshot, RejectsMalformedBytes) {
  Workbook wb = Fixture();
  const Bytes snap = Snapshot(wb, 4);
  ASSERT_TRUE(static_cast<bool>(remove_image(wb, 0, 4)));
  const std::vector<PassthroughPart> parts = wb.passthrough_parts();
  const auto code = [&wb](const Bytes& bytes, std::uint32_t flags = 0) {
    auto result = Restore(wb, 0, bytes, flags);
    return result ? FormulonErrorCode::kOk : result.error().code;
  };
  Bytes magic = snap;
  magic[0] = 'X';
  EXPECT_EQ(code(magic), FormulonErrorCode::kInvalidArgument);
  Bytes version = snap;
  version[4] = 2;
  EXPECT_EQ(code(version), FormulonErrorCode::kInvalidArgument);
  for (std::size_t len = 0; len < snap.size(); ++len) {
    const Bytes cut(snap.begin(), snap.begin() + static_cast<std::ptrdiff_t>(len));
    ASSERT_EQ(code(cut), FormulonErrorCode::kInvalidArgument) << len;
  }
  Bytes trailing = snap;
  trailing.push_back(0);
  EXPECT_EQ(code(trailing), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(code(snap, 2), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(code(Handmade(1, "<xdr:oneCellAnchor><xdr:from>")), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(code(Handmade(1,
                          "<xdr:oneCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"9\"/></xdr:nvSpPr></xdr:sp>"
                          "</xdr:oneCellAnchor>")),
            FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Restore(wb, 0, Bytes{}).error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(restore_image(wb, 0, nullptr, 0, 0).error().code, FormulonErrorCode::kInvalidArgument);
  ASSERT_EQ(wb.passthrough_parts().size(), parts.size());
  for (std::size_t i = 0; i < parts.size(); ++i) {
    EXPECT_EQ(wb.passthrough_parts()[i].bytes, parts[i].bytes) << parts[i].path;
  }
  // The intact capture still restores.
  EXPECT_TRUE(static_cast<bool>(Restore(wb, 0, snap)));
}

TEST(DrawingSnapshot, RefusesWhatItCannotCaptureOrPlace) {
  Workbook wb = Fixture();
  const auto snap_code = [&wb](std::uint32_t id) {
    auto result = snapshot_image(wb, 0, id);
    return result ? FormulonErrorCode::kOk : result.error().code;
  };
  EXPECT_EQ(snap_code(3), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(snap_code(9), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(snapshot_image(wb, 4, 2).error().code, FormulonErrorCode::kInvalidArgument);
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(
      "xl/media/_rels/image1.png.rels", "",
      ToBytes("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>")))));
  EXPECT_EQ(snap_code(2), FormulonErrorCode::kIoDrawingUnparseable);
  const Bytes snap4 = Snapshot(wb, 4);

  // Id 3 belongs to a shape on the target sheet.
  Bytes as_three = snap4;
  const std::string needle = "id=\"4\"";
  const auto at = std::search(as_three.begin(), as_three.end(), needle.begin(), needle.end());
  ASSERT_NE(at, as_three.end());
  *(at + 4) = '3';
  auto clash = Restore(wb, 0, as_three);
  ASSERT_FALSE(static_cast<bool>(clash));
  EXPECT_EQ(clash.error().code, FormulonErrorCode::kInvalidArgument);

  Workbook xlsb = Workbook::create();
  const Bytes drawing = ToBytes(
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/"
      "spreadsheetDrawing\"/>");
  ASSERT_TRUE(static_cast<bool>(xlsb.add_passthrough_part(PassthroughPart(kDrawingPath, kCtDrawing, drawing))));
  xlsb.sheet(0).set_unknown_relationships({UnknownRelationship{"rId1", kRelDrawing, kDrawingPath, false}});
  auto refused = Restore(xlsb, 0, snap4);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kIoDrawingUnparseable);
  EXPECT_EQ(xlsb.passthrough_parts().size(), 1U);
}

}  // namespace
}  // namespace formulon
