// Connectors glued to shapes following an edit of those shapes: both ends glued route again between the two sites,
// one end glued moves the whole connector. The expected boxes are what Mac Excel 16.115 saves after the same edit.

#include <algorithm>
#include <cstdint>
#include <ostream>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "drawing/shape_edit.h"
#include "gtest/gtest.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "workbook.h"

namespace formulon {
namespace {

constexpr const char* kCtDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";
constexpr const char* kDrawingPath = "xl/drawings/drawing1.xml";
constexpr std::int64_t kTop = 2540000;

struct Box {
  std::int64_t x, y, w, h;
};

// A connector's `a:xfrm`: its box and flips.
struct Xfrm {
  std::int64_t x, y, cx, cy;
  bool flip_h, flip_v;
  bool operator==(const Xfrm& o) const {
    return x == o.x && y == o.y && cx == o.cx && cy == o.cy && flip_h == o.flip_h && flip_v == o.flip_v;
  }
};

std::ostream& operator<<(std::ostream& out, const Xfrm& x) {
  return out << "{" << x.x << ", " << x.y << ", " << x.cx << ", " << x.cy << (x.flip_h ? ", flipH" : "")
             << (x.flip_v ? ", flipV" : "") << "}";
}

std::string Num(std::int64_t v) {
  return std::to_string(v);
}

std::string XfrmXml(const Xfrm& x, const std::string& rot = "") {
  return "<a:xfrm" + rot + (x.flip_h ? " flipH=\"1\"" : "") + (x.flip_v ? " flipV=\"1\"" : "") + "><a:off x=\"" +
         Num(x.x) + "\" y=\"" + Num(x.y) + "\"/><a:ext cx=\"" + Num(x.cx) + "\" cy=\"" + Num(x.cy) + "\"/></a:xfrm>";
}

std::string Absolute(const Box& b, const std::string& content) {
  return "<xdr:absoluteAnchor><xdr:pos x=\"" + Num(b.x) + "\" y=\"" + Num(b.y) + "\"/><xdr:ext cx=\"" + Num(b.w) +
         "\" cy=\"" + Num(b.h) + "\"/>" + content + "<xdr:clientData/></xdr:absoluteAnchor>";
}

// A shape of preset `prst` in box `b`; `transform` holds `rot` / flip attributes, `guides` its `a:avLst` children.
std::string Sp(std::uint32_t id, const char* prst, const Box& b, const std::string& transform = "",
               const std::string& guides = "") {
  return "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"" + Num(id) + "\" name=\"Shape " + Num(id) +
         "\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm" + transform + "><a:off x=\"" + Num(b.x) + "\" y=\"" +
         Num(b.y) + "\"/><a:ext cx=\"" + Num(b.w) + "\" cy=\"" + Num(b.h) + "\"/></a:xfrm><a:prstGeom prst=\"" + prst +
         "\"><a:avLst>" + guides + "</a:avLst></a:prstGeom></xdr:spPr></xdr:sp>";
}

std::string Glue(const char* tag, std::uint32_t id, std::uint32_t idx) {
  return std::string("<a:") + tag + " id=\"" + Num(id) + "\" idx=\"" + Num(idx) + "\"/>";
}

std::string Cxn(std::uint32_t id, const char* prst, const Xfrm& x, const std::string& glue,
                const std::string& guides = "") {
  return "<xdr:cxnSp macro=\"\"><xdr:nvCxnSpPr><xdr:cNvPr id=\"" + Num(id) + "\" name=\"Connector " + Num(id) +
         "\"/><xdr:cNvCxnSpPr>" + glue + "</xdr:cNvCxnSpPr></xdr:nvCxnSpPr><xdr:spPr>" + XfrmXml(x) +
         "<a:prstGeom prst=\"" + prst + "\"><a:avLst>" + guides + "</a:avLst></a:prstGeom></xdr:spPr></xdr:cxnSp>";
}

Box BoxOf(const Xfrm& x) {
  return Box{x.x, x.y, x.cx, x.cy};
}

Workbook Book(const std::string& body) {
  const std::string xml =
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
      body + "</xdr:wsDr>";
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(
      PassthroughPart(kDrawingPath, kCtDrawing, std::vector<std::uint8_t>(xml.begin(), xml.end())))));
  wb.sheet(0).set_drawing_rel_target(kDrawingPath);
  return wb;
}

pugi::xml_node Find(const pugi::xml_document& doc, std::uint32_t id) {
  return doc.document_element().find_node([id](const pugi::xml_node& n) {
    return n.child("xdr:nvCxnSpPr").child("xdr:cNvPr").attribute("id").as_uint() == id ||
           n.child("xdr:nvSpPr").child("xdr:cNvPr").attribute("id").as_uint() == id;
  });
}

void Parse(const Workbook& wb, pugi::xml_document& doc) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == kDrawingPath) {
      ASSERT_TRUE(static_cast<bool>(parse_drawing_part(part.bytes, doc)));
    }
  }
}

Xfrm ConnectorXfrm(const Workbook& wb, std::uint32_t id) {
  pugi::xml_document doc;
  Parse(wb, doc);
  const pugi::xml_node x = Find(doc, id).child("xdr:spPr").child("a:xfrm");
  return Xfrm{x.child("a:off").attribute("x").as_llong(),
              x.child("a:off").attribute("y").as_llong(),
              x.child("a:ext").attribute("cx").as_llong(),
              x.child("a:ext").attribute("cy").as_llong(),
              x.attribute("flipH").as_bool(),
              x.attribute("flipV").as_bool()};
}

std::int32_t ConnectorRotation(const Workbook& wb, std::uint32_t id) {
  pugi::xml_document doc;
  Parse(wb, doc);
  return Find(doc, id).child("xdr:spPr").child("a:xfrm").attribute("rot").as_int();
}

std::string GeometryXml(const Workbook& wb, std::uint32_t id) {
  pugi::xml_document doc;
  Parse(wb, doc);
  std::ostringstream out;
  Find(doc, id).child("xdr:spPr").child("a:prstGeom").print(out, "", pugi::format_raw);
  return out.str();
}

void Move(Workbook& wb, std::uint32_t id, const Box& b) {
  ShapeSpec spec;
  spec.fields = shape_field::kAnchor;
  spec.anchor_kind = AnchorKind::kAbsolute;
  spec.from.col_off = b.x;
  spec.from.row_off = b.y;
  spec.cx = b.w;
  spec.cy = b.h;
  ASSERT_TRUE(static_cast<bool>(update_shape(wb, 0, id, spec)));
}

// Shape 2's site 3 (right middle) is the start at (3048000, kTop); shape 3's site 1 (left middle) the end.
const Box kStartShape{2032000, kTop - 254000, 1016000, 508000};
const Box kEndShape{4064000, kTop + 508000, 1016000, 508000};
const Xfrm kBothRoute{3048000, kTop, 1016000, 762000, false, false};

struct MoveRow {
  const char* op;
  Box end_shape;
  Xfrm expected;
};

const MoveRow kMoves[] = {
    {"move", {4699000, kTop + 889000, 1016000, 508000}, {3048000, kTop, 1651000, 1143000, false, false}},
    {"resize", {4064000, kTop + 508000, 1778000, 1016000}, {3048000, kTop, 1016000, 1016000, false, false}},
    {"cross", {254000, kTop + 1270000, 1016000, 508000}, {254000, kTop, 2794000, 1524000, true, false}},
};

TEST(ConnectorRoute, AStraightConnectorGluedAtBothEndsRunsBetweenTheSites) {
  for (const MoveRow& row : kMoves) {
    Workbook wb =
        Book(Absolute(kStartShape, Sp(2, "rect", kStartShape)) + Absolute(kEndShape, Sp(3, "rect", kEndShape)) +
             Absolute(BoxOf(kBothRoute),
                      Cxn(4, "straightConnector1", kBothRoute, Glue("stCxn", 2, 3) + Glue("endCxn", 3, 1))));
    Move(wb, 3, row.end_shape);
    EXPECT_EQ(ConnectorXfrm(wb, 4), row.expected) << row.op;
    // The anchor follows the new box.
    const DrawingObject obj = list_drawing_objects(wb, 0).value()[2];
    EXPECT_EQ(obj.from.col_off, row.expected.x) << row.op;
    EXPECT_EQ(obj.from.row_off, row.expected.y) << row.op;
  }
}

// A bent or curved connector whose shapes move keeps its route while the arrangement holds; once the end shape
// crosses to the other side, the connector goes around the start shape.
TEST(ConnectorRoute, BentAndCurvedConnectorsFollowTheArrangement) {
  const char* families[] = {"bentConnector", "curvedConnector"};
  for (const char* family : families) {
    const std::string three = std::string(family) + "3";
    for (const MoveRow& row : kMoves) {
      Workbook wb =
          Book(Absolute(kStartShape, Sp(2, "rect", kStartShape)) + Absolute(kEndShape, Sp(3, "rect", kEndShape)) +
               Absolute(BoxOf(kBothRoute), Cxn(4, three.c_str(), kBothRoute, Glue("stCxn", 2, 3) + Glue("endCxn", 3, 1),
                                               "<a:gd name=\"adj1\" fmla=\"val 50000\"/>")));
      Move(wb, 3, row.end_shape);
      EXPECT_EQ(ConnectorXfrm(wb, 4), row.expected) << family << " " << row.op;
      const bool crossed = std::string(row.op) == "cross";
      EXPECT_EQ(GeometryXml(wb, 4), std::string("<a:prstGeom prst=\"") + family + (crossed ? "5" : "3") +
                                        "\"><a:avLst><a:gd name=\"adj1\" fmla=\"val " +
                                        (crossed ? "-8182\"/><a:gd name=\"adj2\" fmla=\"val 50000\"/><a:gd "
                                                   "name=\"adj3\" fmla=\"val 108182"
                                                 : "50000") +
                                        "\"/></a:avLst></a:prstGeom>")
          << family << " " << row.op;
    }
  }
}

struct MatrixRow {
  const char* family;
  bool far;
  const char* quadrant;
  std::uint32_t start_site, end_site;
  int segments;
  std::vector<std::int32_t> adjust;
  int quarter_turns;
  bool flip_h, flip_v;
  Box box;
};

// Mac Excel 16.115 routes between two 80 x 40 pt rectangles: shape 2 at (400, 300) pt, shape 3 diagonally off it
// with a 20 pt (near) or 200 x 150 pt (far) gap, glued at every pair of their four sites.
const MatrixRow kMatrix[] = {
    {"elbow", false, "NE", 0, 0, 3, {130000}, 1, true, true, {5842000, 2794000, 762000, 1270000}},
    {"elbow", false, "NE", 0, 1, 2, {}, 1, true, true, {5715000, 3175000, 508000, 762000}},
    {"elbow", false, "NE", 0, 2, 3, {50000}, 1, true, true, {6096000, 3048000, 254000, 1270000}},
    {"elbow", false, "NE", 0, 3, 4, {25000, 112857}, 1, true, true, {6223000, 2667000, 508000, 1778000}},
    {"elbow", false, "NE", 1, 0, 4, {-12857, 122500}, 2, true, false, {5080000, 3048000, 1778000, 1016000}},
    {"elbow", false, "NE", 1, 1, 3, {-18000}, 2, true, false, {5080000, 3302000, 1270000, 762000}},
    {"elbow", false, "NE", 1, 2, 4, {-12857, 75000}, 2, true, false, {5080000, 3556000, 1778000, 508000}},
    {"elbow", false, "NE", 1, 3, 5, {-10000, 50000, 110000}, 2, true, false, {5080000, 3302000, 2286000, 762000}},
    {"elbow", false, "NE", 2, 0, 5, {-18000, 50000, 118000}, 1, true, true, {5588000, 3048000, 1270000, 1270000}},
    {"elbow", false, "NE", 2, 1, 4, {-22500, 83333}, 1, true, true, {5461000, 3429000, 1016000, 762000}},
    {"elbow", false, "NE", 2, 2, 3, {-30000}, 1, true, true, {5842000, 3302000, 762000, 1270000}},
    {"elbow", false, "NE", 2, 3, 4, {-22500, 112857}, 1, true, true, {5969000, 2921000, 1016000, 1778000}},
    {"elbow", false, "NE", 3, 0, 4, {16667, 122500}, 0, false, true, {6096000, 3048000, 762000, 1016000}},
    {"elbow", false, "NE", 3, 1, 3, {50000}, 0, false, true, {6096000, 3302000, 254000, 762000}},
    {"elbow", false, "NE", 3, 2, 2, {}, 0, false, true, {6096000, 3556000, 762000, 508000}},
    {"elbow", false, "NE", 3, 3, 3, {118000}, 0, false, true, {6096000, 3302000, 1270000, 762000}},
    {"elbow", false, "SE", 0, 0, 3, {-30000}, 3, true, false, {5842000, 3556000, 762000, 1270000}},
    {"elbow", false, "SE", 0, 1, 4, {-22500, 83333}, 3, true, false, {5461000, 3937000, 1016000, 762000}},
    {"elbow", false, "SE", 0, 2, 5, {-18000, 50000, 118000}, 3, true, false, {5588000, 3810000, 1270000, 1270000}},
    {"elbow", false, "SE", 0, 3, 4, {-22500, 112857}, 3, true, false, {5969000, 3429000, 1016000, 1778000}},
    {"elbow", false, "SE", 1, 0, 4, {-12857, 75000}, 2, true, true, {5080000, 4064000, 1778000, 508000}},
    {"elbow", false, "SE", 1, 1, 3, {-18000}, 2, true, true, {5080000, 4064000, 1270000, 762000}},
    {"elbow", false, "SE", 1, 2, 4, {-12857, 122500}, 2, true, true, {5080000, 4064000, 1778000, 1016000}},
    {"elbow", false, "SE", 1, 3, 5, {-10000, 50000, 110000}, 2, true, true, {5080000, 4064000, 2286000, 762000}},
    {"elbow", false, "SE", 2, 0, 3, {50000}, 3, true, false, {6096000, 3810000, 254000, 1270000}},
    {"elbow", false, "SE", 2, 1, 2, {}, 3, true, false, {5715000, 4191000, 508000, 762000}},
    {"elbow", false, "SE", 2, 2, 3, {130000}, 3, true, false, {5842000, 4064000, 762000, 1270000}},
    {"elbow", false, "SE", 2, 3, 4, {25000, 112857}, 3, true, false, {6223000, 3683000, 508000, 1778000}},
    {"elbow", false, "SE", 3, 0, 2, {}, 0, false, false, {6096000, 4064000, 762000, 508000}},
    {"elbow", false, "SE", 3, 1, 3, {50000}, 0, false, false, {6096000, 4064000, 254000, 762000}},
    {"elbow", false, "SE", 3, 2, 4, {16667, 122500}, 0, false, false, {6096000, 4064000, 762000, 1016000}},
    {"elbow", false, "SE", 3, 3, 3, {118000}, 0, false, false, {6096000, 4064000, 1270000, 762000}},
    {"elbow", false, "SW", 0, 0, 3, {-30000}, 3, true, true, {4572000, 3556000, 762000, 1270000}},
    {"elbow", false, "SW", 0, 1, 4, {-22500, 112857}, 3, true, true, {4191000, 3429000, 1016000, 1778000}},
    {"elbow", false, "SW", 0, 2, 5, {-18000, 50000, 118000}, 3, true, true, {4318000, 3810000, 1270000, 1270000}},
    {"elbow", false, "SW", 0, 3, 4, {-22500, 83333}, 3, true, true, {4699000, 3937000, 1016000, 762000}},
    {"elbow", false, "SW", 1, 0, 2, {}, 2, false, true, {4318000, 4064000, 762000, 508000}},
    {"elbow", false, "SW", 1, 1, 3, {118000}, 2, false, true, {3810000, 4064000, 1270000, 762000}},
    {"elbow", false, "SW", 1, 2, 4, {16667, 122500}, 2, false, true, {4318000, 4064000, 762000, 1016000}},
    {"elbow", false, "SW", 1, 3, 3, {50000}, 2, false, true, {4826000, 4064000, 254000, 762000}},
    {"elbow", false, "SW", 2, 0, 3, {50000}, 1, false, false, {4826000, 3810000, 254000, 1270000}},
    {"elbow", false, "SW", 2, 1, 4, {25000, 112857}, 1, false, false, {4445000, 3683000, 508000, 1778000}},
    {"elbow", false, "SW", 2, 2, 3, {130000}, 1, false, false, {4572000, 4064000, 762000, 1270000}},
    {"elbow", false, "SW", 2, 3, 2, {}, 1, false, false, {4953000, 4191000, 508000, 762000}},
    {"elbow", false, "SW", 3, 0, 4, {-12857, 75000}, 0, true, false, {4318000, 4064000, 1778000, 508000}},
    {"elbow", false, "SW", 3, 1, 5, {-10000, 50000, 110000}, 0, true, false, {3810000, 4064000, 2286000, 762000}},
    {"elbow", false, "SW", 3, 2, 4, {-12857, 122500}, 0, true, false, {4318000, 4064000, 1778000, 1016000}},
    {"elbow", false, "SW", 3, 3, 3, {-18000}, 0, true, false, {4826000, 4064000, 1270000, 762000}},
    {"elbow", false, "NW", 0, 0, 3, {130000}, 3, false, true, {4572000, 2794000, 762000, 1270000}},
    {"elbow", false, "NW", 0, 1, 4, {25000, 112857}, 3, false, true, {4445000, 2667000, 508000, 1778000}},
    {"elbow", false, "NW", 0, 2, 3, {50000}, 3, false, true, {4826000, 3048000, 254000, 1270000}},
    {"elbow", false, "NW", 0, 3, 2, {}, 3, false, true, {4953000, 3175000, 508000, 762000}},
    {"elbow", false, "NW", 1, 0, 4, {16667, 122500}, 2, false, false, {4318000, 3048000, 762000, 1016000}},
    {"elbow", false, "NW", 1, 1, 3, {118000}, 2, false, false, {3810000, 3302000, 1270000, 762000}},
    {"elbow", false, "NW", 1, 2, 2, {}, 2, false, false, {4318000, 3556000, 762000, 508000}},
    {"elbow", false, "NW", 1, 3, 3, {50000}, 2, false, false, {4826000, 3302000, 254000, 762000}},
    {"elbow", false, "NW", 2, 0, 5, {-18000, 50000, 118000}, 1, true, false, {4318000, 3048000, 1270000, 1270000}},
    {"elbow", false, "NW", 2, 1, 4, {-22500, 112857}, 1, true, false, {4191000, 2921000, 1016000, 1778000}},
    {"elbow", false, "NW", 2, 2, 3, {-30000}, 1, true, false, {4572000, 3302000, 762000, 1270000}},
    {"elbow", false, "NW", 2, 3, 4, {-22500, 83333}, 1, true, false, {4699000, 3429000, 1016000, 762000}},
    {"elbow", false, "NW", 3, 0, 4, {-12857, 122500}, 0, true, true, {4318000, 3048000, 1778000, 1016000}},
    {"elbow", false, "NW", 3, 1, 5, {-10000, 50000, 110000}, 0, true, true, {3810000, 3302000, 2286000, 762000}},
    {"elbow", false, "NW", 3, 2, 4, {-12857, 75000}, 0, true, true, {4318000, 3556000, 1778000, 508000}},
    {"elbow", false, "NW", 3, 3, 3, {-18000}, 0, true, true, {4826000, 3302000, 1270000, 762000}},
    {"elbow", true, "NE", 0, 0, 3, {109474}, 1, true, true, {6159500, 825500, 2413000, 3556000}},
    {"elbow", true, "NE", 0, 1, 2, {}, 1, true, true, {6032500, 1206500, 2159000, 3048000}},
    {"elbow", true, "NE", 0, 2, 3, {50000}, 1, true, true, {6413500, 1079500, 1905000, 3556000}},
    {"elbow", true, "NE", 0, 3, 4, {44118, 105625}, 1, true, true, {6540500, 698500, 2159000, 4064000}},
    {"elbow", true, "NE", 1, 0, 4, {-5625, 108571}, 2, true, false, {5080000, 1397000, 4064000, 2667000}},
    {"elbow", true, "NE", 1, 1, 3, {-6429}, 2, true, false, {5080000, 1651000, 3556000, 2413000}},
    {"elbow", true, "NE", 1, 2, 4, {-5625, 55882}, 2, true, false, {5080000, 1905000, 4064000, 2159000}},
    {"elbow", true, "NE", 1, 3, 5, {-5000, 50000, 105000}, 2, true, false, {5080000, 1651000, 4572000, 2413000}},
    {"elbow", true, "NE", 2, 0, 5, {-7826, 50000, 107826}, 1, true, true, {5905500, 1079500, 2921000, 3556000}},
    {"elbow", true, "NE", 2, 1, 4, {-8571, 58333}, 1, true, true, {5778500, 1460500, 2667000, 3048000}},
    {"elbow", true, "NE", 2, 2, 3, {-9474}, 1, true, true, {6159500, 1333500, 2413000, 3556000}},
    {"elbow", true, "NE", 2, 3, 4, {-8571, 105625}, 1, true, true, {6286500, 952500, 2667000, 4064000}},
    {"elbow", true, "NE", 3, 0, 4, {41667, 108571}, 0, false, true, {6096000, 1397000, 3048000, 2667000}},
    {"elbow", true, "NE", 3, 1, 3, {50000}, 0, false, true, {6096000, 1651000, 2540000, 2413000}},
    {"elbow", true, "NE", 3, 2, 2, {}, 0, false, true, {6096000, 1905000, 3048000, 2159000}},
    {"elbow", true, "NE", 3, 3, 3, {106429}, 0, false, true, {6096000, 1651000, 3556000, 2413000}},
    {"elbow", true, "SE", 0, 0, 3, {-9474}, 3, true, false, {6159500, 3238500, 2413000, 3556000}},
    {"elbow", true, "SE", 0, 1, 4, {-8571, 58333}, 3, true, false, {5778500, 3619500, 2667000, 3048000}},
    {"elbow", true, "SE", 0, 2, 5, {-7826, 50000, 107826}, 3, true, false, {5905500, 3492500, 2921000, 3556000}},
    {"elbow", true, "SE", 0, 3, 4, {-8571, 105625}, 3, true, false, {6286500, 3111500, 2667000, 4064000}},
    {"elbow", true, "SE", 1, 0, 4, {-5625, 55882}, 2, true, true, {5080000, 4064000, 4064000, 2159000}},
    {"elbow", true, "SE", 1, 1, 3, {-6429}, 2, true, true, {5080000, 4064000, 3556000, 2413000}},
    {"elbow", true, "SE", 1, 2, 4, {-5625, 108571}, 2, true, true, {5080000, 4064000, 4064000, 2667000}},
    {"elbow", true, "SE", 1, 3, 5, {-5000, 50000, 105000}, 2, true, true, {5080000, 4064000, 4572000, 2413000}},
    {"elbow", true, "SE", 2, 0, 3, {50000}, 3, true, false, {6413500, 3492500, 1905000, 3556000}},
    {"elbow", true, "SE", 2, 1, 2, {}, 3, true, false, {6032500, 3873500, 2159000, 3048000}},
    {"elbow", true, "SE", 2, 2, 3, {109474}, 3, true, false, {6159500, 3746500, 2413000, 3556000}},
    {"elbow", true, "SE", 2, 3, 4, {44118, 105625}, 3, true, false, {6540500, 3365500, 2159000, 4064000}},
    {"elbow", true, "SE", 3, 0, 2, {}, 0, false, false, {6096000, 4064000, 3048000, 2159000}},
    {"elbow", true, "SE", 3, 1, 3, {50000}, 0, false, false, {6096000, 4064000, 2540000, 2413000}},
    {"elbow", true, "SE", 3, 2, 4, {41667, 108571}, 0, false, false, {6096000, 4064000, 3048000, 2667000}},
    {"elbow", true, "SE", 3, 3, 3, {106429}, 0, false, false, {6096000, 4064000, 3556000, 2413000}},
    {"elbow", true, "SW", 0, 0, 3, {-9474}, 3, true, true, {2603500, 3238500, 2413000, 3556000}},
    {"elbow", true, "SW", 0, 1, 4, {-8571, 105625}, 3, true, true, {2222500, 3111500, 2667000, 4064000}},
    {"elbow", true, "SW", 0, 2, 5, {-7826, 50000, 107826}, 3, true, true, {2349500, 3492500, 2921000, 3556000}},
    {"elbow", true, "SW", 0, 3, 4, {-8571, 58333}, 3, true, true, {2730500, 3619500, 2667000, 3048000}},
    {"elbow", true, "SW", 1, 0, 2, {}, 2, false, true, {2032000, 4064000, 3048000, 2159000}},
    {"elbow", true, "SW", 1, 1, 3, {106429}, 2, false, true, {1524000, 4064000, 3556000, 2413000}},
    {"elbow", true, "SW", 1, 2, 4, {41667, 108571}, 2, false, true, {2032000, 4064000, 3048000, 2667000}},
    {"elbow", true, "SW", 1, 3, 3, {50000}, 2, false, true, {2540000, 4064000, 2540000, 2413000}},
    {"elbow", true, "SW", 2, 0, 3, {50000}, 1, false, false, {2857500, 3492500, 1905000, 3556000}},
    {"elbow", true, "SW", 2, 1, 4, {44118, 105625}, 1, false, false, {2476500, 3365500, 2159000, 4064000}},
    {"elbow", true, "SW", 2, 2, 3, {109474}, 1, false, false, {2603500, 3746500, 2413000, 3556000}},
    {"elbow", true, "SW", 2, 3, 2, {}, 1, false, false, {2984500, 3873500, 2159000, 3048000}},
    {"elbow", true, "SW", 3, 0, 4, {-5625, 55882}, 0, true, false, {2032000, 4064000, 4064000, 2159000}},
    {"elbow", true, "SW", 3, 1, 5, {-5000, 50000, 105000}, 0, true, false, {1524000, 4064000, 4572000, 2413000}},
    {"elbow", true, "SW", 3, 2, 4, {-5625, 108571}, 0, true, false, {2032000, 4064000, 4064000, 2667000}},
    {"elbow", true, "SW", 3, 3, 3, {-6429}, 0, true, false, {2540000, 4064000, 3556000, 2413000}},
    {"elbow", true, "NW", 0, 0, 3, {109474}, 3, false, true, {2603500, 825500, 2413000, 3556000}},
    {"elbow", true, "NW", 0, 1, 4, {44118, 105625}, 3, false, true, {2476500, 698500, 2159000, 4064000}},
    {"elbow", true, "NW", 0, 2, 3, {50000}, 3, false, true, {2857500, 1079500, 1905000, 3556000}},
    {"elbow", true, "NW", 0, 3, 2, {}, 3, false, true, {2984500, 1206500, 2159000, 3048000}},
    {"elbow", true, "NW", 1, 0, 4, {41667, 108571}, 2, false, false, {2032000, 1397000, 3048000, 2667000}},
    {"elbow", true, "NW", 1, 1, 3, {106429}, 2, false, false, {1524000, 1651000, 3556000, 2413000}},
    {"elbow", true, "NW", 1, 2, 2, {}, 2, false, false, {2032000, 1905000, 3048000, 2159000}},
    {"elbow", true, "NW", 1, 3, 3, {50000}, 2, false, false, {2540000, 1651000, 2540000, 2413000}},
    {"elbow", true, "NW", 2, 0, 5, {-7826, 50000, 107826}, 1, true, false, {2349500, 1079500, 2921000, 3556000}},
    {"elbow", true, "NW", 2, 1, 4, {-8571, 105625}, 1, true, false, {2222500, 952500, 2667000, 4064000}},
    {"elbow", true, "NW", 2, 2, 3, {-9474}, 1, true, false, {2603500, 1333500, 2413000, 3556000}},
    {"elbow", true, "NW", 2, 3, 4, {-8571, 58333}, 1, true, false, {2730500, 1460500, 2667000, 3048000}},
    {"elbow", true, "NW", 3, 0, 4, {-5625, 108571}, 0, true, true, {2032000, 1397000, 4064000, 2667000}},
    {"elbow", true, "NW", 3, 1, 5, {-5000, 50000, 105000}, 0, true, true, {1524000, 1651000, 4572000, 2413000}},
    {"elbow", true, "NW", 3, 2, 4, {-5625, 55882}, 0, true, true, {2032000, 1905000, 4064000, 2159000}},
    {"elbow", true, "NW", 3, 3, 3, {-6429}, 0, true, true, {2540000, 1651000, 3556000, 2413000}},
    {"curve", false, "NE", 0, 0, 3, {130000}, 1, true, true, {5842000, 2794000, 762000, 1270000}},
    {"curve", false, "NE", 0, 1, 2, {}, 1, true, true, {5715000, 3175000, 508000, 762000}},
    {"curve", false, "NE", 0, 2, 3, {50000}, 1, true, true, {6096000, 3048000, 254000, 1270000}},
    {"curve", false, "NE", 0, 3, 4, {25000, 112857}, 1, true, true, {6223000, 2667000, 508000, 1778000}},
    {"curve", false, "NE", 1, 0, 4, {-12857, 122500}, 2, true, false, {5080000, 3048000, 1778000, 1016000}},
    {"curve", false, "NE", 1, 1, 3, {-18000}, 2, true, false, {5080000, 3302000, 1270000, 762000}},
    {"curve", false, "NE", 1, 2, 4, {-12857, 75000}, 2, true, false, {5080000, 3556000, 1778000, 508000}},
    {"curve", false, "NE", 1, 3, 5, {-10000, 50000, 110000}, 2, true, false, {5080000, 3302000, 2286000, 762000}},
    {"curve", false, "NE", 2, 0, 5, {-18000, 50000, 118000}, 1, true, true, {5588000, 3048000, 1270000, 1270000}},
    {"curve", false, "NE", 2, 1, 4, {-22500, 83333}, 1, true, true, {5461000, 3429000, 1016000, 762000}},
    {"curve", false, "NE", 2, 2, 3, {-30000}, 1, true, true, {5842000, 3302000, 762000, 1270000}},
    {"curve", false, "NE", 2, 3, 4, {-22500, 112857}, 1, true, true, {5969000, 2921000, 1016000, 1778000}},
    {"curve", false, "NE", 3, 0, 4, {16667, 122500}, 0, false, true, {6096000, 3048000, 762000, 1016000}},
    {"curve", false, "NE", 3, 1, 3, {50000}, 0, false, true, {6096000, 3302000, 254000, 762000}},
    {"curve", false, "NE", 3, 2, 2, {}, 0, false, true, {6096000, 3556000, 762000, 508000}},
    {"curve", false, "NE", 3, 3, 3, {118000}, 0, false, true, {6096000, 3302000, 1270000, 762000}},
    {"curve", false, "SE", 0, 0, 3, {-30000}, 3, true, false, {5842000, 3556000, 762000, 1270000}},
    {"curve", false, "SE", 0, 1, 4, {-22500, 83333}, 3, true, false, {5461000, 3937000, 1016000, 762000}},
    {"curve", false, "SE", 0, 2, 5, {-18000, 50000, 118000}, 3, true, false, {5588000, 3810000, 1270000, 1270000}},
    {"curve", false, "SE", 0, 3, 4, {-22500, 112857}, 3, true, false, {5969000, 3429000, 1016000, 1778000}},
    {"curve", false, "SE", 1, 0, 4, {-12857, 75000}, 2, true, true, {5080000, 4064000, 1778000, 508000}},
    {"curve", false, "SE", 1, 1, 3, {-18000}, 2, true, true, {5080000, 4064000, 1270000, 762000}},
    {"curve", false, "SE", 1, 2, 4, {-12857, 122500}, 2, true, true, {5080000, 4064000, 1778000, 1016000}},
    {"curve", false, "SE", 1, 3, 5, {-10000, 50000, 110000}, 2, true, true, {5080000, 4064000, 2286000, 762000}},
    {"curve", false, "SE", 2, 0, 3, {50000}, 3, true, false, {6096000, 3810000, 254000, 1270000}},
    {"curve", false, "SE", 2, 1, 2, {}, 3, true, false, {5715000, 4191000, 508000, 762000}},
    {"curve", false, "SE", 2, 2, 3, {130000}, 3, true, false, {5842000, 4064000, 762000, 1270000}},
    {"curve", false, "SE", 2, 3, 4, {25000, 112857}, 3, true, false, {6223000, 3683000, 508000, 1778000}},
    {"curve", false, "SE", 3, 0, 2, {}, 0, false, false, {6096000, 4064000, 762000, 508000}},
    {"curve", false, "SE", 3, 1, 3, {50000}, 0, false, false, {6096000, 4064000, 254000, 762000}},
    {"curve", false, "SE", 3, 2, 4, {16667, 122500}, 0, false, false, {6096000, 4064000, 762000, 1016000}},
    {"curve", false, "SE", 3, 3, 3, {118000}, 0, false, false, {6096000, 4064000, 1270000, 762000}},
    {"curve", false, "SW", 0, 0, 3, {-30000}, 3, true, true, {4572000, 3556000, 762000, 1270000}},
    {"curve", false, "SW", 0, 1, 4, {-22500, 112857}, 3, true, true, {4191000, 3429000, 1016000, 1778000}},
    {"curve", false, "SW", 0, 2, 5, {-18000, 50000, 118000}, 3, true, true, {4318000, 3810000, 1270000, 1270000}},
    {"curve", false, "SW", 0, 3, 4, {-22500, 83333}, 3, true, true, {4699000, 3937000, 1016000, 762000}},
    {"curve", false, "SW", 1, 0, 2, {}, 2, false, true, {4318000, 4064000, 762000, 508000}},
    {"curve", false, "SW", 1, 1, 3, {118000}, 2, false, true, {3810000, 4064000, 1270000, 762000}},
    {"curve", false, "SW", 1, 2, 4, {16667, 122500}, 2, false, true, {4318000, 4064000, 762000, 1016000}},
    {"curve", false, "SW", 1, 3, 3, {50000}, 2, false, true, {4826000, 4064000, 254000, 762000}},
    {"curve", false, "SW", 2, 0, 3, {50000}, 1, false, false, {4826000, 3810000, 254000, 1270000}},
    {"curve", false, "SW", 2, 1, 4, {25000, 112857}, 1, false, false, {4445000, 3683000, 508000, 1778000}},
    {"curve", false, "SW", 2, 2, 3, {130000}, 1, false, false, {4572000, 4064000, 762000, 1270000}},
    {"curve", false, "SW", 2, 3, 2, {}, 1, false, false, {4953000, 4191000, 508000, 762000}},
    {"curve", false, "SW", 3, 0, 4, {-12857, 75000}, 0, true, false, {4318000, 4064000, 1778000, 508000}},
    {"curve", false, "SW", 3, 1, 5, {-10000, 50000, 110000}, 0, true, false, {3810000, 4064000, 2286000, 762000}},
    {"curve", false, "SW", 3, 2, 4, {-12857, 122500}, 0, true, false, {4318000, 4064000, 1778000, 1016000}},
    {"curve", false, "SW", 3, 3, 3, {-18000}, 0, true, false, {4826000, 4064000, 1270000, 762000}},
    {"curve", false, "NW", 0, 0, 3, {130000}, 3, false, true, {4572000, 2794000, 762000, 1270000}},
    {"curve", false, "NW", 0, 1, 4, {25000, 112857}, 3, false, true, {4445000, 2667000, 508000, 1778000}},
    {"curve", false, "NW", 0, 2, 3, {50000}, 3, false, true, {4826000, 3048000, 254000, 1270000}},
    {"curve", false, "NW", 0, 3, 2, {}, 3, false, true, {4953000, 3175000, 508000, 762000}},
    {"curve", false, "NW", 1, 0, 4, {16667, 122500}, 2, false, false, {4318000, 3048000, 762000, 1016000}},
    {"curve", false, "NW", 1, 1, 3, {118000}, 2, false, false, {3810000, 3302000, 1270000, 762000}},
    {"curve", false, "NW", 1, 2, 2, {}, 2, false, false, {4318000, 3556000, 762000, 508000}},
    {"curve", false, "NW", 1, 3, 3, {50000}, 2, false, false, {4826000, 3302000, 254000, 762000}},
    {"curve", false, "NW", 2, 0, 5, {-18000, 50000, 118000}, 1, true, false, {4318000, 3048000, 1270000, 1270000}},
    {"curve", false, "NW", 2, 1, 4, {-22500, 112857}, 1, true, false, {4191000, 2921000, 1016000, 1778000}},
    {"curve", false, "NW", 2, 2, 3, {-30000}, 1, true, false, {4572000, 3302000, 762000, 1270000}},
    {"curve", false, "NW", 2, 3, 4, {-22500, 83333}, 1, true, false, {4699000, 3429000, 1016000, 762000}},
    {"curve", false, "NW", 3, 0, 4, {-12857, 122500}, 0, true, true, {4318000, 3048000, 1778000, 1016000}},
    {"curve", false, "NW", 3, 1, 5, {-10000, 50000, 110000}, 0, true, true, {3810000, 3302000, 2286000, 762000}},
    {"curve", false, "NW", 3, 2, 4, {-12857, 75000}, 0, true, true, {4318000, 3556000, 1778000, 508000}},
    {"curve", false, "NW", 3, 3, 3, {-18000}, 0, true, true, {4826000, 3302000, 1270000, 762000}},
    {"curve", true, "NE", 0, 0, 3, {109474}, 1, true, true, {6159500, 825500, 2413000, 3556000}},
    {"curve", true, "NE", 0, 1, 2, {}, 1, true, true, {6032500, 1206500, 2159000, 3048000}},
    {"curve", true, "NE", 0, 2, 3, {50000}, 1, true, true, {6413500, 1079500, 1905000, 3556000}},
    {"curve", true, "NE", 0, 3, 4, {44118, 105625}, 1, true, true, {6540500, 698500, 2159000, 4064000}},
    {"curve", true, "NE", 1, 0, 4, {-5625, 108571}, 2, true, false, {5080000, 1397000, 4064000, 2667000}},
    {"curve", true, "NE", 1, 1, 3, {-6429}, 2, true, false, {5080000, 1651000, 3556000, 2413000}},
    {"curve", true, "NE", 1, 2, 4, {-5625, 55882}, 2, true, false, {5080000, 1905000, 4064000, 2159000}},
    {"curve", true, "NE", 1, 3, 5, {-5000, 50000, 105000}, 2, true, false, {5080000, 1651000, 4572000, 2413000}},
    {"curve", true, "NE", 2, 0, 5, {-7826, 50000, 107826}, 1, true, true, {5905500, 1079500, 2921000, 3556000}},
    {"curve", true, "NE", 2, 1, 4, {-8571, 58333}, 1, true, true, {5778500, 1460500, 2667000, 3048000}},
    {"curve", true, "NE", 2, 2, 3, {-9474}, 1, true, true, {6159500, 1333500, 2413000, 3556000}},
    {"curve", true, "NE", 2, 3, 4, {-8571, 105625}, 1, true, true, {6286500, 952500, 2667000, 4064000}},
    {"curve", true, "NE", 3, 0, 4, {41667, 108571}, 0, false, true, {6096000, 1397000, 3048000, 2667000}},
    {"curve", true, "NE", 3, 1, 3, {50000}, 0, false, true, {6096000, 1651000, 2540000, 2413000}},
    {"curve", true, "NE", 3, 2, 2, {}, 0, false, true, {6096000, 1905000, 3048000, 2159000}},
    {"curve", true, "NE", 3, 3, 3, {106429}, 0, false, true, {6096000, 1651000, 3556000, 2413000}},
    {"curve", true, "SE", 0, 0, 3, {-9474}, 3, true, false, {6159500, 3238500, 2413000, 3556000}},
    {"curve", true, "SE", 0, 1, 4, {-8571, 58333}, 3, true, false, {5778500, 3619500, 2667000, 3048000}},
    {"curve", true, "SE", 0, 2, 5, {-7826, 50000, 107826}, 3, true, false, {5905500, 3492500, 2921000, 3556000}},
    {"curve", true, "SE", 0, 3, 4, {-8571, 105625}, 3, true, false, {6286500, 3111500, 2667000, 4064000}},
    {"curve", true, "SE", 1, 0, 4, {-5625, 55882}, 2, true, true, {5080000, 4064000, 4064000, 2159000}},
    {"curve", true, "SE", 1, 1, 3, {-6429}, 2, true, true, {5080000, 4064000, 3556000, 2413000}},
    {"curve", true, "SE", 1, 2, 4, {-5625, 108571}, 2, true, true, {5080000, 4064000, 4064000, 2667000}},
    {"curve", true, "SE", 1, 3, 5, {-5000, 50000, 105000}, 2, true, true, {5080000, 4064000, 4572000, 2413000}},
    {"curve", true, "SE", 2, 0, 3, {50000}, 3, true, false, {6413500, 3492500, 1905000, 3556000}},
    {"curve", true, "SE", 2, 1, 2, {}, 3, true, false, {6032500, 3873500, 2159000, 3048000}},
    {"curve", true, "SE", 2, 2, 3, {109474}, 3, true, false, {6159500, 3746500, 2413000, 3556000}},
    {"curve", true, "SE", 2, 3, 4, {44118, 105625}, 3, true, false, {6540500, 3365500, 2159000, 4064000}},
    {"curve", true, "SE", 3, 0, 2, {}, 0, false, false, {6096000, 4064000, 3048000, 2159000}},
    {"curve", true, "SE", 3, 1, 3, {50000}, 0, false, false, {6096000, 4064000, 2540000, 2413000}},
    {"curve", true, "SE", 3, 2, 4, {41667, 108571}, 0, false, false, {6096000, 4064000, 3048000, 2667000}},
    {"curve", true, "SE", 3, 3, 3, {106429}, 0, false, false, {6096000, 4064000, 3556000, 2413000}},
    {"curve", true, "SW", 0, 0, 3, {-9474}, 3, true, true, {2603500, 3238500, 2413000, 3556000}},
    {"curve", true, "SW", 0, 1, 4, {-8571, 105625}, 3, true, true, {2222500, 3111500, 2667000, 4064000}},
    {"curve", true, "SW", 0, 2, 5, {-7826, 50000, 107826}, 3, true, true, {2349500, 3492500, 2921000, 3556000}},
    {"curve", true, "SW", 0, 3, 4, {-8571, 58333}, 3, true, true, {2730500, 3619500, 2667000, 3048000}},
    {"curve", true, "SW", 1, 0, 2, {}, 2, false, true, {2032000, 4064000, 3048000, 2159000}},
    {"curve", true, "SW", 1, 1, 3, {106429}, 2, false, true, {1524000, 4064000, 3556000, 2413000}},
    {"curve", true, "SW", 1, 2, 4, {41667, 108571}, 2, false, true, {2032000, 4064000, 3048000, 2667000}},
    {"curve", true, "SW", 1, 3, 3, {50000}, 2, false, true, {2540000, 4064000, 2540000, 2413000}},
    {"curve", true, "SW", 2, 0, 3, {50000}, 1, false, false, {2857500, 3492500, 1905000, 3556000}},
    {"curve", true, "SW", 2, 1, 4, {44118, 105625}, 1, false, false, {2476500, 3365500, 2159000, 4064000}},
    {"curve", true, "SW", 2, 2, 3, {109474}, 1, false, false, {2603500, 3746500, 2413000, 3556000}},
    {"curve", true, "SW", 2, 3, 2, {}, 1, false, false, {2984500, 3873500, 2159000, 3048000}},
    {"curve", true, "SW", 3, 0, 4, {-5625, 55882}, 0, true, false, {2032000, 4064000, 4064000, 2159000}},
    {"curve", true, "SW", 3, 1, 5, {-5000, 50000, 105000}, 0, true, false, {1524000, 4064000, 4572000, 2413000}},
    {"curve", true, "SW", 3, 2, 4, {-5625, 108571}, 0, true, false, {2032000, 4064000, 4064000, 2667000}},
    {"curve", true, "SW", 3, 3, 3, {-6429}, 0, true, false, {2540000, 4064000, 3556000, 2413000}},
    {"curve", true, "NW", 0, 0, 3, {109474}, 3, false, true, {2603500, 825500, 2413000, 3556000}},
    {"curve", true, "NW", 0, 1, 4, {44118, 105625}, 3, false, true, {2476500, 698500, 2159000, 4064000}},
    {"curve", true, "NW", 0, 2, 3, {50000}, 3, false, true, {2857500, 1079500, 1905000, 3556000}},
    {"curve", true, "NW", 0, 3, 2, {}, 3, false, true, {2984500, 1206500, 2159000, 3048000}},
    {"curve", true, "NW", 1, 0, 4, {41667, 108571}, 2, false, false, {2032000, 1397000, 3048000, 2667000}},
    {"curve", true, "NW", 1, 1, 3, {106429}, 2, false, false, {1524000, 1651000, 3556000, 2413000}},
    {"curve", true, "NW", 1, 2, 2, {}, 2, false, false, {2032000, 1905000, 3048000, 2159000}},
    {"curve", true, "NW", 1, 3, 3, {50000}, 2, false, false, {2540000, 1651000, 2540000, 2413000}},
    {"curve", true, "NW", 2, 0, 5, {-7826, 50000, 107826}, 1, true, false, {2349500, 1079500, 2921000, 3556000}},
    {"curve", true, "NW", 2, 1, 4, {-8571, 105625}, 1, true, false, {2222500, 952500, 2667000, 4064000}},
    {"curve", true, "NW", 2, 2, 3, {-9474}, 1, true, false, {2603500, 1333500, 2413000, 3556000}},
    {"curve", true, "NW", 2, 3, 4, {-8571, 58333}, 1, true, false, {2730500, 1460500, 2667000, 3048000}},
    {"curve", true, "NW", 3, 0, 4, {-5625, 108571}, 0, true, true, {2032000, 1397000, 4064000, 2667000}},
    {"curve", true, "NW", 3, 1, 5, {-5000, 50000, 105000}, 0, true, true, {1524000, 1651000, 4572000, 2413000}},
    {"curve", true, "NW", 3, 2, 4, {-5625, 55882}, 0, true, true, {2032000, 1905000, 4064000, 2159000}},
    {"curve", true, "NW", 3, 3, 3, {-6429}, 0, true, true, {2540000, 1651000, 3556000, 2413000}},
};

TEST(ConnectorRoute, BentAndCurvedConnectorsRouteAsExcelDoes) {
  constexpr std::int64_t kPt = 12700;
  const Box start_shape{400 * kPt, 300 * kPt, 80 * kPt, 40 * kPt};
  const auto site = [](const Box& b, std::uint32_t idx) {
    const std::int64_t xs[] = {b.x + b.w / 2, b.x, b.x + b.w / 2, b.x + b.w};
    const std::int64_t ys[] = {b.y, b.y + b.h / 2, b.y + b.h, b.y + b.h / 2};
    return std::make_pair(xs[idx], ys[idx]);
  };
  for (const MatrixRow& row : kMatrix) {
    const std::string q = row.quadrant;
    const std::int64_t gap_x = (row.far ? 200 : 20) * kPt;
    const std::int64_t gap_y = (row.far ? 150 : 20) * kPt;
    const Box end_shape{q[1] == 'E' ? start_shape.x + start_shape.w + gap_x : start_shape.x - start_shape.w - gap_x,
                        q[0] == 'S' ? start_shape.y + start_shape.h + gap_y : start_shape.y - start_shape.h - gap_y,
                        start_shape.w, start_shape.h};
    const auto [sx, sy] = site(start_shape, row.start_site);
    const auto [ex, ey] = site(end_shape, row.end_site);
    const Xfrm authored{std::min(sx, ex), std::min(sy, ey), ex > sx ? ex - sx : sx - ex, ey > sy ? ey - sy : sy - ey,
                        ex < sx,          ey < sy};
    const std::string family = std::string(row.family) == "elbow" ? "bentConnector" : "curvedConnector";
    Workbook wb =
        Book(Absolute(start_shape, Sp(2, "rect", start_shape)) + Absolute(end_shape, Sp(3, "rect", end_shape)) +
             Absolute(BoxOf(authored), Cxn(4, (family + "3").c_str(), authored,
                                           Glue("stCxn", 2, row.start_site) + Glue("endCxn", 3, row.end_site),
                                           "<a:gd name=\"adj1\" fmla=\"val 50000\"/>")));
    // Excel's route was read after a 1 pt move and back.
    Move(wb, 3, Box{end_shape.x + kPt, end_shape.y, end_shape.w, end_shape.h});
    Move(wb, 3, end_shape);
    const std::string what = std::string(row.family) + (row.far ? " far " : " near ") + q + " " +
                             std::to_string(row.start_site) + std::to_string(row.end_site);
    std::string guides;
    for (std::size_t i = 0; i < row.adjust.size(); ++i) {
      guides += "<a:gd name=\"adj" + std::to_string(i + 1U) + "\" fmla=\"val " + std::to_string(row.adjust[i]) + "\"/>";
    }
    EXPECT_EQ(GeometryXml(wb, 4), "<a:prstGeom prst=\"" + family + std::to_string(row.segments) + "\">" +
                                      (guides.empty() ? "<a:avLst/>" : "<a:avLst>" + guides + "</a:avLst>") +
                                      "</a:prstGeom>")
        << what;
    EXPECT_EQ(ConnectorXfrm(wb, 4), (Xfrm{row.box.x, row.box.y, row.box.w, row.box.h, row.flip_h, row.flip_v})) << what;
    EXPECT_EQ(ConnectorRotation(wb, 4), row.quarter_turns * 5400000) << what;
  }
}

struct TransformRow {
  const char* prst;
  const char* transform;
  std::uint32_t idx;
  Xfrm before, after;  // y from kTop
};

// Shape 3 (120 x 80 pt at 330 pt, kTop) is rotated or flipped; the connector runs from shape 2's site 3 at
// (762000, kTop + 635000) to shape 3's site `idx`.
const TransformRow kTransforms[] = {
    {"rightArrow", "flipH", 0, {762000, 0, 4445000, 635000, false, true}, {762000, 0, 3937000, 635000, false, true}},
    {"rightArrow",
     "flipH",
     1,
     {762000, 508000, 3429000, 127000, false, true},
     {762000, 508000, 4953000, 127000, false, true}},
    {"rightArrow",
     "flipH",
     2,
     {762000, 635000, 4445000, 381000, false, false},
     {762000, 635000, 3937000, 381000, false, false}},
    {"rightArrow",
     "flipH",
     3,
     {762000, 508000, 4953000, 127000, false, true},
     {762000, 508000, 3429000, 127000, false, true}},
    {"rightArrow",
     "flipV",
     0,
     {762000, 0, 4445000, 635000, false, true},
     {762000, 635000, 4445000, 381000, false, false}},
    {"rightArrow",
     "flipV",
     1,
     {762000, 508000, 3429000, 127000, false, true},
     {762000, 508000, 3429000, 127000, false, true}},
    {"rightArrow",
     "flipV",
     2,
     {762000, 635000, 4445000, 381000, false, false},
     {762000, 0, 4445000, 635000, false, true}},
    {"rightArrow",
     "flipV",
     3,
     {762000, 508000, 4953000, 127000, false, true},
     {762000, 508000, 4953000, 127000, false, true}},
    {"rightArrow",
     "rot45",
     0,
     {762000, 0, 4445000, 635000, false, true},
     {762000, 328395, 4729815, 306605, false, true}},
    {"rightArrow",
     "rot45",
     1,
     {762000, 508000, 3429000, 127000, false, true},
     {762000, -30815, 3652185, 665815, false, true}},
    {"rightArrow",
     "rot45",
     2,
     {762000, 635000, 4445000, 381000, false, false},
     {762000, 635000, 4011395, 411815, false, false}},
    {"rightArrow",
     "rot45",
     3,
     {762000, 508000, 4953000, 127000, false, true},
     {762000, 635000, 4729815, 411815, false, false}},
    {"rightArrow",
     "rot90",
     0,
     {762000, 0, 4445000, 635000, false, true},
     {762000, 635000, 4699000, 127000, false, false}},
    {"rightArrow",
     "rot90",
     1,
     {762000, 508000, 3429000, 127000, false, true},
     {762000, -254000, 4191000, 889000, false, true}},
    {"rightArrow",
     "rot90",
     2,
     {762000, 635000, 4445000, 381000, false, false},
     {762000, 635000, 3683000, 127000, false, false}},
    {"rightArrow",
     "rot90",
     3,
     {762000, 508000, 4953000, 127000, false, true},
     {762000, 635000, 4191000, 635000, false, false}},
    {"triangle", "flipH", 0, {762000, 0, 4191000, 635000, false, true}, {762000, 0, 4191000, 635000, false, true}},
    {"triangle",
     "flipH",
     1,
     {762000, 508000, 3810000, 127000, false, true},
     {762000, 508000, 4572000, 127000, false, true}},
    {"triangle",
     "flipH",
     2,
     {762000, 635000, 3429000, 381000, false, false},
     {762000, 635000, 4953000, 381000, false, false}},
    {"triangle",
     "flipH",
     3,
     {762000, 635000, 4191000, 381000, false, false},
     {762000, 635000, 4191000, 381000, false, false}},
    {"triangle",
     "flipH",
     4,
     {762000, 635000, 4953000, 381000, false, false},
     {762000, 635000, 3429000, 381000, false, false}},
    {"triangle",
     "flipH",
     5,
     {762000, 508000, 4572000, 127000, false, true},
     {762000, 508000, 3810000, 127000, false, true}},
    {"triangle",
     "flipV",
     0,
     {762000, 0, 4191000, 635000, false, true},
     {762000, 635000, 4191000, 381000, false, false}},
    {"triangle",
     "flipV",
     1,
     {762000, 508000, 3810000, 127000, false, true},
     {762000, 508000, 3810000, 127000, false, true}},
    {"triangle",
     "flipV",
     2,
     {762000, 635000, 3429000, 381000, false, false},
     {762000, 0, 3429000, 635000, false, true}},
    {"triangle",
     "flipV",
     3,
     {762000, 635000, 4191000, 381000, false, false},
     {762000, 0, 4191000, 635000, false, true}},
    {"triangle",
     "flipV",
     4,
     {762000, 635000, 4953000, 381000, false, false},
     {762000, 0, 4953000, 635000, false, true}},
    {"triangle",
     "flipV",
     5,
     {762000, 508000, 4572000, 127000, false, true},
     {762000, 508000, 4572000, 127000, false, true}},
    {"triangle", "rot45", 0, {762000, 0, 4191000, 635000, false, true}, {762000, 148790, 4550210, 486210, false, true}},
    {"triangle",
     "rot45",
     1,
     {762000, 508000, 3810000, 127000, false, true},
     {762000, 238592, 3921592, 396408, false, true}},
    {"triangle",
     "rot45",
     2,
     {762000, 635000, 3429000, 381000, false, false},
     {762000, 328395, 3292974, 306605, false, true}},
    {"triangle",
     "rot45",
     3,
     {762000, 635000, 4191000, 381000, false, false},
     {762000, 635000, 3831790, 232210, false, false}},
    {"triangle",
     "rot45",
     4,
     {762000, 635000, 4953000, 381000, false, false},
     {762000, 635000, 4370605, 771026, false, false}},
    {"triangle",
     "rot45",
     5,
     {762000, 508000, 4572000, 127000, false, true},
     {762000, 635000, 4460408, 142408, false, false}},
    {"triangle", "rot90", 0, {762000, 0, 4191000, 635000, false, true}, {762000, 508000, 4699000, 127000, false, true}},
    {"triangle",
     "rot90",
     1,
     {762000, 508000, 3810000, 127000, false, true},
     {762000, 127000, 4191000, 508000, false, true}},
    {"triangle",
     "rot90",
     2,
     {762000, 635000, 3429000, 381000, false, false},
     {762000, -254000, 3683000, 889000, false, true}},
    {"triangle",
     "rot90",
     3,
     {762000, 635000, 4191000, 381000, false, false},
     {762000, 508000, 3683000, 127000, false, true}},
    {"triangle",
     "rot90",
     4,
     {762000, 635000, 4953000, 381000, false, false},
     {762000, 635000, 3683000, 635000, false, false}},
    {"triangle",
     "rot90",
     5,
     {762000, 508000, 4572000, 127000, false, true},
     {762000, 635000, 4191000, 254000, false, false}},
};

Xfrm Lifted(Xfrm x) {
  x.y += kTop;
  return x;
}

TEST(ConnectorRoute, RotatingOrFlippingTheShapeMovesTheGluedEnd) {
  const Box anchor_shape{254000, kTop + 381000, 508000, 508000};
  const Box target{4191000, kTop, 1524000, 1016000};
  for (const TransformRow& row : kTransforms) {
    const Xfrm before = Lifted(row.before);
    Workbook wb = Book(Absolute(anchor_shape, Sp(2, "rect", anchor_shape)) + Absolute(target, Sp(3, row.prst, target)) +
                       Absolute(BoxOf(before), Cxn(4, "straightConnector1", before,
                                                   Glue("stCxn", 2, 3) + Glue("endCxn", 3, row.idx))));
    ShapeSpec spec;
    const std::string transform = row.transform;
    if (transform == "flipH") {
      spec.fields = shape_field::kFlipH;
      spec.flip_h = true;
    } else if (transform == "flipV") {
      spec.fields = shape_field::kFlipV;
      spec.flip_v = true;
    } else {
      spec.fields = shape_field::kRotation;
      spec.rotation = (transform == "rot45" ? 45 : 90) * 60000;
    }
    ASSERT_TRUE(static_cast<bool>(update_shape(wb, 0, 3, spec)));
    EXPECT_EQ(ConnectorXfrm(wb, 4), Lifted(row.after)) << row.prst << " " << row.transform << " idx " << row.idx;
  }
}

// Glued at its end only, the connector moves with the shape by the same distance, its free start included.
TEST(ConnectorRoute, AConnectorGluedAtOneEndMovesWithTheShape) {
  struct Row {
    Xfrm before;
    std::int64_t dx, dy;
    Xfrm after;
  };
  const Row rows[] = {
      {{3810000, 4064000, 1270000, 762000, false, true},
       635000,
       381000,
       {4445000, 4445000, 1270000, 762000, false, true}},
      {{3810000, 4064000, 1270000, 762000, false, true},
       -635000,
       -381000,
       {3175000, 3683000, 1270000, 762000, false, true}},
      {{3810000, 3302000, 1270000, 762000, false, false},
       635000,
       381000,
       {4445000, 3683000, 1270000, 762000, false, false}},
      {{3810000, 3302000, 1270000, 762000, false, false},
       -635000,
       -381000,
       {3175000, 2921000, 1270000, 762000, false, false}},
  };
  const Box shape{5080000, 3810000, 1016000, 508000};
  for (const char* prst : {"straightConnector1", "bentConnector3"}) {
    for (const Row& row : rows) {
      const std::string guide = std::string(prst) == "bentConnector3" ? "<a:gd name=\"adj1\" fmla=\"val 50000\"/>" : "";
      Workbook wb = Book(Absolute(shape, Sp(3, "rect", shape)) +
                         Absolute(BoxOf(row.before), Cxn(4, prst, row.before, Glue("endCxn", 3, 1), guide)));
      Move(wb, 3, Box{shape.x + row.dx, shape.y + row.dy, shape.w, shape.h});
      EXPECT_EQ(ConnectorXfrm(wb, 4), row.after) << prst << " " << row.dx;
      const std::string guides = guide.empty() ? "<a:avLst/>" : "<a:avLst>" + guide + "</a:avLst>";
      EXPECT_EQ(GeometryXml(wb, 4), std::string("<a:prstGeom prst=\"") + prst + "\">" + guides + "</a:prstGeom>");
    }
  }
}

// A stated formula guide places the site by its value: `*/ 60000 1 2` lands where `val 30000` does.
TEST(ConnectorRoute, AFormulaGuidePlacesTheSiteByItsValue) {
  const Box start_shape{508000, 508000, 1016000, 508000};
  const Box p{3048000, 1016000, 1524000, 1016000};
  const Xfrm before{1524000, 762000, 2921000, 762000, false, false};
  for (const char* fmla : {"*/ 60000 1 2", "val 30000"}) {
    Workbook wb =
        Book(Absolute(start_shape, Sp(2, "rect", start_shape)) +
             Absolute(p, Sp(3, "parallelogram", p, "", std::string("<a:gd name=\"adj\" fmla=\"") + fmla + "\"/>")) +
             Absolute(BoxOf(before), Cxn(4, "straightConnector1", before, Glue("stCxn", 2, 3) + Glue("endCxn", 3, 2))));
    Move(wb, 3, Box{p.x + 635000, p.y + 381000, p.w, p.h});
    EXPECT_EQ(ConnectorXfrm(wb, 4), (Xfrm{1524000, 762000, 3530600, 1143000, false, false})) << fmla;
  }
}

// An edit that leaves every glued site where it was moves no connector, even one not lying on its sites.
TEST(ConnectorRoute, AnEditLeavingTheSitesInPlaceMovesNothing) {
  const Xfrm off_site{100, 200, 300, 400, true, false};
  Workbook wb = Book(
      Absolute(kStartShape, Sp(2, "rect", kStartShape)) + Absolute(kEndShape, Sp(3, "rect", kEndShape)) +
      Absolute(BoxOf(off_site), Cxn(4, "straightConnector1", off_site, Glue("stCxn", 2, 3) + Glue("endCxn", 3, 1))));
  ShapeSpec rename;
  rename.fields = shape_field::kName;
  rename.name = "Renamed";
  ASSERT_TRUE(static_cast<bool>(update_shape(wb, 0, 3, rename)));
  // A rect's left middle stays put when the rect is flipped.
  ShapeSpec flip;
  flip.fields = shape_field::kFlipV;
  flip.flip_v = true;
  ASSERT_TRUE(static_cast<bool>(update_shape(wb, 0, 3, flip)));
  EXPECT_EQ(ConnectorXfrm(wb, 4), off_site);
}

// A glued end names a site its shape has; one past the list moves nothing, so the other glued end alone moves the
// connector.
TEST(ConnectorRoute, AnEndNamingNoSiteCountsAsFree) {
  Workbook wb = Book(Absolute(kStartShape, Sp(2, "rect", kStartShape)) + Absolute(kEndShape, Sp(3, "rect", kEndShape)) +
                     Absolute(BoxOf(kBothRoute),
                              Cxn(4, "straightConnector1", kBothRoute, Glue("stCxn", 2, 9) + Glue("endCxn", 3, 1))));
  Move(wb, 3, Box{kEndShape.x + 635000, kEndShape.y + 381000, kEndShape.w, kEndShape.h});
  EXPECT_EQ(ConnectorXfrm(wb, 4), (Xfrm{3048000 + 635000, kTop + 381000, 1016000, 762000, false, false}));
}

std::string Group(std::uint32_t id, const Box& b, const std::string& children) {
  return "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"" + Num(id) + "\" name=\"Group " + Num(id) +
         "\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x=\"" + Num(b.x) + "\" y=\"" + Num(b.y) +
         "\"/><a:ext cx=\"" + Num(b.w) + "\" cy=\"" + Num(b.h) + "\"/><a:chOff x=\"" + Num(b.x) + "\" y=\"" + Num(b.y) +
         "\"/><a:chExt cx=\"" + Num(b.w) + "\" cy=\"" + Num(b.h) + "\"/></a:xfrm></xdr:grpSpPr>" + children +
         "</xdr:grpSp>";
}

// Moving a group moves the sites of its children: a connector outside it follows, one inside it with both of its
// shapes keeps its place in the group.
TEST(ConnectorRoute, AGroupMoveCarriesTheSitesOfItsChildren) {
  const Box group{kEndShape.x, kEndShape.y, kEndShape.w, kEndShape.h};
  Workbook wb = Book(Absolute(kStartShape, Sp(2, "rect", kStartShape)) +
                     Absolute(group, Group(5, group, Sp(3, "rect", kEndShape))) +
                     Absolute(BoxOf(kBothRoute),
                              Cxn(4, "straightConnector1", kBothRoute, Glue("stCxn", 2, 3) + Glue("endCxn", 3, 1))));
  Move(wb, 5, Box{group.x + 635000, group.y + 381000, group.w, group.h});
  EXPECT_EQ(ConnectorXfrm(wb, 4), kMoves[0].expected);

  const Box both{kStartShape.x, kStartShape.y, kEndShape.x + kEndShape.w - kStartShape.x,
                 kEndShape.y + kEndShape.h - kStartShape.y};
  Workbook inside = Book(
      Absolute(both, Group(5, both,
                           Sp(2, "rect", kStartShape) + Sp(3, "rect", kEndShape) +
                               Cxn(4, "straightConnector1", kBothRoute, Glue("stCxn", 2, 3) + Glue("endCxn", 3, 1)))));
  Move(inside, 5, Box{both.x + 635000, both.y + 381000, both.w, both.h});
  EXPECT_EQ(ConnectorXfrm(inside, 4), kBothRoute);
}

}  // namespace
}  // namespace formulon
