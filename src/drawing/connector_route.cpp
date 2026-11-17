#include "drawing/connector_route.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "drawing/drawing_edit_internal.h"
#include "drawing/drawing_geometry.h"
#include "drawing/drawing_xml.h"
#include "io/ooxml/preset_geometry.h"
#include "io/xml_utils.h"
#include "utils/elementary_math.h"
#include "workbook.h"

namespace formulon {
namespace drawing_edit {
namespace {

struct Point {
  double x = 0.0;
  double y = 0.0;
};

constexpr double kRadiansPerUnit = 3.14159265358979323846 / (180.0 * 60000.0);
constexpr double kQuarterTurn = 5400000.0;
// How far a bent or curved connector runs past a shape it has to go around: 18 pt.
constexpr double kClearance = 228600.0;
// The side a routed connector's box has where its two ends line up: 1 pt.
constexpr double kMinimumSide = 12700.0;

std::int64_t emu(const pugi::xml_node& node, const char* name) {
  return node.attribute(name).as_llong();
}

double emu_d(const pugi::xml_node& node, const char* name) {
  return static_cast<double>(emu(node, name));
}

// Maps a point of a frame to the sheet: per axis, `scale * p + shift`.
struct Frame {
  double sx = 1.0, sy = 1.0, tx = 0.0, ty = 0.0;
  Point to_sheet(const Point& p) const { return Point{sx * p.x + tx, sy * p.y + ty}; }
  Point from_sheet(const Point& p) const { return Point{(p.x - tx) / sx, (p.y - ty) / sy}; }
};

// The frame the `a:xfrm` of `content` is stated in: the child space of every group holding it, mapped through
// the group's `off` / `ext` against its `chOff` / `chExt`.
Frame frame_of(const pugi::xml_node& content) {
  Frame f;
  for (pugi::xml_node node = content.parent(); node; node = node.parent()) {
    if (local_name(node) != "grpSp") {
      continue;
    }
    const pugi::xml_node xfrm = child_local(child_local(node, "grpSpPr"), "xfrm");
    const pugi::xml_node off = child_local(xfrm, "off");
    const pugi::xml_node ext = child_local(xfrm, "ext");
    const pugi::xml_node ch_off = child_local(xfrm, "chOff");
    const pugi::xml_node ch_ext = child_local(xfrm, "chExt");
    const double kx = emu(ch_ext, "cx") != 0 ? emu_d(ext, "cx") / emu_d(ch_ext, "cx") : 1.0;
    const double ky = emu(ch_ext, "cy") != 0 ? emu_d(ext, "cy") / emu_d(ch_ext, "cy") : 1.0;
    f = Frame{kx * f.sx, ky * f.sy, kx * (f.tx - emu_d(ch_off, "x")) + emu_d(off, "x"),
              ky * (f.ty - emu_d(ch_off, "y")) + emu_d(off, "y")};
  }
  return f;
}

// A site placed on the sheet: its point, the direction it leaves the shape in (60000ths of a degree, clockwise),
// and two opposite corners of the shape's unrotated box; `on_site` is false for an index past the preset's list,
// which leaves only the box.
struct PlacedSite {
  bool on_site = true;
  Point at;
  double angle = 0.0;
  Point box[2];
};

// Site `idx` of the shape or picture `content`: the preset's site on the current size and stated adjust guides,
// flipped and rotated with the object, then placed through its frame.
std::optional<PlacedSite> place_site(const pugi::xml_node& content, std::uint32_t idx) {
  const pugi::xml_node props = child_local(content, "spPr");
  const pugi::xml_node geom = child_local(props, "prstGeom");
  const std::optional<std::size_t> preset = preset_geometry::index_of(geom.attribute("prst").value());
  const pugi::xml_node xfrm = child_local(props, "xfrm");
  if (!preset || !xfrm) {
    return std::nullopt;
  }
  const double w = emu_d(child_local(xfrm, "ext"), "cx");
  const double h = emu_d(child_local(xfrm, "ext"), "cy");
  // A stated guide is placed by its value, a formula one included.
  std::vector<preset_geometry::AdjustGuide> stated;
  for (const pugi::xml_node& gd : children_local(child_local(geom, "avLst"), "gd")) {
    if (const std::optional<double> v = preset_geometry::guide_value(gd.attribute("fmla").value(), w, h)) {
      stated.push_back(preset_geometry::AdjustGuide{gd.attribute("name").value(), *v});
    }
  }
  const std::vector<preset_geometry::ConnectionSite> sites = preset_geometry::connection_sites(*preset, w, h, stated);
  const XfrmTransform t = read_xfrm_transform(xfrm);
  const double rot = static_cast<double>(t.rot) * kRadiansPerUnit;
  const double c = t.rot != 0 ? elementary::cos(rot) : 1.0;
  const double s = t.rot != 0 ? elementary::sin(rot) : 0.0;
  const Point off{emu_d(child_local(xfrm, "off"), "x"), emu_d(child_local(xfrm, "off"), "y")};
  const Frame frame = frame_of(content);
  // A point of the unflipped, unrotated shape, flipped and turned about the centre, then put on the sheet.
  const auto place = [&](double x, double y) {
    x = t.flip_h ? w - x : x;
    y = t.flip_v ? h - y : y;
    const double dx = x - w / 2.0;
    const double dy = y - h / 2.0;
    return frame.to_sheet(Point{off.x + w / 2.0 + dx * c - dy * s, off.y + h / 2.0 + dx * s + dy * c});
  };
  PlacedSite placed;
  // Routing goes around the box the shape has before it is rotated.
  placed.box[0] = frame.to_sheet(off);
  placed.box[1] = frame.to_sheet(Point{off.x + w, off.y + h});
  placed.on_site = idx < sites.size();
  if (!placed.on_site) {
    return placed;
  }
  placed.at = place(sites[idx].x, sites[idx].y);
  double angle = sites[idx].angle;
  angle = t.flip_h ? 2.0 * kQuarterTurn - angle : angle;
  angle = t.flip_v ? -angle : angle;
  placed.angle = angle + static_cast<double>(t.rot);
  return placed;
}

// The shapes and pictures connectors can be glued to, by id; the first of a repeated id.
std::vector<std::pair<std::uint32_t, pugi::xml_node>> glue_targets(const pugi::xml_node& root) {
  std::vector<std::pair<std::uint32_t, pugi::xml_node>> out;
  for_each_element(root, [&out](const pugi::xml_node& node) {
    const std::string_view name = local_name(node);
    if (name != "sp" && name != "pic") {
      return;
    }
    const std::optional<std::uint32_t> id = io::parse_xml_u32_attr_strict(content_cnvpr(node).attribute("id"));
    if (id && std::none_of(out.begin(), out.end(), [&id](const auto& t) { return t.first == *id; })) {
      out.emplace_back(*id, node);
    }
  });
  return out;
}

const GluedEnd* find_end(const std::vector<GluedEnd>& ends, const pugi::xml_node& connector, bool start) {
  for (const GluedEnd& end : ends) {
    if (end.connector == connector && end.start == start) {
      return &end;
    }
  }
  return nullptr;
}

// Whether a glued end's point moved; an end at no site never moves.
bool moved(const GluedEnd* before, const GluedEnd& after) {
  if (before != nullptr && !before->on_site && !after.on_site) {
    return false;
  }
  return before == nullptr || before->on_site != after.on_site || before->x != after.x || before->y != after.y;
}

bool changed(const GluedEnd* before, const GluedEnd& after) {
  return moved(before, after) || before->direction != after.direction || before->left != after.left ||
         before->top != after.top || before->right != after.right || before->bottom != after.bottom;
}

// The start or end point of the connector `cxn` in its own frame: a corner of its box, flipped and turned with it.
Point end_point(const pugi::xml_node& cxn, bool start) {
  const pugi::xml_node xfrm = child_local(child_local(cxn, "spPr"), "xfrm");
  const XfrmTransform t = read_xfrm_transform(xfrm);
  const double w = emu_d(child_local(xfrm, "ext"), "cx");
  const double h = emu_d(child_local(xfrm, "ext"), "cy");
  const double dx = ((start != t.flip_h) ? 0.0 : w) - w / 2.0;
  const double dy = ((start != t.flip_v) ? 0.0 : h) - h / 2.0;
  const double rot = static_cast<double>(t.rot) * kRadiansPerUnit;
  const double c = t.rot != 0 ? elementary::cos(rot) : 1.0;
  const double s = t.rot != 0 ? elementary::sin(rot) : 0.0;
  const double cx = emu_d(child_local(xfrm, "off"), "x") + w / 2.0;
  const double cy = emu_d(child_local(xfrm, "off"), "y") + h / 2.0;
  return Point{cx + dx * c - dy * s, cy + dx * s + dy * c};
}

// `v` turned `quarters` quarter turns counterclockwise: the inverse of a clockwise `a:xfrm` rotation.
Point unturn(Point v, int quarters) {
  for (int i = 0; i < quarters; ++i) {
    v = Point{v.y, -v.x};
  }
  return v;
}

constexpr Point kDirections[] = {{1.0, 0.0}, {0.0, 1.0}, {-1.0, 0.0}, {0.0, -1.0}};

std::int32_t adjust(double v, double span) {
  return static_cast<std::int32_t>(std::llround(v / span * 100000.0));
}

void write_geometry(pugi::xml_node geom, std::string_view family, int segments, const std::vector<std::int32_t>& adj) {
  const std::string prst = std::string(family) + std::to_string(segments);
  geom.attribute("prst").set_value(prst.c_str());
  const std::string ap(io::qname_prefix_colon(geom.name()));
  pugi::xml_node guides = child_local(geom, "avLst");
  if (!guides) {
    guides = geom.prepend_child((ap + "avLst").c_str());
  }
  while (guides.first_child()) {
    guides.remove_child(guides.first_child());
  }
  for (std::size_t i = 0; i < adj.size(); ++i) {
    pugi::xml_node gd = guides.append_child((ap + "gd").c_str());
    gd.append_attribute("name").set_value(("adj" + std::to_string(i + 1U)).c_str());
    gd.append_attribute("fmla").set_value(("val " + std::to_string(adj[i])).c_str());
  }
}

// A bent or curved connector glued at both ends, routed as Excel routes it. The connector's frame is turned so
// that it leaves the start along its local x axis and flipped so that the end lies at (w, h); in that frame the
// path turns halfway from the start to the end's shape, or runs `kClearance` past the shapes' unrotated boxes
// where it has to come back, and the number of segments follows from the two directions.
void route_elbow(pugi::xml_node xfrm, pugi::xml_node geom, std::string_view family, const GluedEnd& start,
                 const GluedEnd& end) {
  int r = start.direction;
  const Point delta = unturn(Point{static_cast<double>(end.x - start.x), static_cast<double>(end.y - start.y)}, r);
  bool flip_h = delta.x < 0.0;
  bool flip_v = delta.y < 0.0;
  // Excel writes these two frames as their half-turned equivalents.
  if ((r == 3 && !flip_h && !flip_v) || (r == 1 && !flip_h && flip_v)) {
    r = (r + 2) % 4;
    flip_h = !flip_h;
    flip_v = !flip_v;
  }
  // A sheet vector or point in the connector's local frame, the start at the origin.
  const auto local = [&](Point v) {
    v = unturn(v, r);
    return Point{flip_h ? -v.x : v.x, flip_v ? -v.y : v.y};
  };
  // An end at no site has no box to go around, only its point.
  const auto local_box = [&](const GluedEnd& e) {
    double box[4] = {0.0, 0.0, 0.0, 0.0};  // left, right, top, bottom
    bool first = true;
    const double xs[2] = {e.on_site ? e.left : static_cast<double>(e.x),
                          e.on_site ? e.right : static_cast<double>(e.x)};
    const double ys[2] = {e.on_site ? e.top : static_cast<double>(e.y),
                          e.on_site ? e.bottom : static_cast<double>(e.y)};
    for (const double x : xs) {
      for (const double y : ys) {
        const Point p = local(Point{x - static_cast<double>(start.x), y - static_cast<double>(start.y)});
        box[0] = first ? p.x : std::min(box[0], p.x);
        box[1] = first ? p.x : std::max(box[1], p.x);
        box[2] = first ? p.y : std::min(box[2], p.y);
        box[3] = first ? p.y : std::max(box[3], p.y);
        first = false;
      }
    }
    return std::array<double, 4>{box[0], box[1], box[2], box[3]};
  };
  const Point span = local(Point{static_cast<double>(end.x - start.x), static_cast<double>(end.y - start.y)});
  const std::int64_t cx = static_cast<std::int64_t>(std::llround(span.x));
  const std::int64_t cy = static_cast<std::int64_t>(std::llround(span.y));
  // A side of no length counts as 1 pt.
  const double w = cx != 0 ? span.x : kMinimumSide;
  const double h = cy != 0 ? span.y : kMinimumSide;
  const std::array<double, 4> b1 = local_box(start);
  const std::array<double, 4> b2 = local_box(end);
  const bool out_forward = local(kDirections[start.direction]).x > 0.0;
  const Point to_end = local(kDirections[end.direction]);
  // Where the shapes overlap across the path, it runs past the far side of both instead of between them, and
  // clears only the shape it leaves or enters.
  const bool apart = b1[3] <= b2[2];
  const double beyond_x = (apart ? std::max(b1[1], b2[1]) : b2[1]) + kClearance;
  const double beyond_y = std::max(b1[3], b2[3]) + kClearance;
  const double gap_x = b2[0] / 2.0;
  const double gap_y = apart ? (b1[3] + b2[2]) / 2.0 : beyond_y;
  const double back_x = (apart ? std::min(b1[0], b2[0]) : b1[0]) - kClearance;
  int segments = 3;
  std::vector<std::int32_t> adj;
  if (out_forward) {
    if (to_end.x < 0.0) {
      adj = {adjust(gap_x, w)};
    } else if (to_end.x > 0.0) {
      // The end's shape standing across the way out is gone around.
      if (b2[0] > 0.0 && b2[2] < 0.0 && b2[3] > 0.0) {
        segments = 5;
        adj = {adjust(gap_x, w), adjust(beyond_y, h), adjust(beyond_x, w)};
      } else {
        adj = {adjust(beyond_x, w)};
      }
    } else if (to_end.y < 0.0) {
      segments = 2;
    } else {
      segments = 4;
      adj = {adjust(gap_x, w), adjust(beyond_y, h)};
    }
  } else if (to_end.x < 0.0) {
    // The start's shape standing across the way back is gone around on the side away from the end.
    if (b1[2] < h && h < b1[3]) {
      segments = 5;
      adj = {adjust(back_x, w), adjust(b1[2] - kClearance, h), adjust((b1[1] + b2[0]) / 2.0, w)};
    } else {
      adj = {adjust(back_x, w)};
    }
  } else if (to_end.x > 0.0) {
    segments = 5;
    adj = {adjust(back_x, w), adjust(gap_y, h), adjust(beyond_x, w)};
  } else {
    segments = 4;
    adj = {adjust(std::min(0.0, b2[0]) - kClearance, w), adjust(to_end.y < 0.0 ? gap_y : beyond_y, h)};
  }
  write_geometry(geom, family, segments, adj);
  // The box is the start-to-end box, turned with the frame about its centre.
  io::set_attr(child_local(xfrm, "off"), "x", (start.x + end.x - cx) / 2);
  io::set_attr(child_local(xfrm, "off"), "y", (start.y + end.y - cy) / 2);
  io::set_attr(child_local(xfrm, "ext"), "cx", cx != 0 ? cx : static_cast<std::int64_t>(kMinimumSide));
  io::set_attr(child_local(xfrm, "ext"), "cy", cy != 0 ? cy : static_cast<std::int64_t>(kMinimumSide));
  write_xfrm_transform(xfrm, XfrmTransform{static_cast<std::int32_t>(r) * 5400000, flip_h, flip_v});
}

// Both ends glued: a straight connector runs from its start site to its end site, its flips giving the direction;
// a bent or curved one is routed around the two shapes.
void route_connector(pugi::xml_node connector, pugi::xml_node xfrm, const GluedEnd& start, const GluedEnd& end) {
  const pugi::xml_node geom = child_local(child_local(connector, "spPr"), "prstGeom");
  const std::string_view prst = geom.attribute("prst").value();
  for (const std::string_view family : {std::string_view("bentConnector"), std::string_view("curvedConnector")}) {
    if (prst.substr(0, family.size()) == family) {
      route_elbow(xfrm, geom, family, start, end);
      return;
    }
  }
  io::set_attr(child_local(xfrm, "off"), "x", std::min(start.x, end.x));
  io::set_attr(child_local(xfrm, "off"), "y", std::min(start.y, end.y));
  io::set_attr(child_local(xfrm, "ext"), "cx", end.x > start.x ? end.x - start.x : start.x - end.x);
  io::set_attr(child_local(xfrm, "ext"), "cy", end.y > start.y ? end.y - start.y : start.y - end.y);
  write_xfrm_transform(xfrm, XfrmTransform{0, end.x < start.x, end.y < start.y});
}

// Writes the anchor of the top-level connector `anchor` from its `a:xfrm`, keeping the anchor's kind.
void place_connector_anchor(const Geometry& geo, pugi::xml_node anchor, const pugi::xml_node& xfrm) {
  const DrawingObject obj = read_drawing_object(anchor, {});
  Placement p;
  p.kind = obj.anchor_kind;
  p.edit_as = obj.edit_as;
  p.off_x = emu(child_local(xfrm, "off"), "x");
  p.off_y = emu(child_local(xfrm, "off"), "y");
  p.cx = emu(child_local(xfrm, "ext"), "cx");
  p.cy = emu(child_local(xfrm, "ext"), "cy");
  const bool swap = swaps_box(xfrm);
  p.box_w = swap ? p.cy : p.cx;
  p.box_h = swap ? p.cx : p.cy;
  // The box shares the centre of the unrotated rectangle.
  const std::int64_t box_x = std::max<std::int64_t>(p.off_x - (p.box_w - p.cx) / 2, 0);
  const std::int64_t box_y = std::max<std::int64_t>(p.off_y - (p.box_h - p.cy) / 2, 0);
  if (p.kind == AnchorKind::kAbsolute) {
    p.from = AnchorPoint{0, 0, box_y, box_x};
  } else {
    (void)geo.advance(true, p.from.row, p.from.row_off, box_y);
    (void)geo.advance(false, p.from.col, p.from.col_off, box_x);
  }
  p.to = p.from;
  (void)geo.advance(true, p.to.row, p.to.row_off, p.box_h);
  (void)geo.advance(false, p.to.col, p.to.col_off, p.box_w);
  place_anchor(anchor, p);
}

bool in_group(const pugi::xml_node& content) {
  for (pugi::xml_node node = content.parent(); node; node = node.parent()) {
    if (local_name(node) == "grpSp") {
      return true;
    }
  }
  return false;
}

}  // namespace

std::vector<GluedEnd> glued_ends(const pugi::xml_node& root) {
  const std::vector<std::pair<std::uint32_t, pugi::xml_node>> targets = glue_targets(root);
  std::vector<GluedEnd> out;
  for_each_element(root, [&](const pugi::xml_node& node) {
    if (local_name(node) != "cxnSp") {
      return;
    }
    const pugi::xml_node cxn = child_local(child_local(node, "nvCxnSpPr"), "cNvCxnSpPr");
    const Frame own = frame_of(node);
    for (const bool start : {true, false}) {
      const pugi::xml_node glue = child_local(cxn, start ? "stCxn" : "endCxn");
      const std::optional<std::uint32_t> id = io::parse_xml_u32_attr_strict(glue.attribute("id"));
      const std::optional<std::uint32_t> idx = io::parse_xml_u32_attr_strict(glue.attribute("idx"));
      const auto target =
          std::find_if(targets.begin(), targets.end(), [&id](const auto& t) { return id && t.first == *id; });
      if (!idx || target == targets.end() || own.sx == 0.0 || own.sy == 0.0) {
        continue;
      }
      if (const std::optional<PlacedSite> site = place_site(target->second, *idx)) {
        GluedEnd glued;
        glued.connector = node;
        glued.start = start;
        glued.on_site = site->on_site;
        // An end at no site stays where the connector has it and counts as leaving rightwards.
        const Point at = site->on_site ? own.from_sheet(site->at) : end_point(node, start);
        glued.x = std::llround(at.x);
        glued.y = std::llround(at.y);
        // Frames scale without turning, so the direction is the sheet's, snapped to the nearest axis.
        const double quarters = std::fmod(site->angle / kQuarterTurn, 4.0);
        glued.direction = static_cast<int>(std::llround(quarters < 0.0 ? quarters + 4.0 : quarters) % 4);
        const Point a = own.from_sheet(site->box[0]);
        const Point b = own.from_sheet(site->box[1]);
        glued.left = std::min(a.x, b.x);
        glued.right = std::max(a.x, b.x);
        glued.top = std::min(a.y, b.y);
        glued.bottom = std::max(a.y, b.y);
        out.push_back(glued);
      }
    }
  });
  return out;
}

void follow_glued_shapes(const Workbook& wb, std::size_t sheet, const pugi::xml_node& root,
                         const std::vector<GluedEnd>& before) {
  const std::vector<GluedEnd> after = glued_ends(root);
  std::optional<Geometry> geo;
  std::vector<pugi::xml_node> done;
  for (const GluedEnd& end : before) {
    const pugi::xml_node connector = end.connector;
    if (std::find(done.begin(), done.end(), connector) != done.end()) {
      continue;
    }
    done.push_back(connector);
    const GluedEnd* start_after = find_end(after, connector, true);
    const GluedEnd* end_after = find_end(after, connector, false);
    const GluedEnd* start_before = find_end(before, connector, true);
    const GluedEnd* end_before = find_end(before, connector, false);
    const pugi::xml_node xfrm = child_local(child_local(connector, "spPr"), "xfrm");
    if (!xfrm) {
      continue;
    }
    if (start_after != nullptr && end_after != nullptr) {
      if (!changed(start_before, *start_after) && !changed(end_before, *end_after)) {
        continue;
      }
      route_connector(connector, xfrm, *start_after, *end_after);
    } else if (const GluedEnd* now = start_after != nullptr ? start_after : end_after) {
      // Glued at one end, the connector moves with it, its free end included.
      const GluedEnd* was = start_after != nullptr ? start_before : end_before;
      if (was == nullptr || !moved(was, *now)) {
        continue;
      }
      const pugi::xml_node off = child_local(xfrm, "off");
      io::set_attr(off, "x", emu(off, "x") + (now->x - was->x));
      io::set_attr(off, "y", emu(off, "y") + (now->y - was->y));
    } else {
      continue;
    }
    if (!in_group(connector)) {
      if (!geo) {
        geo.emplace(wb, wb.sheet(sheet));
      }
      place_connector_anchor(*geo, top_of(root, connector), xfrm);
    }
  }
}

}  // namespace drawing_edit
}  // namespace formulon
