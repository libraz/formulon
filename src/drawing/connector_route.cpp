#include "drawing/connector_route.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <optional>
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

using preset_geometry::Point;

constexpr double kRadiansPerUnit = 3.14159265358979323846 / (180.0 * 60000.0);

std::int64_t emu(const pugi::xml_node& node, const char* name) {
  return node.attribute(name).as_llong();
}

double emu_d(const pugi::xml_node& node, const char* name) {
  return static_cast<double>(emu(node, name));
}

// Maps a point of a frame to the sheet: per axis, `scale * p + shift`.
struct Frame {
  double sx = 1.0, sy = 1.0, tx = 0.0, ty = 0.0;
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

// Site `idx` of the shape or picture `content` on the sheet: the preset's site on the current size and stated
// adjust guides, flipped and rotated with the object, then placed through its frame.
std::optional<Point> site_on_sheet(const pugi::xml_node& content, std::uint32_t idx) {
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
  const std::vector<Point> sites = preset_geometry::connection_sites(*preset, w, h, stated);
  if (idx >= sites.size()) {
    return std::nullopt;
  }
  Point p = sites[idx];
  const XfrmTransform t = read_xfrm_transform(xfrm);
  p.x = t.flip_h ? w - p.x : p.x;
  p.y = t.flip_v ? h - p.y : p.y;
  if (t.rot != 0) {
    const double angle = static_cast<double>(t.rot) * kRadiansPerUnit;
    const double c = elementary::cos(angle);
    const double s = elementary::sin(angle);
    const double dx = p.x - w / 2.0;
    const double dy = p.y - h / 2.0;
    p = Point{w / 2.0 + dx * c - dy * s, h / 2.0 + dx * s + dy * c};
  }
  const pugi::xml_node off = child_local(xfrm, "off");
  const Frame f = frame_of(content);
  return Point{f.sx * (p.x + emu_d(off, "x")) + f.tx, f.sy * (p.y + emu_d(off, "y")) + f.ty};
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

bool moved(const GluedEnd* before, const GluedEnd& after) {
  return before == nullptr || before->x != after.x || before->y != after.y;
}

// Both ends glued: the connector runs from its start site to its end site, its flips giving the direction. A
// bent or curved connector keeps its preset and adjust values.
void route_connector(pugi::xml_node xfrm, const GluedEnd& start, const GluedEnd& end) {
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
      if (const std::optional<Point> site = site_on_sheet(target->second, *idx)) {
        out.push_back(GluedEnd{node, start, std::llround((site->x - own.tx) / own.sx),
                               std::llround((site->y - own.ty) / own.sy)});
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
      if (!moved(start_before, *start_after) && !moved(end_before, *end_after)) {
        continue;
      }
      route_connector(xfrm, *start_after, *end_after);
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
