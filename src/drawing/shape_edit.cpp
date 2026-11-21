#include "drawing/shape_edit.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <iterator>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "chart/chart_edit.h"
#include "chart/chart_edit_internal.h"
#include "chart/chart_ns.h"
#include "chart/chart_xml.h"
#include "drawing/connector_route.h"
#include "drawing/drawing_color.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_edit_internal.h"
#include "drawing/drawing_geometry.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml/preset_geometry.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {
namespace {

using chart_ns::kCtDrawing;
using chart_ns::kNsA;
using chart_ns::kNsXdr;
constexpr const char* kNoShape = "object_id does not name a shape";
constexpr std::string_view kFillNames[] = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"};

bool is_preset(std::string_view prst) {
  return preset_geometry::index_of(prst).has_value();
}

using drawing_edit::anchor_point_in_range;
using drawing_edit::in_emu_range;
using drawing_edit::invalid;
using drawing_edit::is_shape_kind;

bool is_fill(pugi::xml_node node) {
  return std::find(std::begin(kFillNames), std::end(kFillNames), local_name(node)) != std::end(kFillNames);
}

// Prefix (with its colon, empty for the default namespace) binding `uri` at `node`.
std::optional<std::string> bound_prefix(pugi::xml_node node, std::string_view uri) {
  for (pugi::xml_node n = node; n; n = n.parent()) {
    for (pugi::xml_attribute attr = n.first_attribute(); attr; attr = attr.next_attribute()) {
      const std::string_view name = attr.name();
      if (attr.value() == uri && (name == "xmlns" || strings::starts_with(name, "xmlns:"))) {
        return name == "xmlns" ? std::string() : std::string(name.substr(6)) + ":";
      }
    }
  }
  return std::nullopt;
}

// The `a:` prefix at `node`, declared on the drawing root when nothing binds it.
std::string a_prefix(pugi::xml_node node) {
  if (const auto bound = bound_prefix(node, kNsA)) {
    return *bound;
  }
  pugi::xml_node top = node.root().first_child();
  while (top && top.type() != pugi::node_element) {
    top = top.next_sibling();
  }
  top.append_attribute("xmlns:a").set_value(kNsA);
  return "a:";
}

std::string xdr_prefix(pugi::xml_node node) {
  return bound_prefix(node, kNsXdr).value_or("xdr:");
}

pugi::xml_node props_of(pugi::xml_node content) {
  return child_local(content, local_name(content) == "grpSp" ? "grpSpPr" : "spPr");
}

// Drops the glue at both ends of a connector; any other kind of `content` is left as it is.
void detach_connector(pugi::xml_node content) {
  pugi::xml_node cxn = child_local(child_local(content, "nvCxnSpPr"), "cNvCxnSpPr");
  cxn.remove_child(child_local(cxn, "stCxn"));
  cxn.remove_child(child_local(cxn, "endCxn"));
}

std::string read_text(pugi::xml_node content) {
  std::string out;
  bool first = true;
  for (pugi::xml_node p : children_local(child_local(content, "txBody"), "p")) {
    out += first ? "" : "\n";
    first = false;
    for (pugi::xml_node run = p.first_child(); run; run = run.next_sibling()) {
      const std::string_view name = local_name(run);
      if (name == "br") {
        out += '\n';
      } else if (name == "r" || name == "fld") {
        out += child_local(run, "t").text().get();
      }
    }
  }
  return out;
}

// One glued end; a missing or unreadable `id` or `idx` reads as a free end.
std::optional<ShapeConnection> read_connection(const pugi::xml_node& node) {
  const std::optional<std::uint32_t> id = io::parse_xml_u32_attr_strict(node.attribute("id"));
  const std::optional<std::uint32_t> idx = io::parse_xml_u32_attr_strict(node.attribute("idx"));
  if (!id || !idx) {
    return std::nullopt;
  }
  return ShapeConnection{*id, *idx};
}

constexpr const char* kTextVertTokens[] = {"horz",        "vert",           "vert270",      "eaVert",
                                           "wordArtVert", "wordArtVertRtl", "mongolianVert"};

std::optional<ShapeTextVert> read_text_vert(const pugi::xml_node& content) {
  const pugi::xml_attribute vert = child_local(child_local(content, "txBody"), "bodyPr").attribute("vert");
  for (std::size_t i = 0; vert && i < std::size(kTextVertTokens); ++i) {
    if (std::string_view(vert.value()) == kTextVertTokens[i]) {
      return static_cast<ShapeTextVert>(i);
    }
  }
  return std::nullopt;
}

/// The value of a guide formula `val <int32>`; unset for any other formula.
std::optional<std::int32_t> adjust_value_of(std::string_view fmla) {
  constexpr std::string_view kVal = "val ";
  if (fmla.substr(0, kVal.size()) != kVal) {
    return std::nullopt;
  }
  std::string_view digits = fmla.substr(kVal.size());
  const bool negative = !digits.empty() && digits.front() == '-';
  digits.remove_prefix(negative ? 1U : 0U);
  if (digits.empty() || digits.size() > 10U) {
    return std::nullopt;
  }
  std::int64_t v = 0;
  for (const char c : digits) {
    if (!strings::is_ascii_digit(c)) {
      return std::nullopt;
    }
    v = v * 10 + (c - '0');
  }
  v = negative ? -v : v;
  if (v < INT32_MIN || v > INT32_MAX) {
    return std::nullopt;
  }
  return static_cast<std::int32_t>(v);
}

// The `val` guides of `geom` its preset defines, each once; the save leaves any other guide out.
std::vector<ShapeAdjust> read_adjust_values(const pugi::xml_node& geom) {
  std::vector<ShapeAdjust> out;
  const std::optional<std::size_t> preset = preset_geometry::index_of(geom.attribute("prst").value());
  if (!preset) {
    return out;
  }
  const std::vector<preset_geometry::AdjustGuide> defined = preset_geometry::adjust_guides(*preset);
  std::vector<std::string_view> seen;
  for (const pugi::xml_node& gd : children_local(child_local(geom, "avLst"), "gd")) {
    const std::string_view name = gd.attribute("name").value();
    const bool known = std::any_of(defined.begin(), defined.end(),
                                   [name](const preset_geometry::AdjustGuide& g) { return g.name == name; });
    const bool repeated = std::find(seen.begin(), seen.end(), name) != seen.end();
    seen.push_back(name);
    if (!known || repeated) {
      continue;
    }
    if (const std::optional<std::int32_t> v = adjust_value_of(gd.attribute("fmla").value())) {
      out.push_back(ShapeAdjust{std::string(name), *v});
    }
  }
  return out;
}

Expected<void, Error> check_adjust_values(const std::vector<ShapeAdjustPatch>& patches) {
  if (patches.size() > kMaxAdjustPatches) {
    return make_error(FormulonErrorCode::kSecurityResourceLimit, "adjust_values holds too many guides",
                      limit_context("adjust_values", patches.size(), kMaxAdjustPatches));
  }
  for (std::size_t i = 0; i < patches.size(); ++i) {
    if (!drawing_color::is_drawing_identifier(patches[i].name)) {
      return invalid("An adjust value name must be [A-Za-z][A-Za-z0-9]* of at most 32 characters");
    }
    for (std::size_t j = 0; j < i; ++j) {
      if (patches[j].name == patches[i].name) {
        return invalid("adjust_values names a guide twice");
      }
    }
  }
  return Expected<void, Error>::Ok();
}

ShapeInfo read_shape(const Workbook& wb, pugi::xml_node anchor, DrawingObject object) {
  ShapeInfo info;
  info.object = std::move(object);
  pugi::xml_node content = anchor_content(anchor);
  pugi::xml_node props = props_of(content);
  info.prst = child_local(props, "prstGeom").attribute("prst").value();
  info.adjust_values = read_adjust_values(child_local(props, "prstGeom"));
  info.text = read_text(content);
  info.text_box = child_local(child_local(content, "nvSpPr"), "cNvSpPr").attribute("txBox").as_bool();
  info.text_vert = read_text_vert(content);
  const drawing_edit::XfrmTransform transform = drawing_edit::read_xfrm_transform(child_local(props, "xfrm"));
  info.rotation = transform.rot;
  info.flip_h = transform.flip_h;
  info.flip_v = transform.flip_v;
  info.paint_fill = read_paint_fill(wb, props);
  info.paint_line = read_paint_line(wb, props);
  const pugi::xml_node cxn = child_local(child_local(content, "nvCxnSpPr"), "cNvCxnSpPr");
  info.start_connection = read_connection(child_local(cxn, "stCxn"));
  info.end_connection = read_connection(child_local(cxn, "endCxn"));
  return info;
}

pugi::xml_node find_shape(pugi::xml_node root, std::uint32_t object_id) {
  return find_drawing_object(root, object_id, is_shape_kind);
}

// The `child` element of `parent`, created (first or last) when absent.
pugi::xml_node ensure_child(pugi::xml_node parent, const std::string& qname, bool first = false) {
  pugi::xml_node found = child_local(parent, local_name(qname.c_str()));
  if (found) {
    return found;
  }
  return first ? parent.prepend_child(qname.c_str()) : parent.append_child(qname.c_str());
}

// The child of `parent` after which a new fill, outline or geometry belongs:
// the last of the xfrm, geometry and (for an outline) fill elements.
pugi::xml_node last_leading_child(pugi::xml_node parent, bool past_fill) {
  pugi::xml_node last;
  for (pugi::xml_node child = parent.first_child(); child; child = child.next_sibling()) {
    const std::string_view name = local_name(child);
    if (name == "xfrm" || name == "prstGeom" || name == "custGeom" || (past_fill && is_fill(child))) {
      last = child;
    }
  }
  return last;
}

pugi::xml_node insert_after_or_first(pugi::xml_node parent, pugi::xml_node ref, const std::string& qname) {
  return ref ? parent.insert_child_after(qname.c_str(), ref) : parent.prepend_child(qname.c_str());
}

pugi::xml_node ensure_xfrm(pugi::xml_node props) {
  if (pugi::xml_node xfrm = child_local(props, "xfrm")) {
    return xfrm;
  }
  const std::string ap = a_prefix(props);
  pugi::xml_node xfrm = props.prepend_child((ap + "xfrm").c_str());
  xfrm.append_child((ap + "off").c_str());
  xfrm.append_child((ap + "ext").c_str());
  return xfrm;
}

// Turns `anchor` into an empty anchor of `kind`: its tag, `editAs` and position children change.
void convert_anchor(pugi::xml_node anchor, AnchorKind kind) {
  const std::string xp = xdr_prefix(anchor);
  anchor.set_name((xp + (kind == AnchorKind::kOneCell   ? "oneCellAnchor"
                         : kind == AnchorKind::kTwoCell ? "twoCellAnchor"
                                                        : "absoluteAnchor"))
                      .c_str());
  anchor.remove_attribute("editAs");
  for (const char* name : {"pos", "from", "to", "ext"}) {
    anchor.remove_child(child_local(anchor, name));
  }
  // Prepended last-first, so the children end up as from + to, from + ext or pos + ext.
  const auto prepend_marker = [&](const char* tag) {
    pugi::xml_node marker = anchor.prepend_child((xp + tag).c_str());
    for (const char* part : {"col", "colOff", "row", "rowOff"}) {
      marker.append_child((xp + part).c_str()).text().set(0);
    }
  };
  if (kind == AnchorKind::kTwoCell) {
    prepend_marker("to");
  } else {
    anchor.prepend_child((xp + "ext").c_str());
  }
  if (kind == AnchorKind::kAbsolute) {
    anchor.prepend_child((xp + "pos").c_str());
  } else {
    prepend_marker("from");
  }
}

// The top-left corner of `anchor` in EMU from the sheet's origin.
std::pair<std::int64_t, std::int64_t> anchor_origin(const Workbook& wb, std::size_t sheet, pugi::xml_node anchor) {
  const DrawingObject obj = read_drawing_object(anchor, {});
  if (obj.anchor_kind == AnchorKind::kAbsolute) {
    return {obj.from.col_off, obj.from.row_off};
  }
  const Geometry geo(wb, wb.sheet(sheet));
  return {geo.offset(false, obj.from.col) + obj.from.col_off, geo.offset(true, obj.from.row) + obj.from.row_off};
}

// Writes the anchor fields of `spec` into the existing `anchor` and the matching `a:xfrm`, converting the
// anchor's kind when `spec` names another one.
Expected<void, Error> place(const Workbook& wb, std::size_t sheet, pugi::xml_node anchor, pugi::xml_node content,
                            const ShapeSpec& spec) {
  const bool absolute = spec.anchor_kind == AnchorKind::kAbsolute;
  const bool two_cell = spec.anchor_kind == AnchorKind::kTwoCell;
  if (absolute ? !in_emu_range(spec.from.row_off) || !in_emu_range(spec.from.col_off)
               : !anchor_point_in_range(spec.from)) {
    return invalid("anchor cells and offsets must be in range");
  }
  if (two_cell ? !anchor_point_in_range(spec.to) || spec.to.row < spec.from.row || spec.to.col < spec.from.col
               : !in_emu_range(spec.cx) || !in_emu_range(spec.cy)) {
    return invalid("anchor box and size must be in range");
  }
  if (read_drawing_object(anchor, {}).anchor_kind != spec.anchor_kind) {
    convert_anchor(anchor, spec.anchor_kind);
  }
  std::int64_t off_x = spec.from.col_off;
  std::int64_t off_y = spec.from.row_off;
  std::int64_t cx = spec.cx;
  std::int64_t cy = spec.cy;
  std::optional<Geometry> geo;
  if (absolute) {
    pugi::xml_node pos = child_local(anchor, "pos");
    io::set_attr(pos, "x", off_x);
    io::set_attr(pos, "y", off_y);
  } else {
    geo.emplace(wb, wb.sheet(sheet));
    write_anchor_point(child_local(anchor, "from"), spec.from);
    off_x += geo->offset(false, spec.from.col);
    off_y += geo->offset(true, spec.from.row);
  }
  if (two_cell) {
    write_anchor_point(child_local(anchor, "to"), spec.to);
    cx = geo->offset(false, spec.to.col) + spec.to.col_off - off_x;
    cy = geo->offset(true, spec.to.row) + spec.to.row_off - off_y;
    pugi::xml_attribute edit_as = anchor.attribute("editAs");
    if (spec.edit_as == EditAs::kTwoCell) {
      anchor.remove_attribute(edit_as);
    } else {
      (edit_as ? edit_as : anchor.append_attribute("editAs"))
          .set_value(spec.edit_as == EditAs::kOneCell ? "oneCell" : "absolute");
    }
  } else {
    pugi::xml_node ext = child_local(anchor, "ext");
    io::set_attr(ext, "cx", cx);
    io::set_attr(ext, "cy", cy);
  }
  pugi::xml_node xfrm = ensure_xfrm(props_of(content));
  const std::string ap = a_prefix(xfrm);
  pugi::xml_node off = ensure_child(xfrm, ap + "off", true);
  io::set_attr(off, "x", off_x);
  io::set_attr(off, "y", off_y);
  pugi::xml_node ext = ensure_child(xfrm, ap + "ext");
  io::set_attr(ext, "cx", std::max<std::int64_t>(cx, 0));
  io::set_attr(ext, "cy", std::max<std::int64_t>(cy, 0));
  return Expected<void, Error>::Ok();
}

/// The shape's `txBody`, created with one empty paragraph when absent.
pugi::xml_node ensure_text_body(pugi::xml_node content) {
  pugi::xml_node body = child_local(content, "txBody");
  if (!body) {
    const std::string ap = a_prefix(content);
    body = content.append_child((xdr_prefix(content) + "txBody").c_str());
    pugi::xml_node pr = body.append_child((ap + "bodyPr").c_str());
    pr.append_attribute("vertOverflow").set_value("clip");
    pr.append_attribute("wrap").set_value("square");
    pr.append_attribute("rtlCol").set_value("0");
    pr.append_attribute("anchor").set_value("t");
    body.append_child((ap + "lstStyle").c_str());
    body.append_child((ap + "p").c_str());
  }
  return body;
}

void set_text_vert(pugi::xml_node content, std::optional<ShapeTextVert> vert) {
  if (!vert && !child_local(content, "txBody")) {
    return;
  }
  pugi::xml_node body = ensure_text_body(content);
  pugi::xml_node pr = child_local(body, "bodyPr");
  if (!pr) {
    pr = body.prepend_child((a_prefix(content) + "bodyPr").c_str());
  }
  if (!vert) {
    pr.remove_attribute("vert");
    return;
  }
  (pr.attribute("vert") ? pr.attribute("vert") : pr.append_attribute("vert"))
      .set_value(kTextVertTokens[static_cast<std::size_t>(*vert)]);
}

void set_text(pugi::xml_node content, const std::string& text) {
  const std::string ap = a_prefix(content);
  pugi::xml_node body = ensure_text_body(content);
  // The first paragraph's properties and first run's formatting serve every new paragraph.
  pugi::xml_node first = child_local(body, "p");
  pugi::xml_node ppr = child_local(first, "pPr");
  pugi::xml_node rpr = child_local(child_local(first, "r"), "rPr");
  if (!rpr) {
    rpr = child_local(first, "endParaRPr");
  }
  std::vector<pugi::xml_node> old = children_local(body, "p");
  std::size_t start = 0;
  while (true) {
    const std::size_t end = text.find('\n', start);
    pugi::xml_node p = body.append_child((ap + "p").c_str());
    if (ppr) {
      p.append_copy(ppr);
    }
    const std::string line = text.substr(start, end == std::string::npos ? end : end - start);
    if (!line.empty()) {
      pugi::xml_node run = p.append_child((ap + "r").c_str());
      if (rpr) {
        run.append_copy(rpr).set_name((ap + "rPr").c_str());
      }
      run.append_child((ap + "t").c_str()).text().set(line.c_str());
    } else if (rpr) {
      p.append_copy(rpr).set_name((ap + "endParaRPr").c_str());
    }
    if (end == std::string::npos) {
      break;
    }
    start = end + 1U;
  }
  for (pugi::xml_node p : old) {
    body.remove_child(p);
  }
}

// Writes the chart-model fill and outline patches of `spec` into `props`.
Expected<void, Error> apply_paint_records(pugi::xml_node props, const ShapeSpec& spec) {
  if (chart_edit::fill_empty(spec.paint_fill) && chart_edit::line_empty(spec.paint_line)) {
    return Expected<void, Error>::Ok();
  }
  // The shared writer spells DrawingML elements `a:`; bind it locally when the part uses another prefix.
  if (bound_prefix(props, kNsA).value_or(std::string()) != "a:") {
    props.append_attribute("xmlns:a").set_value(kNsA);
  }
  RETURN_IF_ERROR(write_paint_properties(props, spec.paint_fill, spec.paint_line));
  // An `a:ln` left with nothing in it states nothing.
  pugi::xml_node ln = child_local(props, "ln");
  if (ln && !ln.first_child() && !ln.first_attribute()) {
    props.remove_child(ln);
  }
  return Expected<void, Error>::Ok();
}

/// Sets or removes each named guide of `geom`'s `a:avLst`; a new guide goes last.
void apply_adjust_values(pugi::xml_node geom, const std::string& ap, const std::vector<ShapeAdjustPatch>& patches) {
  pugi::xml_node guides = ensure_child(geom, ap + "avLst");
  for (const ShapeAdjustPatch& patch : patches) {
    pugi::xml_node gd = guides.find_child([&patch](const pugi::xml_node& n) {
      return local_name(n) == "gd" && patch.name == n.attribute("name").value();
    });
    if (!patch.value) {
      guides.remove_child(gd);
      continue;
    }
    if (!gd) {
      gd = guides.append_child((ap + "gd").c_str());
      gd.append_attribute("name").set_value(patch.name.c_str());
    }
    const std::string fmla = "val " + std::to_string(*patch.value);
    (gd.attribute("fmla") ? gd.attribute("fmla") : gd.append_attribute("fmla")).set_value(fmla.c_str());
  }
}

// Applies the non-anchor fields of `spec` to `content`; `anchor` names its `cNvPr`.
Expected<void, Error> apply_fields(pugi::xml_node anchor, pugi::xml_node content, const ShapeSpec& spec) {
  const std::uint32_t f = spec.fields;
  const std::string_view kind = local_name(content);
  const bool group = kind == "grpSp";
  pugi::xml_node props = props_of(content);
  if (!props) {
    return invalid("Shape has no shape properties");
  }
  if ((f & shape_field::kPrst) != 0U && (group || !is_preset(spec.prst))) {
    return invalid("preset_geometry must be a DrawingML shape type on a shape or connector");
  }
  RETURN_IF_ERROR(check_adjust_values(spec.adjust_values));
  if (!spec.adjust_values.empty()) {
    const pugi::xml_node geom = child_local(props, "prstGeom");
    if (!geom && (f & shape_field::kPrst) == 0U) {
      return invalid("adjust_values apply only to preset geometry");
    }
    // Excel refuses to open a file whose avLst names a guide its preset does not define.
    const std::string_view prst =
        (f & shape_field::kPrst) != 0U ? std::string_view(spec.prst) : std::string_view(geom.attribute("prst").value());
    const std::optional<std::size_t> preset = preset_geometry::index_of(prst);
    const std::vector<preset_geometry::AdjustGuide> guides =
        preset ? preset_geometry::adjust_guides(*preset) : std::vector<preset_geometry::AdjustGuide>{};
    for (const ShapeAdjustPatch& patch : spec.adjust_values) {
      if (std::none_of(guides.begin(), guides.end(),
                       [&patch](const preset_geometry::AdjustGuide& g) { return g.name == patch.name; })) {
        return invalid("adjust_values names a guide preset_geometry does not define");
      }
    }
  }
  if ((f & (shape_field::kText | shape_field::kTextBox)) != 0U && kind != "sp") {
    return invalid("text and text_box apply only to a shape");
  }
  if ((f & shape_field::kTextVert) != 0U && kind != "sp") {
    return invalid("text_vert applies only to a shape");
  }
  if ((f & shape_field::kTextVert) != 0U && spec.text_vert &&
      static_cast<std::size_t>(*spec.text_vert) >= std::size(kTextVertTokens)) {
    return invalid("text_vert must be a DrawingML text direction");
  }
  if ((f & shape_field::kText) != 0U && std::any_of(spec.text.begin(), spec.text.end(), [](unsigned char c) {
        return c < 0x20U && c != '\n' && c != '\t';
      })) {
    return invalid("text must not hold a control character");
  }
  pugi::xml_node cnvpr = anchor_cnvpr(anchor);
  if ((f & shape_field::kName) != 0U) {
    (cnvpr.attribute("name") ? cnvpr.attribute("name") : cnvpr.append_attribute("name")).set_value(spec.name.c_str());
  }
  if ((f & shape_field::kHidden) != 0U) {
    drawing_edit::set_cnvpr_hidden(cnvpr, spec.hidden);
  }
  if ((f & shape_field::kTextBox) != 0U) {
    pugi::xml_node cnv = child_local(child_local(content, "nvSpPr"), "cNvSpPr");
    if (spec.text_box) {
      (cnv.attribute("txBox") ? cnv.attribute("txBox") : cnv.append_attribute("txBox")).set_value("1");
    } else {
      cnv.remove_attribute("txBox");
    }
  }
  if ((f & shape_field::kPrst) != 0U) {
    pugi::xml_node geom = child_local(props, "prstGeom");
    if (!geom) {
      pugi::xml_node custom = child_local(props, "custGeom");
      geom = insert_after_or_first(props, custom ? custom : last_leading_child(props, false),
                                   a_prefix(props) + "prstGeom");
      props.remove_child(custom);
    }
    if (spec.prst != geom.attribute("prst").value()) {
      pugi::xml_node guides = ensure_child(geom, a_prefix(props) + "avLst");
      while (guides.first_child()) {
        guides.remove_child(guides.first_child());
      }
      (geom.attribute("prst") ? geom.attribute("prst") : geom.append_attribute("prst")).set_value(spec.prst.c_str());
    }
  }
  if (!spec.adjust_values.empty()) {
    apply_adjust_values(child_local(props, "prstGeom"), a_prefix(props), spec.adjust_values);
  }
  if ((f & (shape_field::kRotation | shape_field::kFlipH | shape_field::kFlipV)) != 0U) {
    const drawing_edit::XfrmTransform current = drawing_edit::read_xfrm_transform(child_local(props, "xfrm"));
    drawing_edit::XfrmTransform wanted = current;
    if ((f & shape_field::kRotation) != 0U) {
      wanted.rot = spec.rotation;
    }
    if ((f & shape_field::kFlipH) != 0U) {
      wanted.flip_h = spec.flip_h;
    }
    if ((f & shape_field::kFlipV) != 0U) {
      wanted.flip_v = spec.flip_v;
    }
    if (wanted != current) {
      drawing_edit::write_xfrm_transform(ensure_xfrm(props), wanted);
      // Excel detaches a connector at both ends when it is flipped or rotated.
      detach_connector(content);
    }
  }
  if ((f & shape_field::kText) != 0U) {
    set_text(content, spec.text);
  }
  if ((f & shape_field::kTextVert) != 0U) {
    set_text_vert(content, spec.text_vert);
  }
  return apply_paint_records(props, spec);
}

bool is_connector_prst(const std::string& prst) {
  return prst == "line" || prst == "straightConnector1" || strings::starts_with(prst, "bentConnector") ||
         strings::starts_with(prst, "curvedConnector");
}

void add_marker(pugi::xml_node parent, const std::string& xp, const char* tag) {
  pugi::xml_node marker = parent.append_child((xp + tag).c_str());
  for (const char* part : {"col", "colOff", "row", "rowOff"}) {
    marker.append_child((xp + part).c_str()).text().set(0);
  }
}

}  // namespace

Expected<std::vector<ShapeInfo>, Error> list_shapes(const Workbook& wb, std::size_t sheet) {
  ASSIGN_OR_RETURN(std::string path, sheet_drawing_path(wb, sheet));
  std::vector<ShapeInfo> shapes;
  if (path.empty()) {
    return shapes;
  }
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
  for (pugi::xml_node anchor : drawing_anchors(doc.document_element(), false)) {
    DrawingObject obj = read_drawing_object(anchor, {});
    if (is_shape_kind(obj.kind)) {
      shapes.push_back(read_shape(wb, anchor, std::move(obj)));
    }
  }
  return shapes;
}

Expected<std::uint32_t, Error> insert_shape(Workbook& wb, std::size_t sheet, const ShapeSpec& spec) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if ((spec.fields & shape_field::kAnchor) == 0U) {
    return invalid("placement.anchor is required");
  }
  const std::string prst = (spec.fields & shape_field::kPrst) != 0U && !spec.prst.empty() ? spec.prst : "rect";
  if (!is_preset(prst)) {
    return invalid("preset_geometry must be a DrawingML shape type");
  }
  const bool connector = is_connector_prst(prst);
  const bool fresh = path.empty();
  pugi::xml_document doc;
  if (fresh) {
    path = "xl/drawings/drawing1.xml";
    for (std::uint32_t n = 2; io::ooxml::find_passthrough_part(wb, path) != nullptr; ++n) {
      path = "xl/drawings/drawing" + std::to_string(n) + ".xml";
    }
    doc.append_child("xdr:wsDr").append_attribute("xmlns:xdr").set_value(kNsXdr);
  } else {
    RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
  }
  pugi::xml_node root = doc.document_element();
  const std::string xp = xdr_prefix(root);
  const std::string ap = a_prefix(root);
  ASSIGN_OR_RETURN(const std::uint32_t id, drawing_edit::allocate_drawing_id(wb.sheet(sheet), doc));

  const char* anchor_tag = spec.anchor_kind == AnchorKind::kOneCell   ? "oneCellAnchor"
                           : spec.anchor_kind == AnchorKind::kTwoCell ? "twoCellAnchor"
                                                                      : "absoluteAnchor";
  pugi::xml_node anchor = root.append_child((xp + anchor_tag).c_str());
  if (spec.anchor_kind == AnchorKind::kAbsolute) {
    anchor.append_child((xp + "pos").c_str());
  } else {
    add_marker(anchor, xp, "from");
    if (spec.anchor_kind == AnchorKind::kTwoCell) {
      add_marker(anchor, xp, "to");
    }
  }
  if (spec.anchor_kind != AnchorKind::kTwoCell) {
    anchor.append_child((xp + "ext").c_str());
  }
  pugi::xml_node content = anchor.append_child((xp + (connector ? "cxnSp" : "sp")).c_str());
  content.append_attribute("macro").set_value("");
  if (!connector) {
    content.append_attribute("textlink").set_value("");
  }
  pugi::xml_node nv = content.append_child((xp + (connector ? "nvCxnSpPr" : "nvSpPr")).c_str());
  pugi::xml_node cnvpr = nv.append_child((xp + "cNvPr").c_str());
  cnvpr.append_attribute("id").set_value(id);
  cnvpr.append_attribute("name").set_value(
      ((spec.fields & shape_field::kName) != 0U ? spec.name : "Shape " + std::to_string(id)).c_str());
  nv.append_child((xp + (connector ? "cNvCxnSpPr" : "cNvSpPr")).c_str());
  pugi::xml_node props = content.append_child((xp + "spPr").c_str());
  pugi::xml_node geom = props.append_child((ap + "prstGeom").c_str());
  geom.append_attribute("prst").set_value(prst.c_str());
  geom.append_child((ap + "avLst").c_str());
  anchor.append_child((xp + "clientData").c_str());

  ShapeSpec rest = spec;
  rest.fields &= ~(shape_field::kName | shape_field::kPrst | shape_field::kAnchor);
  auto placed = place(wb, sheet, anchor, content, spec);
  if (placed) {
    placed = apply_fields(anchor, content, rest);
  }
  if (!placed) {
    return std::move(placed.error());
  }
  if (fresh) {
    RETURN_IF_ERROR(wb.add_passthrough_part(PassthroughPart(path, kCtDrawing, {})));
  }
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  if (fresh) {
    wb.sheet(sheet).set_drawing_rel_target(path);
  }
  return id;
}

namespace {

// Applies `spec` to the shape with `object_id` in a parsed copy of the drawing, storing the part only when
// `store` is set; with it unset the call is the full check of the edit.
Expected<void, Error> edit_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id, const ShapeSpec& spec,
                                 bool store) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  pugi::xml_document doc;
  if (path.empty()) {
    return drawing_edit::not_found(kNoShape);
  }
  RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
  pugi::xml_node anchor = find_shape(doc.document_element(), object_id);
  if (!anchor) {
    return drawing_edit::refuse_object(doc.document_element(), object_id, kNoShape);
  }
  pugi::xml_node content = anchor_content(anchor);
  const std::vector<drawing_edit::GluedEnd> glued = drawing_edit::glued_ends(doc.document_element());
  if ((spec.fields & shape_field::kAnchor) != 0U) {
    const std::pair<std::int64_t, std::int64_t> before = anchor_origin(wb, sheet, anchor);
    RETURN_IF_ERROR(place(wb, sheet, anchor, content, spec));
    // Excel detaches a connector that is moved; resizing it from a fixed top-left keeps its glue.
    if (anchor_origin(wb, sheet, anchor) != before) {
      detach_connector(content);
    }
  }
  ShapeSpec rest = spec;
  rest.fields &= ~shape_field::kAnchor;
  RETURN_IF_ERROR(apply_fields(anchor, content, rest));
  drawing_edit::follow_glued_shapes(wb, sheet, doc.document_element(), glued);
  if (!store) {
    return Expected<void, Error>::Ok();
  }
  return io::ooxml::store_part_dom(wb, path, doc);
}

}  // namespace

Expected<void, Error> update_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id, const ShapeSpec& spec) {
  return edit_shape(wb, sheet, object_id, spec, true);
}

Expected<void, Error> update_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                   const DrawingCommonPatch* common, const ShapeSpec* shape) {
  if (shape == nullptr) {
    return common != nullptr ? update_drawing(wb, sheet, object_id, *common) : Expected<void, Error>::Ok();
  }
  RETURN_IF_ERROR(edit_shape(wb, sheet, object_id, *shape, false));
  if (common == nullptr) {
    return edit_shape(wb, sheet, object_id, *shape, true);
  }
  // The common half checks itself before it writes; the shape half, checked above, then only writes. Should
  // that write still fail, the drawing goes back to its bytes before the common half.
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  std::vector<std::uint8_t> before;
  if (const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path)) {
    before = part->bytes;
  }
  RETURN_IF_ERROR(update_drawing(wb, sheet, object_id, *common));
  if (auto applied = edit_shape(wb, sheet, object_id, *shape, true); !applied) {
    (void)wb.replace_passthrough_part(path, std::move(before));
    return applied;
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> remove_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  pugi::xml_document doc;
  if (path.empty()) {
    return drawing_edit::not_found(kNoShape);
  }
  RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
  pugi::xml_node root = doc.document_element();
  pugi::xml_node target = find_shape(root, object_id);
  if (!target) {
    return drawing_edit::refuse_object(root, object_id, kNoShape);
  }
  return drawing_edit::drop_objects(wb, sheet, path, doc, {drawing_edit::top_of(root, target)});
}

}  // namespace formulon
