#include "drawing/shape_edit.h"

#include <algorithm>
#include <cctype>
#include <cstdlib>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "drawing/drawing_edit.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_defs.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/number_text.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr const char* kCtDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";
constexpr const char* kNsXdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
constexpr const char* kNsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
constexpr std::int64_t kEmuLimit = std::int64_t{1} << 31;
constexpr std::string_view kFillNames[] = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"};

Error invalid(const char* message) {
  return make_error(FormulonErrorCode::kInvalidArgument, message);
}

bool is_shape_kind(DrawingObjectKind kind) {
  return kind == DrawingObjectKind::kShape || kind == DrawingObjectKind::kConnector ||
         kind == DrawingObjectKind::kGroup;
}

bool in_range(std::int64_t value) {
  return value >= 0 && value < kEmuLimit;
}

bool is_fill(pugi::xml_node node) {
  return std::find(std::begin(kFillNames), std::end(kFillNames), local_name(node)) != std::end(kFillNames);
}

// Prefix (with its colon, empty for the default namespace) binding `uri` at `node`.
std::optional<std::string> bound_prefix(pugi::xml_node node, std::string_view uri) {
  for (pugi::xml_node n = node; n; n = n.parent()) {
    for (pugi::xml_attribute attr = n.first_attribute(); attr; attr = attr.next_attribute()) {
      const std::string_view name = attr.name();
      if (attr.value() == uri && (name == "xmlns" || name.rfind("xmlns:", 0) == 0)) {
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

bool parse_rgb(std::string_view text, std::uint32_t& rgb) {
  if (text.size() != 6U || text.find_first_not_of("0123456789abcdefABCDEF") != std::string_view::npos) {
    return false;
  }
  rgb = static_cast<std::uint32_t>(std::strtoul(std::string(text).c_str(), nullptr, 16));
  return true;
}

pugi::xml_node props_of(pugi::xml_node content) {
  return child_local(content, local_name(content) == "grpSp" ? "grpSpPr" : "spPr");
}

void read_paint(pugi::xml_node parent, ShapePaint& kind, std::uint32_t& rgb) {
  for (pugi::xml_node child = parent.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element || !is_fill(child)) {
      continue;
    }
    const std::string_view name = local_name(child);
    kind = name == "noFill" ? ShapePaint::kNone : ShapePaint::kOther;
    if (name == "solidFill" && parse_rgb(child_local(child, "srgbClr").attribute("val").value(), rgb)) {
      kind = ShapePaint::kSolid;
    }
    return;
  }
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

ShapeInfo read_shape(pugi::xml_node anchor, DrawingObject object) {
  ShapeInfo info;
  info.object = std::move(object);
  pugi::xml_node content = anchor_content(anchor);
  pugi::xml_node props = props_of(content);
  info.prst = child_local(props, "prstGeom").attribute("prst").value();
  info.text = read_text(content);
  info.text_box = child_local(child_local(content, "nvSpPr"), "cNvSpPr").attribute("txBox").as_bool();
  read_paint(props, info.fill, info.fill_rgb);
  pugi::xml_node ln = child_local(props, "ln");
  read_paint(ln, info.line, info.line_rgb);
  info.line_width_emu = ln.attribute("w").as_llong();
  info.rotation = child_local(props, "xfrm").attribute("rot").as_int();
  info.hidden = anchor_cnvpr(anchor).attribute("hidden").as_bool();
  return info;
}

Expected<void, Error> load_drawing(const Workbook& wb, const std::string& path, pugi::xml_document& doc) {
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path);
  if (part == nullptr) {
    return make_error(FormulonErrorCode::kIoDrawingUnparseable, "drawing: part is missing", path);
  }
  return parse_drawing_part(part->bytes, doc);
}

Expected<std::string, Error> editable_path(const Workbook& wb, std::size_t sheet) {
  if (sheet >= wb.sheet_count()) {
    return invalid("drawing: sheet index out of range");
  }
  const Sheet& ws = wb.sheet(sheet);
  for (const UnknownRelationship& rel : ws.unknown_relationships()) {
    if (rel.type == io::kRelDrawing) {
      return make_error(FormulonErrorCode::kIoDrawingUnparseable,
                        "drawing: drawing retained from an xlsb package is not editable", ws.name());
    }
  }
  if (ws.is_opaque_ooxml_sheet()) {
    return invalid("drawing: only a worksheet's drawing is editable");
  }
  return ws.drawing_rel_target();
}

pugi::xml_node find_shape(pugi::xml_node root, std::uint32_t object_id) {
  for (pugi::xml_node anchor : drawing_anchors(root, false)) {
    const DrawingObject obj = read_drawing_object(anchor, {});
    if (is_shape_kind(obj.kind) && obj.object_id == object_id) {
      return anchor;
    }
  }
  return {};
}

void set_llong_attr(pugi::xml_node node, const char* name, std::int64_t value) {
  pugi::xml_attribute attr = node.attribute(name);
  (attr ? attr : node.append_attribute(name)).set_value(static_cast<long long>(value));
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

// Sets the fill element of `parent` (a shape-properties or `a:ln` element).
void set_paint(pugi::xml_node parent, bool is_line, bool none, std::uint32_t rgb) {
  const std::string ap = a_prefix(parent);
  pugi::xml_node cur;
  for (pugi::xml_node child = parent.first_child(); child && !cur; child = child.next_sibling()) {
    cur = child.type() == pugi::node_element && is_fill(child) ? child : pugi::xml_node();
  }
  std::uint32_t current = 0;
  if (cur && (none ? local_name(cur) == "noFill"
                   : local_name(cur) == "solidFill" &&
                         parse_rgb(child_local(cur, "srgbClr").attribute("val").value(), current) && current == rgb)) {
    return;
  }
  const std::string name = ap + (none ? "noFill" : "solidFill");
  pugi::xml_node fresh =
      cur ? parent.insert_child_before(name.c_str(), cur)
          : insert_after_or_first(parent, is_line ? pugi::xml_node() : last_leading_child(parent, false), name);
  if (!none) {
    char hex[8];
    format_hex(hex, sizeof hex, rgb & 0xFFFFFFU, 6, true);
    fresh.append_child((ap + "srgbClr").c_str()).append_attribute("val").set_value(hex);
  }
  if (cur) {
    parent.remove_child(cur);
  }
}

pugi::xml_node ensure_ln(pugi::xml_node props) {
  if (pugi::xml_node ln = child_local(props, "ln")) {
    return ln;
  }
  return insert_after_or_first(props, last_leading_child(props, true), a_prefix(props) + "ln");
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

bool marker_ok(const AnchorPoint& p) {
  return in_range(p.row_off) && in_range(p.col_off) && p.row < Sheet::kMaxRows && p.col < Sheet::kMaxCols;
}

// Writes the anchor fields of `spec` into the existing `anchor` and the matching `a:xfrm`.
Expected<void, Error> place(const Workbook& wb, std::size_t sheet, pugi::xml_node anchor, pugi::xml_node content,
                            const ShapeSpec& spec) {
  const bool absolute = spec.anchor_kind == AnchorKind::kAbsolute;
  const bool two_cell = spec.anchor_kind == AnchorKind::kTwoCell;
  if (read_drawing_object(anchor, {}).anchor_kind != spec.anchor_kind) {
    return invalid("shape: the anchor kind cannot change");
  }
  if (absolute ? !in_range(spec.from.row_off) || !in_range(spec.from.col_off) : !marker_ok(spec.from)) {
    return invalid("shape: anchor cell or offset out of range");
  }
  if (two_cell ? !marker_ok(spec.to) || spec.to.row < spec.from.row || spec.to.col < spec.from.col
               : !in_range(spec.cx) || !in_range(spec.cy)) {
    return invalid("shape: anchor box or size out of range");
  }
  std::int64_t off_x = spec.from.col_off;
  std::int64_t off_y = spec.from.row_off;
  std::int64_t cx = spec.cx;
  std::int64_t cy = spec.cy;
  if (absolute) {
    pugi::xml_node pos = child_local(anchor, "pos");
    set_llong_attr(pos, "x", off_x);
    set_llong_attr(pos, "y", off_y);
  } else {
    write_anchor_point(child_local(anchor, "from"), spec.from);
    off_x += sheet_offset_emu(wb, sheet, false, spec.from.col);
    off_y += sheet_offset_emu(wb, sheet, true, spec.from.row);
  }
  if (two_cell) {
    write_anchor_point(child_local(anchor, "to"), spec.to);
    cx = sheet_offset_emu(wb, sheet, false, spec.to.col) + spec.to.col_off - off_x;
    cy = sheet_offset_emu(wb, sheet, true, spec.to.row) + spec.to.row_off - off_y;
    pugi::xml_attribute edit_as = anchor.attribute("editAs");
    if (spec.edit_as == EditAs::kTwoCell) {
      anchor.remove_attribute(edit_as);
    } else {
      (edit_as ? edit_as : anchor.append_attribute("editAs"))
          .set_value(spec.edit_as == EditAs::kOneCell ? "oneCell" : "absolute");
    }
  } else {
    pugi::xml_node ext = child_local(anchor, "ext");
    set_llong_attr(ext, "cx", cx);
    set_llong_attr(ext, "cy", cy);
  }
  pugi::xml_node xfrm = ensure_xfrm(props_of(content));
  const std::string ap = a_prefix(xfrm);
  pugi::xml_node off = ensure_child(xfrm, ap + "off", true);
  set_llong_attr(off, "x", off_x);
  set_llong_attr(off, "y", off_y);
  pugi::xml_node ext = ensure_child(xfrm, ap + "ext");
  set_llong_attr(ext, "cx", std::max<std::int64_t>(cx, 0));
  set_llong_attr(ext, "cy", std::max<std::int64_t>(cy, 0));
  return Expected<void, Error>::Ok();
}

void set_text(pugi::xml_node content, const std::string& text) {
  const std::string ap = a_prefix(content);
  pugi::xml_node body = child_local(content, "txBody");
  if (!body) {
    body = content.append_child((xdr_prefix(content) + "txBody").c_str());
    pugi::xml_node pr = body.append_child((ap + "bodyPr").c_str());
    pr.append_attribute("vertOverflow").set_value("clip");
    pr.append_attribute("wrap").set_value("square");
    pr.append_attribute("rtlCol").set_value("0");
    pr.append_attribute("anchor").set_value("t");
    body.append_child((ap + "lstStyle").c_str());
  }
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

bool is_token(const std::string& s) {
  return !s.empty() && std::all_of(s.begin(), s.end(), [](unsigned char c) { return std::isalnum(c) != 0; });
}

// Applies the non-anchor fields of `spec` to `content`; `anchor` names its `cNvPr`.
Expected<void, Error> apply_fields(pugi::xml_node anchor, pugi::xml_node content, const ShapeSpec& spec) {
  const std::uint32_t f = spec.fields;
  const std::string_view kind = local_name(content);
  const bool group = kind == "grpSp";
  pugi::xml_node props = props_of(content);
  if (!props) {
    return invalid("shape: no shape properties");
  }
  const auto paint_ok = [](ShapePaint p) { return p == ShapePaint::kSolid || p == ShapePaint::kNone; };
  if (((f & shape_field::kFill) != 0U && !paint_ok(spec.fill)) ||
      ((f & shape_field::kLine) != 0U && !paint_ok(spec.line))) {
    return invalid("shape: a fill or outline is solid or none");
  }
  if ((f & shape_field::kPrst) != 0U && (group || !is_token(spec.prst))) {
    return invalid("shape: preset geometry must be an alphanumeric token on a shape or connector");
  }
  if ((f & (shape_field::kText | shape_field::kTextBox)) != 0U && kind != "sp") {
    return invalid("shape: only a shape holds text");
  }
  if ((f & shape_field::kText) != 0U && std::any_of(spec.text.begin(), spec.text.end(), [](unsigned char c) {
        return c < 0x20U && c != '\n' && c != '\t';
      })) {
    return invalid("shape: text holds a control character");
  }
  pugi::xml_node cnvpr = anchor_cnvpr(anchor);
  if ((f & shape_field::kName) != 0U) {
    (cnvpr.attribute("name") ? cnvpr.attribute("name") : cnvpr.append_attribute("name")).set_value(spec.name.c_str());
  }
  if ((f & shape_field::kHidden) != 0U) {
    if (spec.hidden) {
      (cnvpr.attribute("hidden") ? cnvpr.attribute("hidden") : cnvpr.append_attribute("hidden")).set_value("1");
    } else {
      cnvpr.remove_attribute("hidden");
    }
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
  if ((f & shape_field::kFill) != 0U) {
    set_paint(props, false, spec.fill == ShapePaint::kNone, spec.fill_rgb);
  }
  if ((f & shape_field::kLine) != 0U) {
    set_paint(ensure_ln(props), true, spec.line == ShapePaint::kNone, spec.line_rgb);
  }
  if ((f & shape_field::kLineWidth) != 0U) {
    if (spec.line_width_emu < 0 || spec.line_width_emu >= kEmuLimit) {
      return invalid("shape: line width out of range");
    }
    set_llong_attr(ensure_ln(props), "w", spec.line_width_emu);
  }
  if ((f & shape_field::kRotation) != 0U) {
    pugi::xml_node xfrm = ensure_xfrm(props);
    if (spec.rotation == 0) {
      xfrm.remove_attribute("rot");
    } else {
      set_llong_attr(xfrm, "rot", spec.rotation);
    }
  }
  if ((f & shape_field::kText) != 0U) {
    set_text(content, spec.text);
  }
  return Expected<void, Error>::Ok();
}

bool is_connector_prst(const std::string& prst) {
  return prst == "line" || prst == "straightConnector1" || prst.rfind("bentConnector", 0) == 0 ||
         prst.rfind("curvedConnector", 0) == 0;
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
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
  for (pugi::xml_node anchor : drawing_anchors(doc.document_element(), false)) {
    DrawingObject obj = read_drawing_object(anchor, {});
    if (is_shape_kind(obj.kind)) {
      shapes.push_back(read_shape(anchor, std::move(obj)));
    }
  }
  return shapes;
}

Expected<std::uint32_t, Error> insert_shape(Workbook& wb, std::size_t sheet, const ShapeSpec& spec) {
  ASSIGN_OR_RETURN(std::string path, editable_path(wb, sheet));
  if ((spec.fields & shape_field::kAnchor) == 0U) {
    return invalid("shape: an anchor is required");
  }
  const std::string prst = (spec.fields & shape_field::kPrst) != 0U ? spec.prst : "rect";
  if (!is_token(prst)) {
    return invalid("shape: preset geometry must be an alphanumeric token");
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
    RETURN_IF_ERROR(load_drawing(wb, path, doc));
  }
  pugi::xml_node root = doc.document_element();
  const std::string xp = xdr_prefix(root);
  const std::string ap = a_prefix(root);
  std::uint32_t id = 1;
  doc.find_node([&id](pugi::xml_node n) {
    if (n.type() == pugi::node_element && local_name(n) == "cNvPr") {
      id = std::max(id, n.attribute("id").as_uint());
    }
    return false;
  });
  if (id == std::numeric_limits<std::uint32_t>::max()) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "drawing: object id space exhausted");
  }
  ++id;

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

Expected<void, Error> update_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id, const ShapeSpec& spec) {
  ASSIGN_OR_RETURN(std::string path, editable_path(wb, sheet));
  pugi::xml_document doc;
  if (path.empty()) {
    return invalid("shape: no shape with that id");
  }
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
  pugi::xml_node anchor = find_shape(doc.document_element(), object_id);
  if (!anchor) {
    return invalid("shape: no shape with that id");
  }
  pugi::xml_node content = anchor_content(anchor);
  if ((spec.fields & shape_field::kAnchor) != 0U) {
    RETURN_IF_ERROR(place(wb, sheet, anchor, content, spec));
  }
  ShapeSpec rest = spec;
  rest.fields &= ~shape_field::kAnchor;
  RETURN_IF_ERROR(apply_fields(anchor, content, rest));
  return io::ooxml::store_part_dom(wb, path, doc);
}

Expected<void, Error> remove_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id) {
  ASSIGN_OR_RETURN(std::string path, editable_path(wb, sheet));
  pugi::xml_document doc;
  if (path.empty()) {
    return invalid("shape: no shape with that id");
  }
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
  pugi::xml_node root = doc.document_element();
  pugi::xml_node top = find_shape(root, object_id);
  if (!top) {
    return invalid("shape: no shape with that id");
  }
  while (top.parent() != root) {
    top = top.parent();
  }
  root.remove_child(top);
  return io::ooxml::store_part_dom(wb, path, doc);
}

}  // namespace formulon
