#include "drawing/drawing_edit.h"

#include <algorithm>
#include <limits>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "chart/chart_ns.h"
#include "drawing/connector_route.h"
#include "drawing/drawing_edit_internal.h"
#include "drawing/drawing_geometry.h"
#include "drawing/image_header.h"
#include "drawing/shape_edit.h"
#include "drawing/slicer_edit.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml/relationship_xml.h"
#include "io/ooxml_defs.h"
#include "io/worksheet_child_refs.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "new_workbook_defaults.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {
namespace drawing_edit {

namespace {

constexpr const char* kRelImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

constexpr const char* kEmptyRels =
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>";

constexpr std::int64_t kEmuPerPixel = 9525;

}  // namespace

Error refuse_object(const pugi::xml_node& root, std::uint32_t object_id, const char* message) {
  const auto any_kind = [](DrawingObjectKind) { return true; };
  return find_drawing_object(root, object_id, +any_kind) ? invalid(message) : not_found(message);
}

Expected<std::string, Error> load_object_drawing(const Workbook& wb, std::size_t sheet, pugi::xml_document& doc) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if (path.empty()) {
    return not_found(kNoObject);
  }
  RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
  return path;
}

Expected<std::uint32_t, Error> allocate_drawing_id(Sheet& sheet, const pugi::xml_document& doc) {
  std::vector<std::uint32_t> used;
  for_each_element(doc, [&used](const pugi::xml_node& node) {
    if (local_name(node) == "cNvPr") {
      used.push_back(node.attribute("id").as_uint());
    }
  });
  std::uint32_t id = std::max(sheet.next_drawing_id(), 2U);
  while (id != std::numeric_limits<std::uint32_t>::max() && std::find(used.begin(), used.end(), id) != used.end()) {
    ++id;
  }
  if (id == std::numeric_limits<std::uint32_t>::max()) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "The drawing's object id space is exhausted");
  }
  sheet.set_next_drawing_id(id + 1U);
  return id;
}

namespace {

const char* kind_word(const NewWorkbookDefaults& row, DrawingObjectKind kind) {
  return kind == DrawingObjectKind::kChart     ? row.chart_word
         : kind == DrawingObjectKind::kPicture ? row.picture_word
                                               : nullptr;
}

}  // namespace

std::string default_drawing_name(const Workbook& wb, DrawingObjectKind kind, std::uint32_t id) {
  const char* word = kind_word(new_workbook_defaults(wb.excel_profile().locale), kind);
  return word == nullptr ? std::string() : std::string(word) + " " + std::to_string(id - 1U);
}

bool is_default_drawing_name(const Workbook& wb, DrawingObjectKind kind, std::string_view name) {
  const auto matches = [&](const char* word) {
    const std::string_view w = word == nullptr ? std::string_view() : std::string_view(word);
    if (w.empty() || name.size() <= w.size() + 1U || !strings::starts_with(name, w) || name[w.size()] != ' ') {
      return false;
    }
    const std::string_view digits = name.substr(w.size() + 1U);
    return std::all_of(digits.begin(), digits.end(), strings::is_ascii_digit);
  };
  return matches(kind_word(new_workbook_defaults(wb.excel_profile().locale), kind)) ||
         matches(kind_word(new_workbook_defaults(ExcelLocale::kEnUS), kind));
}

std::vector<std::string> relationship_refs(const pugi::xml_node& node) {
  std::vector<std::string> used;
  const auto collect = [&used](const pugi::xml_node& element) {
    for (pugi::xml_attribute attr = element.first_attribute(); attr; attr = attr.next_attribute()) {
      const std::string_view name = attr.name();
      if (name.find(':') != std::string_view::npos && !strings::starts_with(name, "xmlns:")) {
        used.emplace_back(attr.value());
      }
    }
  };
  collect(node);
  for_each_element(node, collect);
  return used;
}

std::string unused_part_path(const Workbook& wb, const std::string& stem, const char* suffix,
                             const std::vector<std::string>& reserved) {
  for (std::uint32_t n = 1;; ++n) {
    const std::string base = stem + std::to_string(n) + ".";
    const bool taken =
        std::any_of(wb.passthrough_parts().begin(), wb.passthrough_parts().end(),
                    [&](const PassthroughPart& part) { return strings::starts_with(part.path, base); }) ||
        std::any_of(reserved.begin(), reserved.end(),
                    [&](const std::string& path) { return strings::starts_with(path, base); });
    if (!taken) {
      return base + suffix;
    }
  }
}

Expected<void, Error> check_package_growth(const Workbook& wb, std::size_t extra_parts, std::size_t extra_bytes) {
  std::size_t total_bytes = extra_bytes;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    total_bytes += part.bytes.size();
  }
  return check_package_limits(wb.passthrough_parts().size() + wb.sheet_count() + extra_parts + 2U, total_bytes);
}

Expected<void, Error> prepare_drawing(const Workbook& wb, std::string& path, pugi::xml_document& doc) {
  if (path.empty()) {
    path = unused_part_path(wb, "xl/drawings/drawing", "xml");
    doc.append_child("xdr:wsDr");
    return Expected<void, Error>::Ok();
  }
  return load_drawing_part(wb, path, doc);
}

namespace {

// Removes the relationships of `rels_doc` no element of the drawing names (`used`), and the parts they
// leave unreferenced.
Expected<void, Error> drop_unused_rels(Workbook& wb, const std::string& rels_path, pugi::xml_document& rels_doc,
                                       const std::vector<io::ooxml::DrawingRel>& rels,
                                       const std::vector<std::string>& used) {
  std::vector<std::string> candidates;
  pugi::xml_node root = io::ooxml::relationship_root(rels_doc);
  if (!root) {
    return unparseable("Drawing rels part has an invalid Relationships root", rels_path);
  }
  for (pugi::xml_node rel = root.first_child(); rel;) {
    pugi::xml_node next = rel.next_sibling();
    if (!io::ooxml::is_relationship_element(rel, root)) {
      rel = next;
      continue;
    }
    const std::string id = rel.attribute("Id").value();
    if (std::find(used.begin(), used.end(), id) == used.end()) {
      for (const io::ooxml::DrawingRel& entry : rels) {
        if (entry.id == id && !entry.external) {
          candidates.push_back(entry.target);
        }
      }
      root.remove_child(rel);
    }
    rel = next;
  }
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, rels_path, rels_doc));
  io::ooxml::remove_orphans(wb, std::move(candidates));
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<void, Error> drop_objects(Workbook& wb, std::size_t sheet, const std::string& path, pugi::xml_document& doc,
                                   const std::vector<pugi::xml_node>& tops) {
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, path));
  std::vector<std::uint32_t> gone;
  for (const pugi::xml_node& top : tops) {
    for_each_element(top, [&gone](const pugi::xml_node& node) {
      if (local_name(node) == "cNvPr") {
        gone.push_back(node.attribute("id").as_uint());
      }
    });
    top.parent().remove_child(top);
  }
  // A connector glued to a removed object loses that end, as in Excel; an id still present (a restored copy) keeps it.
  for_each_element(doc.document_element(), [&gone](const pugi::xml_node& node) {
    if (local_name(node) == "cNvPr") {
      gone.erase(std::remove(gone.begin(), gone.end(), node.attribute("id").as_uint()), gone.end());
    }
  });
  std::vector<pugi::xml_node> dangling;
  for_each_element(doc.document_element(), [&gone, &dangling](const pugi::xml_node& node) {
    const std::string_view name = local_name(node);
    if ((name == "stCxn" || name == "endCxn") &&
        std::find(gone.begin(), gone.end(), node.attribute("id").as_uint()) != gone.end()) {
      dangling.push_back(node);
    }
  });
  for (const pugi::xml_node& node : dangling) {
    node.parent().remove_child(node);
  }
  const std::vector<std::string> used = relationship_refs(doc);
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  const std::string rels_path = io::ooxml::rels_path_for_part(path);
  pugi::xml_document rels_doc;
  if (!rels.empty() && io::ooxml::load_part_dom(wb, rels_path, "drawing", rels_doc)) {
    RETURN_IF_ERROR(drop_unused_rels(wb, rels_path, rels_doc, rels, used));
  }
  if (doc.document_element() && !io::has_element_child(doc.document_element())) {
    std::string part = path;  // `path` may be the sheet's own link, cleared next.
    wb.sheet(sheet).set_drawing_rel_target({});
    io::ooxml::remove_orphans(wb, {std::move(part)});
  }
  return Expected<void, Error>::Ok();
}

namespace {

std::string marker_xml(const char* tag, const AnchorPoint& p) {
  return std::string("<xdr:") + tag + "><xdr:col>" + std::to_string(p.col) + "</xdr:col><xdr:colOff>" +
         std::to_string(p.col_off) + "</xdr:colOff><xdr:row>" + std::to_string(p.row) + "</xdr:row><xdr:rowOff>" +
         std::to_string(p.row_off) + "</xdr:rowOff></xdr:" + tag + ">";
}

}  // namespace

Expected<void, Error> add_empty_part(Workbook& wb, const std::string& path, const char* content_type) {
  if (io::ooxml::find_passthrough_part(wb, path) != nullptr) {
    return Expected<void, Error>::Ok();
  }
  return wb.add_passthrough_part(PassthroughPart(path, content_type, {}));
}

void ensure_default_content_type(Workbook& wb, const char* extension, const char* content_type) {
  for (const DefaultContentType& def : wb.default_content_types()) {
    if (def.extension == extension) {
      return;
    }
  }
  std::vector<DefaultContentType> defaults = wb.default_content_types();
  defaults.push_back({extension, content_type});
  wb.set_default_content_types(std::move(defaults));
}

namespace {

bool cell_in_range(std::uint32_t row, std::uint32_t col) {
  return row < Sheet::kMaxRows && col < Sheet::kMaxCols;
}

// Whether a marker's offsets, the object's size and its cell all lie inside the sheet.
bool box_in_range(std::int64_t row_off, std::int64_t col_off, std::int64_t cx, std::int64_t cy, std::uint32_t row,
                  std::uint32_t col) {
  return in_emu_range(row_off) && in_emu_range(col_off) && in_emu_range(cx) && in_emu_range(cy) &&
         cell_in_range(row, col);
}

// The `edit`-axis line and offset of a marker.
io::AnchorLine axis_line(const AnchorPoint& p, bool row_axis) {
  return row_axis ? io::AnchorLine{p.row, p.row_off} : io::AnchorLine{p.col, p.col_off};
}

AnchorPoint with_axis_line(AnchorPoint p, bool row_axis, const io::AnchorLine& at) {
  (row_axis ? p.row : p.col) = at.line;
  (row_axis ? p.row_off : p.col_off) = at.offset;
  return p;
}

bool same_point(const AnchorPoint& a, const AnchorPoint& b) {
  return a.row == b.row && a.col == b.col && a.row_off == b.row_off && a.col_off == b.col_off;
}

}  // namespace

bool anchor_point_in_range(const AnchorPoint& p) {
  return in_emu_range(p.row_off) && in_emu_range(p.col_off) && cell_in_range(p.row, p.col);
}

pugi::xml_node top_of(const pugi::xml_node& root, pugi::xml_node node) {
  while (node.parent() != root) {
    node = node.parent();
  }
  return node;
}

pugi::xml_node insertion_point(const pugi::xml_node& root, const pugi::xml_node& moving, std::size_t index) {
  std::size_t n = 0;
  for (const pugi::xml_node& item : drawing_anchors(root, false)) {
    const pugi::xml_node top = top_of(root, item);
    if (top != moving && n++ == index) {
      return top;
    }
  }
  return {};
}

namespace {

std::string prefix_of(const pugi::xml_node& node) {
  return std::string(io::qname_prefix_colon(node.name()));
}

// Writes marker `tag`, creating it after `after` (or first) when absent.
pugi::xml_node put_marker(pugi::xml_node anchor, const pugi::xml_node& after, const char* tag, const AnchorPoint& p) {
  pugi::xml_node marker = child_local(anchor, tag);
  if (!marker) {
    const std::string pre = prefix_of(anchor);
    marker = after ? anchor.insert_child_after((pre + tag).c_str(), after) : anchor.prepend_child((pre + tag).c_str());
    for (const char* part : {"col", "colOff", "row", "rowOff"}) {
      marker.append_child((pre + part).c_str());
    }
  }
  write_anchor_point(marker, p);
  return marker;
}

}  // namespace

pugi::xml_node object_xfrm(const pugi::xml_node& anchor) {
  const pugi::xml_node content = anchor_content(anchor);
  const pugi::xml_node own = child_local(content, "xfrm");
  return own ? own : child_local(child_local(content, "spPr"), "xfrm");
}

bool swaps_box(const pugi::xml_node& xfrm) {
  std::int64_t deg = xfrm.attribute("rot").as_llong() / 60000 % 360;
  deg = deg < 0 ? deg + 360 : deg;
  return (deg >= 45 && deg < 135) || (deg >= 225 && deg < 315);
}

void place_anchor(pugi::xml_node anchor, const Placement& p) {
  const DrawingObject current = read_drawing_object(anchor, {});
  const bool two = p.kind == AnchorKind::kTwoCell;
  if (current.anchor_kind != p.kind) {
    anchor.set_name((prefix_of(anchor) + (two ? "twoCellAnchor" : "oneCellAnchor")).c_str());
    anchor.remove_attribute("editAs");
    for (const char* name : {"pos", "ext", "to"}) {
      anchor.remove_child(child_local(anchor, name));
    }
  }
  // An absent `editAs` means twoCell, so a converted anchor starts there.
  const EditAs edit_as = current.anchor_kind == AnchorKind::kTwoCell ? current.edit_as : EditAs::kTwoCell;
  if (two && p.edit_as != edit_as) {
    if (p.edit_as == EditAs::kTwoCell) {
      anchor.remove_attribute("editAs");
    } else {
      pugi::xml_attribute attr = anchor.attribute("editAs");
      (attr ? attr : anchor.append_attribute("editAs"))
          .set_value(p.edit_as == EditAs::kOneCell ? "oneCell" : "absolute");
    }
  }
  if (p.kind == AnchorKind::kAbsolute) {
    // Only an absolute anchor stays absolute: its `xdr:pos` and `xdr:ext` are already there.
    pugi::xml_node pos = child_local(anchor, "pos");
    io::set_attr(pos, "x", p.from.col_off);
    io::set_attr(pos, "y", p.from.row_off);
    pugi::xml_node ext = child_local(anchor, "ext");
    io::set_attr(ext, "cx", p.box_w);
    io::set_attr(ext, "cy", p.box_h);
  } else if (const pugi::xml_node from = put_marker(anchor, {}, "from", p.from); two) {
    put_marker(anchor, from, "to", p.to);
  } else {
    pugi::xml_node ext = child_local(anchor, "ext");
    if (!ext) {
      ext = anchor.insert_child_after((prefix_of(anchor) + "ext").c_str(), from);
    }
    io::set_attr(ext, "cx", p.box_w);
    io::set_attr(ext, "cy", p.box_h);
  }
  const pugi::xml_node xfrm = object_xfrm(anchor);
  if (const pugi::xml_node off = child_local(xfrm, "off")) {
    io::set_attr(off, "x", p.off_x);
    io::set_attr(off, "y", p.off_y);
  }
  if (const pugi::xml_node ext = child_local(xfrm, "ext")) {
    io::set_attr(ext, "cx", p.cx);
    io::set_attr(ext, "cy", p.cy);
  }
}

std::string next_rel_id(const std::vector<io::ooxml::DrawingRel>& rels) {
  for (std::uint32_t n = 1;; ++n) {
    const std::string candidate = "rId" + std::to_string(n);
    if (std::none_of(rels.begin(), rels.end(),
                     [&candidate](const io::ooxml::DrawingRel& rel) { return rel.id == candidate; })) {
      return candidate;
    }
  }
}

Expected<void, Error> load_rels_doc(const Workbook& wb, const std::string& rels_path, pugi::xml_document& rels_doc) {
  if (io::ooxml::find_passthrough_part(wb, rels_path) != nullptr) {
    return io::ooxml::load_part_dom(wb, rels_path, "drawing", rels_doc);
  }
  rels_doc.load_string(kEmptyRels);
  return Expected<void, Error>::Ok();
}

void append_rel(pugi::xml_document& rels_doc, const std::string& id, std::string_view type, const std::string& target,
                bool external) {
  pugi::xml_node root = io::ooxml::relationship_root(rels_doc);
  pugi::xml_node node = root.append_child(io::ooxml::relationship_child_qname(root).c_str());
  node.append_attribute("Id").set_value(id.c_str());
  node.append_attribute("Type").set_value(std::string(type).c_str());
  node.append_attribute("Target").set_value(target.c_str());
  if (external) {
    node.append_attribute("TargetMode").set_value("External");
  }
}

namespace {

// Canonical markers and `a:off` of an inserted object: the offset lies inside
// its cell and `to` is `from` plus the object's size under the display geometry.
struct InsertBox {
  AnchorPoint from, to;
  std::int64_t off_x = 0, off_y = 0;
};

// Validates an insertion of a `cx` x `cy` EMU object adding `parts` parts of
// `len` bytes in total, and lays out its box.
Expected<InsertBox, Error> insert_box(const Workbook& wb, std::size_t sheet, const ImageInsertOptions& options,
                                      std::int64_t cx, std::int64_t cy, std::size_t parts, std::size_t len) {
  if (options.anchor_kind == AnchorKind::kAbsolute) {
    return invalid("anchor.kind must be a one-cell or two-cell anchor");
  }
  if (!box_in_range(options.row_off, options.col_off, cx, cy, options.row, options.col)) {
    return invalid("anchor offsets, size and cells must be in range");
  }
  RETURN_IF_ERROR(check_package_growth(wb, parts, len));
  const Geometry geo(wb, wb.sheet(sheet));
  InsertBox box;
  box.from = AnchorPoint{options.row, options.col, options.row_off, options.col_off};
  if (!geo.advance(true, box.from.row, box.from.row_off, 0) || !geo.advance(false, box.from.col, box.from.col_off, 0)) {
    return invalid("anchor must lie within the sheet");
  }
  box.to = box.from;
  if (!geo.advance(true, box.to.row, box.to.row_off, cy) || !geo.advance(false, box.to.col, box.to.col_off, cx)) {
    return invalid("anchor.width_emu and anchor.height_emu must keep the object within the sheet");
  }
  box.off_x = geo.offset(false, box.from.col) + box.from.col_off;
  box.off_y = geo.offset(true, box.from.row) + box.from.row_off;
  return box;
}

std::string anchor_open_xml(const ImageInsertOptions& options, const InsertBox& box, std::int64_t cx, std::int64_t cy) {
  if (options.anchor_kind == AnchorKind::kOneCell) {
    return "<xdr:oneCellAnchor>" + marker_xml("from", box.from) + "<xdr:ext cx=\"" + std::to_string(cx) + "\" cy=\"" +
           std::to_string(cy) + "\"/>";
  }
  std::string xml = "<xdr:twoCellAnchor";
  xml += options.edit_as == EditAs::kOneCell    ? " editAs=\"oneCell\""
         : options.edit_as == EditAs::kAbsolute ? " editAs=\"absolute\""
                                                : "";
  return xml + ">" + marker_xml("from", box.from) + marker_xml("to", box.to);
}

const char* anchor_close_xml(const ImageInsertOptions& options) {
  return options.anchor_kind == AnchorKind::kOneCell ? "<xdr:clientData/></xdr:oneCellAnchor>"
                                                     : "<xdr:clientData/></xdr:twoCellAnchor>";
}

std::string xfrm_xml(const char* tag, const InsertBox& box, std::int64_t cx, std::int64_t cy) {
  return std::string("<") + tag + "><a:off x=\"" + std::to_string(box.off_x) + "\" y=\"" + std::to_string(box.off_y) +
         "\"/><a:ext cx=\"" + std::to_string(cx) + "\" cy=\"" + std::to_string(cy) + "\"/></" + tag + ">";
}

// Appends the single anchor element `xml` to `doc` and names its object.
Expected<void, Error> append_anchor(pugi::xml_document& doc, const std::string& xml, const std::string& name,
                                    const std::string& descr) {
  pugi::xml_node root = doc.document_element();
  const pugi::xml_node last = root.last_child();
  if (!root.append_buffer(xml.data(), xml.size())) {
    return invalid("Anchor could not be built");
  }
  const pugi::xml_node added = last ? last.next_sibling() : root.first_child();
  if (!added || added != root.last_child() || !anchor_cnvpr(added)) {
    return invalid("Anchor could not be built");
  }
  io::ensure_attr(root, "xmlns:xdr", chart_ns::kNsXdr);
  io::ensure_attr(root, "xmlns:a", chart_ns::kNsA);
  pugi::xml_node cnvpr = anchor_cnvpr(added);
  cnvpr.append_attribute("name").set_value(name.c_str());
  if (!descr.empty()) {
    cnvpr.append_attribute("descr").set_value(descr.c_str());
  }
  return Expected<void, Error>::Ok();
}

// Writes the drawing `doc` with a new relationship `rid` of `rel_type` to
// `target`, its rels part, and the sheet's drawing relationship when `fresh`.
Expected<void, Error> commit_drawing(Workbook& wb, std::size_t sheet, const std::string& path, bool fresh,
                                     const pugi::xml_document& doc, const std::string& rid, std::string_view rel_type,
                                     const std::string& target) {
  const std::string rels_path = io::ooxml::rels_path_for_part(path);
  pugi::xml_document rels_doc;
  RETURN_IF_ERROR(load_rels_doc(wb, rels_path, rels_doc));
  append_rel(rels_doc, rid, rel_type, relative_target(io::ooxml::dir_of(path), target), false);
  RETURN_IF_ERROR(add_empty_part(wb, path, chart_ns::kCtDrawing));
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  RETURN_IF_ERROR(add_empty_part(wb, rels_path, ""));
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, rels_path, rels_doc));
  if (fresh) {
    wb.sheet(sheet).set_drawing_rel_target(path);
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

}  // namespace drawing_edit

using namespace drawing_edit;

Expected<std::vector<DrawingObject>, Error> list_drawing_objects(const Workbook& wb, std::size_t sheet) {
  ASSIGN_OR_RETURN(std::string path, sheet_drawing_path(wb, sheet));
  if (path.empty()) {
    return std::vector<DrawingObject>{};
  }
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, path));
  std::vector<DrawingObject> objects = read_drawing_objects(doc, rels);
  for (DrawingObject& obj : objects) {
    if (const PassthroughPart* media = io::ooxml::find_passthrough_part(wb, obj.media_path)) {
      if (auto info = probe_image(media->bytes.data(), media->bytes.size())) {
        obj.image_format = info.value().format;
      }
    }
  }
  return objects;
}

Expected<std::uint32_t, Error> insert_image(Workbook& wb, std::size_t sheet, const std::uint8_t* bytes, std::size_t len,
                                            const ImageInsertOptions& options) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if (len > kMaxImageBytes) {
    return make_error(FormulonErrorCode::kSecurityResourceLimit, "The image exceeds the size limit",
                      limit_context("image_bytes", len, kMaxImageBytes));
  }
  ASSIGN_OR_RETURN(ImageInfo info, probe_image(bytes, len));
  const std::int64_t cx = options.width_emu != 0 ? options.width_emu : info.px_width * kEmuPerPixel;
  const std::int64_t cy = options.height_emu != 0 ? options.height_emu : info.px_height * kEmuPerPixel;
  ASSIGN_OR_RETURN(InsertBox box, insert_box(wb, sheet, options, cx, cy, 1, len));

  const bool fresh = path.empty();
  pugi::xml_document doc;
  RETURN_IF_ERROR(prepare_drawing(wb, path, doc));
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, path));
  ASSIGN_OR_RETURN(const std::uint32_t id, allocate_drawing_id(wb.sheet(sheet), doc));
  const std::string rid = next_rel_id(rels);
  const std::string media_path = unused_part_path(wb, "xl/media/image", image_extension(info.format));
  const std::string xml =
      anchor_open_xml(options, box, cx, cy) + "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" + std::to_string(id) +
      "\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip "
      "xmlns:r=\"" +
      std::string(io::kNsOfficeRelationships) + "\" r:embed=\"" + rid +
      "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>" + xfrm_xml("a:xfrm", box, cx, cy) +
      "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>" + anchor_close_xml(options);
  RETURN_IF_ERROR(append_anchor(
      doc, xml, options.name.empty() ? default_drawing_name(wb, DrawingObjectKind::kPicture, id) : options.name,
      options.descr));
  RETURN_IF_ERROR(wb.add_passthrough_part(
      PassthroughPart(media_path, std::string(), std::vector<std::uint8_t>(bytes, bytes + len))));
  ensure_default_content_type(wb, image_extension(info.format), image_content_type(info.format));
  RETURN_IF_ERROR(commit_drawing(wb, sheet, path, fresh, doc, rid, kRelImage, media_path));
  return id;
}

Expected<std::uint32_t, Error> insert_graphic_frame(Workbook& wb, std::size_t sheet,
                                                    const ImageInsertOptions& placement,
                                                    std::string_view graphic_data_xml,
                                                    std::string_view target_part_path, std::string_view rel_type) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if (placement.width_emu == 0 || placement.height_emu == 0) {
    return invalid("anchor.width_emu and anchor.height_emu must be non-zero");
  }
  if (!io::ooxml::is_safe_part_name(target_part_path) || rel_type.empty()) {
    return invalid("Graphic frame target is not a part name");
  }
  const std::int64_t cx = placement.width_emu;
  const std::int64_t cy = placement.height_emu;
  ASSIGN_OR_RETURN(InsertBox box, insert_box(wb, sheet, placement, cx, cy, 0, 0));

  const bool fresh = path.empty();
  pugi::xml_document doc;
  RETURN_IF_ERROR(prepare_drawing(wb, path, doc));
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, path));
  ASSIGN_OR_RETURN(const std::uint32_t id, allocate_drawing_id(wb.sheet(sheet), doc));
  const std::string rid = next_rel_id(rels);
  std::string data(graphic_data_xml);
  for (std::size_t at = data.find("{rId}"); at != std::string::npos; at = data.find("{rId}", at + rid.size())) {
    data.replace(at, 5, rid);
  }
  const std::string xml = anchor_open_xml(placement, box, cx, cy) +
                          "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" + std::to_string(id) +
                          "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>" + xfrm_xml("xdr:xfrm", box, cx, cy) +
                          "<a:graphic>" + data + "</a:graphic></xdr:graphicFrame>" + anchor_close_xml(placement);
  RETURN_IF_ERROR(append_anchor(
      doc, xml, placement.name.empty() ? default_drawing_name(wb, DrawingObjectKind::kChart, id) : placement.name,
      placement.descr));
  RETURN_IF_ERROR(commit_drawing(wb, sheet, path, fresh, doc, rid, rel_type, std::string(target_part_path)));
  return id;
}

namespace {

bool is_slicer_object(const pugi::xml_node& anchor, DrawingObjectKind /*kind*/) {
  return is_slicer_frame(anchor);
}

std::optional<std::size_t> slicer_index_of(const std::vector<SlicerInfo>& slicers, std::uint32_t object_id) {
  for (std::size_t i = 0; i < slicers.size(); ++i) {
    if (slicers[i].object.object_id == object_id) {
      return i;
    }
  }
  return std::nullopt;
}

}  // namespace

Expected<void, Error> remove_drawing(Workbook& wb, std::size_t sheet, std::uint32_t object_id) {
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  const pugi::xml_node root = doc.document_element();
  if (const pugi::xml_node target = find_drawing_object(root, object_id, is_editable)) {
    return drop_objects(wb, sheet, path, doc, {top_of(root, target)});
  }
  // A slicer frame is checked before the shapes: its `mc:Fallback` holds a shape with the same id.
  if (find_drawing_object(root, object_id, is_slicer_object)) {
    ASSIGN_OR_RETURN(std::vector<SlicerInfo> slicers, list_slicers(wb, sheet));
    if (const std::optional<std::size_t> index = slicer_index_of(slicers, object_id)) {
      return remove_slicer(wb, sheet, *index);
    }
    return refuse_object(root, object_id, kNoObject);
  }
  if (find_drawing_object(root, object_id, is_shape_kind)) {
    return remove_shape(wb, sheet, object_id);
  }
  return refuse_object(root, object_id, kNoObject);
}

std::vector<std::pair<std::uint32_t, std::int64_t>> slicer_frame_extents(const Workbook& wb, std::size_t sheet,
                                                                         bool row_axis) {
  std::vector<std::pair<std::uint32_t, std::int64_t>> extents;
  const std::string path = wb.sheet(sheet).drawing_rel_target();
  pugi::xml_document doc;
  if (path.empty() || !load_drawing_part(wb, path, doc)) {
    return extents;
  }
  const Geometry geo(wb, wb.sheet(sheet));
  for (const pugi::xml_node& anchor : drawing_anchors(doc.document_element(), true)) {
    const DrawingObject obj = read_drawing_object(anchor, {});
    if (obj.anchor_kind == AnchorKind::kTwoCell && obj.edit_as == EditAs::kOneCell && is_slicer_frame(anchor)) {
      const auto at = [&](const AnchorPoint& p) {
        return geo.offset(row_axis, row_axis ? p.row : p.col) + (row_axis ? p.row_off : p.col_off);
      };
      extents.emplace_back(obj.object_id, at(obj.to) - at(obj.from));
    }
  }
  return extents;
}

std::vector<std::uint32_t> shift_drawing_anchors(
    Workbook& wb, std::size_t sheet, std::uint32_t index, std::uint32_t count, bool is_delete, bool row_axis,
    const std::vector<std::pair<std::uint32_t, std::int64_t>>& slicer_extents) {
  std::vector<std::uint32_t> removed_ids;
  const std::string path = wb.sheet(sheet).drawing_rel_target();
  pugi::xml_document doc;
  if (path.empty() || count == 0 || !load_drawing_part(wb, path, doc)) {
    return removed_ids;
  }
  const Geometry geo(wb, wb.sheet(sheet));
  const StructuralEdit edit{index, count, is_delete, row_axis};
  const pugi::xml_node root = doc.document_element();
  const auto line = [row_axis](const AnchorPoint& p) { return row_axis ? p.row : p.col; };
  std::vector<pugi::xml_node> removed;
  bool changed = false;
  for (const pugi::xml_node& anchor : drawing_anchors(root, true)) {
    const DrawingObject obj = read_drawing_object(anchor, {});
    if (obj.edit_as == EditAs::kAbsolute) {
      continue;
    }
    const pugi::xml_node from_node = child_local(anchor, "from");
    const pugi::xml_node to_node = child_local(anchor, "to");
    const bool two_cell = obj.anchor_kind == AnchorKind::kTwoCell;
    if (two_cell && line(obj.to) < index) {
      continue;  // The edit lies wholly after the object.
    }
    io::AnchorLine from_line = axis_line(obj.from, row_axis);
    io::AnchorLine to_line_at = axis_line(obj.to, row_axis);
    if (two_cell && obj.edit_as == EditAs::kTwoCell && io::cell_anchor_deleted(from_line, to_line_at, edit)) {
      const pugi::xml_node top = top_of(root, anchor);
      if (std::find(removed.begin(), removed.end(), top) == removed.end()) {
        removed.push_back(top);
      }
      removed_ids.push_back(obj.object_id);
      continue;
    }
    std::int64_t extent = row_axis ? obj.cy : obj.cx;
    for (const auto& [id, size] : slicer_extents) {
      if (id == obj.object_id && extent == 0 && is_slicer_frame(anchor)) {
        extent = size;
      }
    }
    // A move-only object with no stored size (a form control or OLE object
    // keeps its xfrm zero) keeps its line span.
    io::move_cell_anchor(from_line, to_line_at, !(two_cell && obj.edit_as == EditAs::kOneCell && extent == 0), edit);
    const AnchorPoint from = with_axis_line(obj.from, row_axis, from_line);
    AnchorPoint to = with_axis_line(obj.to, row_axis, to_line_at);
    if (two_cell && obj.edit_as == EditAs::kOneCell && extent > 0) {
      // A move-only object keeps its size: `to` is re-derived from `from`.
      std::uint32_t& to_line = row_axis ? to.row : to.col;
      std::int64_t& to_off = row_axis ? to.row_off : to.col_off;
      to_line = line(from);
      to_off = row_axis ? from.row_off : from.col_off;
      geo.advance(row_axis, to_line, to_off, extent);
    }
    if (!same_point(from, obj.from)) {
      write_anchor_point(from_node, from);
      changed = true;
    }
    if (two_cell && !same_point(to, obj.to)) {
      write_anchor_point(to_node, to);
      changed = true;
    }
  }
  if (!removed.empty()) {
    static_cast<void>(drop_objects(wb, sheet, path, doc, removed));
  } else if (changed) {
    static_cast<void>(io::ooxml::store_part_dom(wb, path, doc));
  }
  return removed_ids;
}

std::int64_t sheet_offset_emu(const Workbook& wb, std::size_t sheet, bool row_axis, std::uint32_t line) {
  return Geometry(wb, wb.sheet(sheet)).offset(row_axis, line);
}

Expected<void, Error> set_drawing_name(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                       std::string_view name) {
  if (name.empty()) {
    return invalid("name must not be empty");
  }
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  const pugi::xml_node target = find_drawing_object(doc.document_element(), object_id, is_nameable);
  if (!target) {
    return refuse_object(doc.document_element(), object_id, kNoShapeOrObject);
  }
  pugi::xml_node cnvpr = anchor_cnvpr(target);
  if (!cnvpr) {
    return invalid(kNoShapeOrObject);
  }
  if (std::string_view(cnvpr.attribute("name").value()) == name) {
    return Expected<void, Error>::Ok();
  }
  io::set_attr(cnvpr, "name", name);
  return io::ooxml::store_part_dom(wb, path, doc);
}

Expected<void, Error> set_drawing_anchor(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                         const DrawingAnchor& anchor) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  pugi::xml_document doc;
  if (path.empty()) {
    return not_found(kNoObject);
  }
  RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
  const pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_drawing_object(root, object_id, is_editable);
  if (!target) {
    return refuse_object(root, object_id, kNoObject);
  }
  const Geometry geo(wb, wb.sheet(sheet));
  const DrawingObject obj = read_drawing_object(target, {});
  const bool absolute = anchor.anchor_kind == AnchorKind::kAbsolute;
  if (absolute && obj.anchor_kind != AnchorKind::kAbsolute) {
    return invalid(kNotAbsolute);
  }
  if (absolute && (anchor.row != 0 || anchor.col != 0)) {
    return invalid(kAbsoluteCells);
  }
  std::int64_t cur_cx = obj.cx;
  std::int64_t cur_cy = obj.cy;
  if (obj.kind == DrawingObjectKind::kChart && obj.anchor_kind == AnchorKind::kTwoCell) {
    // A chart frame's size is its anchor box; Excel leaves its `xdr:xfrm` zero.
    cur_cx = geo.span(false, obj.from, obj.to);
    cur_cy = geo.span(true, obj.from, obj.to);
  }
  Placement p;
  p.kind = anchor.anchor_kind;
  p.edit_as = anchor.edit_as;
  p.cx = anchor.width_emu != 0 ? anchor.width_emu : cur_cx;
  p.cy = anchor.height_emu != 0 ? anchor.height_emu : cur_cy;
  if (!box_in_range(anchor.row_off, anchor.col_off, p.cx, p.cy, anchor.row, anchor.col)) {
    return invalid("anchor offsets, size and cells must be in range");
  }
  p.from = AnchorPoint{anchor.row, anchor.col, anchor.row_off, anchor.col_off};
  const bool two = p.kind == AnchorKind::kTwoCell;
  if (obj.anchor_kind == p.kind && (!two || obj.edit_as == p.edit_as) && same_point(obj.from, p.from) &&
      cur_cx == p.cx && cur_cy == p.cy) {
    return Expected<void, Error>::Ok();
  }

  if (!geo.advance(true, p.from.row, p.from.row_off, 0) || !geo.advance(false, p.from.col, p.from.col_off, 0)) {
    return invalid("anchor must lie within the sheet");
  }
  const bool swap = swaps_box(object_xfrm(target));
  p.box_w = swap ? p.cy : p.cx;
  p.box_h = swap ? p.cx : p.cy;
  p.to = p.from;
  // An absolute position is an offset from A1, so the same walk bounds it.
  if (!geo.advance(true, p.to.row, p.to.row_off, p.box_h) || !geo.advance(false, p.to.col, p.to.col_off, p.box_w)) {
    return invalid("anchor.width_emu and anchor.height_emu must keep the object within the sheet");
  }
  // `a:off` is the unrotated rectangle sharing the box's centre.
  p.off_x = geo.offset(false, p.from.col) + p.from.col_off + (p.box_w - p.cx) / 2;
  p.off_y = geo.offset(true, p.from.row) + p.from.row_off + (p.box_h - p.cy) / 2;
  const std::vector<GluedEnd> glued = glued_ends(root);
  const pugi::xml_node top = top_of(root, target);
  for (const pugi::xml_node& other : drawing_anchors(root, true)) {
    const DrawingObject candidate = read_drawing_object(other, {});
    if (top_of(root, other) == top && is_editable(candidate.kind) && candidate.object_id == object_id) {
      place_anchor(other, p);
    }
  }
  follow_glued_shapes(wb, sheet, root, glued);
  return io::ooxml::store_part_dom(wb, path, doc);
}

Expected<void, Error> set_drawing_z_order(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                          std::uint32_t index) {
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_drawing_object(root, object_id, is_orderable);
  if (!target) {
    return refuse_object(root, object_id, kNoOrderable);
  }
  const std::vector<pugi::xml_node> items = drawing_anchors(root, false);
  if (index >= items.size()) {
    return make_error(FormulonErrorCode::kIndexOutOfRange, "z_order_index must be less than the object count");
  }
  const pugi::xml_node top = top_of(root, target);
  std::size_t current = 0;
  while (top_of(root, items[current]) != top) {
    ++current;
  }
  const pugi::xml_node before = insertion_point(root, top, index);
  if (before == insertion_point(root, top, current)) {
    return Expected<void, Error>::Ok();
  }
  if (before) {
    root.insert_move_before(top, before);
  } else {
    root.append_move(top);
  }
  return io::ooxml::store_part_dom(wb, path, doc);
}

namespace {

constexpr const char* kNoDrawingObject = "object_id does not name a picture, chart, shape or slicer";

enum class DrawingRoute { kFrame, kShape, kSlicer };

bool any_anchor_field(const DrawingAnchorPatch& p) {
  return p.kind || p.edit_as || p.from_row || p.from_col || p.from_row_off || p.from_col_off || p.to_row || p.to_col ||
         p.to_row_off || p.to_col_off || p.width_emu || p.height_emu;
}

// `patch` merged into the anchor of `obj` (whose box is `box_cx` x `box_cy`).
// `to` is meaningful for a two-cell result, `cx`/`cy` for any other.
struct MergedAnchor {
  AnchorKind kind = AnchorKind::kTwoCell;
  EditAs edit_as = EditAs::kTwoCell;
  AnchorPoint from;
  AnchorPoint to;
  std::int64_t cx = 0;
  std::int64_t cy = 0;
};

Expected<MergedAnchor, Error> merge_anchor(const Geometry& geo, const DrawingObject& obj, std::int64_t box_cx,
                                           std::int64_t box_cy, const DrawingAnchorPatch& p) {
  MergedAnchor m;
  m.kind = p.kind.value_or(obj.anchor_kind);
  m.edit_as = p.edit_as.value_or(obj.edit_as);
  if (m.kind == AnchorKind::kAbsolute && obj.anchor_kind != AnchorKind::kAbsolute) {
    return invalid(kNotAbsolute);
  }
  const bool two = m.kind == AnchorKind::kTwoCell;
  const bool has_to = p.to_row || p.to_col || p.to_row_off || p.to_col_off;
  const bool has_size = p.width_emu || p.height_emu;
  // A read anchor carries both the end marker and the size: the field the resulting kind holds wins.
  if (two && has_size && !has_to) {
    return invalid("anchor.width_emu and anchor.height_emu do not apply to a two-cell anchor");
  }
  if (!two && has_to && !has_size) {
    return invalid("anchor.to_row and anchor.to_col apply only to a two-cell anchor");
  }
  m.from = obj.from;
  m.from.row = p.from_row.value_or(m.from.row);
  m.from.col = p.from_col.value_or(m.from.col);
  m.from.row_off = p.from_row_off.value_or(m.from.row_off);
  m.from.col_off = p.from_col_off.value_or(m.from.col_off);
  if (m.kind == AnchorKind::kAbsolute && (m.from.row != 0 || m.from.col != 0)) {
    return invalid(kAbsoluteCells);
  }
  if (!two) {
    m.cx = p.width_emu.value_or(box_cx);
    m.cy = p.height_emu.value_or(box_cy);
    return m;
  }
  if (obj.anchor_kind == AnchorKind::kTwoCell) {
    m.to = obj.to;
  } else {
    // A converted anchor ends where the current size reaches from the new start.
    m.to = m.from;
    if (m.from.row >= Sheet::kMaxRows || m.from.col >= Sheet::kMaxCols ||
        !geo.advance(true, m.to.row, m.to.row_off, box_cy) || !geo.advance(false, m.to.col, m.to.col_off, box_cx)) {
      return invalid("anchor must lie within the sheet");
    }
  }
  m.to.row = p.to_row.value_or(m.to.row);
  m.to.col = p.to_col.value_or(m.to.col);
  m.to.row_off = p.to_row_off.value_or(m.to.row_off);
  m.to.col_off = p.to_col_off.value_or(m.to.col_off);
  if (!cell_in_range(m.to.row, m.to.col)) {
    return invalid("anchor offsets, size and cells must be in range");
  }
  m.cx = geo.span(false, m.from, m.to);
  m.cy = geo.span(true, m.from, m.to);
  if (m.cx < 0 || m.cy < 0) {
    return invalid("anchor.to_row and anchor.to_col must not precede anchor.from_row and anchor.from_col");
  }
  return m;
}

// Routes a merged anchor through the edit of the object's kind.
Expected<void, Error> apply_merged_anchor(Workbook& wb, std::size_t sheet, std::uint32_t object_id, DrawingRoute route,
                                          const MergedAnchor& m, bool swapped) {
  switch (route) {
    case DrawingRoute::kFrame: {
      DrawingAnchor a;
      a.anchor_kind = m.kind;
      a.edit_as = m.edit_as;
      a.row = m.from.row;
      a.col = m.from.col;
      a.row_off = m.from.row_off;
      a.col_off = m.from.col_off;
      // `set_drawing_anchor` takes the unrotated size of a picture laid out in a swapped box.
      a.width_emu = swapped ? m.cy : m.cx;
      a.height_emu = swapped ? m.cx : m.cy;
      if (a.width_emu == 0 || a.height_emu == 0) {
        return invalid("anchor.width_emu and anchor.height_emu must be non-zero");
      }
      return set_drawing_anchor(wb, sheet, object_id, a);
    }
    case DrawingRoute::kShape: {
      ShapeSpec spec;
      spec.fields = shape_field::kAnchor;
      spec.anchor_kind = m.kind;
      spec.edit_as = m.edit_as;
      spec.from = m.from;
      spec.to = m.to;
      spec.cx = m.cx;
      spec.cy = m.cy;
      return update_shape(wb, sheet, object_id, spec);
    }
    case DrawingRoute::kSlicer: {
      if (m.kind != AnchorKind::kTwoCell) {
        return invalid("anchor.kind must be a two-cell anchor for a slicer");
      }
      ASSIGN_OR_RETURN(std::vector<SlicerInfo> slicers, list_slicers(wb, sheet));
      if (const std::optional<std::size_t> index = slicer_index_of(slicers, object_id)) {
        SlicerSpec spec;
        spec.fields = slicer_field::kAnchor;
        spec.from = m.from;
        spec.to = m.to;
        return update_slicer(wb, sheet, *index, spec);
      }
      return not_found(kNoDrawingObject);
    }
  }
  return invalid(kNoDrawingObject);
}

// Every `cNvPr` of `top` carrying `object_id`: both branches of an `mc:AlternateContent`.
std::vector<pugi::xml_node> object_cnvprs(const pugi::xml_node& top, std::uint32_t object_id) {
  std::vector<pugi::xml_node> out;
  for_each_element(top, [&](const pugi::xml_node& node) {
    if (local_name(node) == "cNvPr" && node.attribute("id").as_uint() == object_id) {
      out.push_back(node);
    }
  });
  return out;
}

// Applies the `cNvPr` fields of `patch`; true when anything changed.
bool apply_cnvpr_fields(pugi::xml_node cnvpr, const DrawingCommonPatch& patch) {
  bool changed = false;
  if (patch.name && std::string_view(cnvpr.attribute("name").value()) != *patch.name) {
    io::set_attr(cnvpr, "name", *patch.name);
    changed = true;
  }
  if (patch.description && std::string_view(cnvpr.attribute("descr").value()) != *patch.description) {
    if (patch.description->empty()) {
      cnvpr.remove_attribute("descr");
    } else {
      io::set_attr(cnvpr, "descr", *patch.description);
    }
    changed = true;
  }
  if (patch.hidden && cnvpr.attribute("hidden").as_bool() != *patch.hidden) {
    set_cnvpr_hidden(cnvpr, *patch.hidden);
    changed = true;
  }
  return changed;
}

}  // namespace

Expected<void, Error> update_drawing(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                     const DrawingCommonPatch& patch) {
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  pugi::xml_node root = doc.document_element();
  DrawingRoute route = DrawingRoute::kFrame;
  // A slicer frame is checked before the shapes: its `mc:Fallback` holds a shape with the same id.
  pugi::xml_node target = find_drawing_object(root, object_id, is_slicer_object);
  if (target) {
    route = DrawingRoute::kSlicer;
  } else if ((target = find_drawing_object(root, object_id, is_editable))) {
    route = DrawingRoute::kFrame;
  } else if ((target = find_drawing_object(root, object_id, is_shape_kind))) {
    route = DrawingRoute::kShape;
  } else {
    return refuse_object(root, object_id, kNoDrawingObject);
  }
  if (patch.name && route == DrawingRoute::kSlicer) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "name cannot be changed on a slicer");
  }
  if (patch.name && patch.name->empty()) {
    return invalid("name must not be empty");
  }

  std::optional<MergedAnchor> merged;
  bool swapped = false;
  if (any_anchor_field(patch.anchor)) {
    const Geometry geo(wb, wb.sheet(sheet));
    const DrawingObject obj = read_drawing_object(target, {});
    swapped =
        route == DrawingRoute::kFrame && obj.kind == DrawingObjectKind::kPicture && swaps_box(object_xfrm(target));
    std::int64_t box_cx = swapped ? obj.cy : obj.cx;
    std::int64_t box_cy = swapped ? obj.cx : obj.cy;
    if (obj.anchor_kind == AnchorKind::kTwoCell) {
      box_cx = geo.span(false, obj.from, obj.to);
      box_cy = geo.span(true, obj.from, obj.to);
    }
    ASSIGN_OR_RETURN(MergedAnchor m, merge_anchor(geo, obj, box_cx, box_cy, patch.anchor));
    merged = m;
  }

  // The anchor edit stores the part itself; keep the bytes so a later failure leaves the drawing as it was.
  std::vector<std::uint8_t> before;
  if (const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path)) {
    before = part->bytes;
  }
  if (merged) {
    RETURN_IF_ERROR(apply_merged_anchor(wb, sheet, object_id, route, *merged, swapped));
    RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
    root = doc.document_element();
    const auto any_kind = [](DrawingObjectKind) { return true; };
    target = find_drawing_object(root, object_id, +any_kind);
    if (!target) {
      (void)wb.replace_passthrough_part(path, std::move(before));
      return not_found(kNoDrawingObject);
    }
  }
  bool changed = false;
  for (pugi::xml_node cnvpr : object_cnvprs(top_of(root, target), object_id)) {
    changed = apply_cnvpr_fields(cnvpr, patch) || changed;
  }
  if (!changed) {
    return Expected<void, Error>::Ok();
  }
  if (auto stored = io::ooxml::store_part_dom(wb, path, doc); !stored) {
    if (merged) {
      (void)wb.replace_passthrough_part(path, std::move(before));
    }
    return stored;
  }
  return Expected<void, Error>::Ok();
}

namespace {

bool is_picture_kind(DrawingObjectKind kind) {
  return kind == DrawingObjectKind::kPicture;
}

constexpr const char* kNoPicture = "object_id does not name a picture";

}  // namespace

Expected<std::optional<PictureTransform>, Error> get_picture_transform(const Workbook& wb, std::size_t sheet,
                                                                       std::uint32_t object_id) {
  ASSIGN_OR_RETURN(std::string path, sheet_drawing_path(wb, sheet));
  if (path.empty()) {
    return std::optional<PictureTransform>{};
  }
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
  const pugi::xml_node target = find_drawing_object(doc.document_element(), object_id, is_picture_kind);
  if (!target) {
    return std::optional<PictureTransform>{};
  }
  const XfrmTransform t = read_xfrm_transform(object_xfrm(target));
  return std::optional<PictureTransform>{PictureTransform{t.rot, t.flip_h, t.flip_v}};
}

Expected<void, Error> update_picture(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                     const DrawingCommonPatch* common, const PictureTransformPatch* picture) {
  const bool has_picture = picture != nullptr && (picture->rotation || picture->flip_h || picture->flip_v);
  if (!has_picture) {
    return common != nullptr ? update_drawing(wb, sheet, object_id, *common) : Expected<void, Error>::Ok();
  }
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  const pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_drawing_object(root, object_id, is_picture_kind);
  if (!target) {
    return refuse_object(root, object_id, kNoPicture);
  }
  pugi::xml_node props = child_local(anchor_content(target), "spPr");
  if (!props) {
    return invalid("The picture has no shape properties");
  }
  pugi::xml_node xfrm = child_local(props, "xfrm");
  const XfrmTransform current = read_xfrm_transform(xfrm);
  XfrmTransform wanted = current;
  wanted.rot = picture->rotation.value_or(current.rot);
  wanted.flip_h = picture->flip_h.value_or(current.flip_h);
  wanted.flip_v = picture->flip_v.value_or(current.flip_v);
  std::vector<std::uint8_t> before;
  const bool changed = wanted != current;
  if (changed) {
    if (const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path)) {
      before = part->bytes;
    }
    if (!xfrm) {
      xfrm = props.prepend_child(
          (prefix_of(child_local(child_local(anchor_content(target), "blipFill"), "blip")) + "xfrm").c_str());
    }
    const std::vector<GluedEnd> glued = glued_ends(root);
    write_xfrm_transform(xfrm, wanted);
    follow_glued_shapes(wb, sheet, root, glued);
    RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  }
  if (common == nullptr) {
    return Expected<void, Error>::Ok();
  }
  auto applied = update_drawing(wb, sheet, object_id, *common);
  if (!applied && changed) {
    (void)wb.replace_passthrough_part(path, std::move(before));
  }
  return applied;
}

}  // namespace formulon
