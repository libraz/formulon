#include "drawing/drawing_edit.h"

#include <algorithm>
#include <limits>
#include <string_view>
#include <utility>
#include <vector>

#include "chart/chart_ns.h"
#include "drawing/drawing_edit_internal.h"
#include "drawing/drawing_geometry.h"
#include "drawing/image_header.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml/relationship_xml.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "new_workbook_defaults.h"
#include "passthrough_part.h"
#include "sheet.h"
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

constexpr std::int64_t kEmuLimit = std::int64_t{1} << 31;

}  // namespace

Expected<std::string, Error> load_object_drawing(const Workbook& wb, std::size_t sheet, pugi::xml_document& doc) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if (path.empty()) {
    return invalid(kNoObject);
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
    return make_error(FormulonErrorCode::kPreconditionFailed, "drawing: object id space exhausted");
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

Expected<void, Error> prepare_drawing(const Workbook& wb, std::string& path, pugi::xml_document& doc) {
  if (path.empty()) {
    path = unused_part_path(wb, "xl/drawings/drawing", "xml");
    doc.append_child("xdr:wsDr");
    return Expected<void, Error>::Ok();
  }
  return load_drawing_part(wb, path, doc);
}

Expected<void, Error> drop_objects(Workbook& wb, const std::string& path, pugi::xml_document& doc,
                                   const std::vector<pugi::xml_node>& tops) {
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, path));
  for (const pugi::xml_node& top : tops) {
    top.parent().remove_child(top);
  }
  const std::vector<std::string> used = relationship_refs(doc);
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  const std::string rels_path = io::ooxml::rels_path_for_part(path);
  pugi::xml_document rels_doc;
  if (rels.empty() || !io::ooxml::load_part_dom(wb, rels_path, "drawing", rels_doc)) {
    return Expected<void, Error>::Ok();
  }
  std::vector<std::string> candidates;
  pugi::xml_node root = io::ooxml::relationship_root(rels_doc);
  if (!root) {
    return unparseable("drawing: rels part has an invalid Relationships root", rels_path);
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

bool in_emu_range(std::int64_t value) {
  return value >= 0 && value < kEmuLimit;
}

// Moves one marker for an edit along its axis; a marker inside a deleted span
// lands on the deletion boundary with offset 0.
AnchorPoint shift_point(AnchorPoint p, std::uint32_t index, std::uint32_t count, bool is_delete, bool row_axis) {
  std::uint32_t& line = row_axis ? p.row : p.col;
  std::int64_t& off = row_axis ? p.row_off : p.col_off;
  if (line < index) {
    return p;
  }
  if (!is_delete) {
    const std::uint32_t limit = a1::axis_extent(row_axis);
    const std::uint64_t shifted = static_cast<std::uint64_t>(line) + count;
    line = static_cast<std::uint32_t>(std::min<std::uint64_t>(shifted, limit - 1U));
  } else if (line >= index + count) {
    line -= count;
  } else {
    line = index;
    off = 0;
  }
  return p;
}

bool same_point(const AnchorPoint& a, const AnchorPoint& b) {
  return a.row == b.row && a.col == b.col && a.row_off == b.row_off && a.col_off == b.col_off;
}

}  // namespace

pugi::xml_node top_of(const pugi::xml_node& root, pugi::xml_node node) {
  while (node.parent() != root) {
    node = node.parent();
  }
  return node;
}

pugi::xml_node find_object(const pugi::xml_node& root, std::uint32_t object_id) {
  return find_drawing_object(root, object_id, is_editable);
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

// A picture's `spPr/a:xfrm` or a graphic frame's `xdr:xfrm`.
pugi::xml_node object_xfrm(const pugi::xml_node& anchor) {
  const pugi::xml_node content = anchor_content(anchor);
  const pugi::xml_node own = child_local(content, "xfrm");
  return own ? own : child_local(child_local(content, "spPr"), "xfrm");
}

// Excel lays a picture rotated into [45, 135) or [225, 315) degrees out in a
// box with width and height swapped.
bool swaps_box(const pugi::xml_node& xfrm) {
  std::int64_t deg = xfrm.attribute("rot").as_llong() / 60000 % 360;
  deg = deg < 0 ? deg + 360 : deg;
  return (deg >= 45 && deg < 135) || (deg >= 225 && deg < 315);
}

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

// A resolved `set_drawing_anchor` placement.
struct Placement {
  AnchorKind kind = AnchorKind::kOneCell;
  EditAs edit_as = EditAs::kTwoCell;
  AnchorPoint from, to;
  std::int64_t box_w = 0, box_h = 0;
  std::int64_t off_x = 0, off_y = 0;
  std::int64_t cx = 0, cy = 0;
};

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
  const pugi::xml_node from = put_marker(anchor, {}, "from", p.from);
  if (two) {
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

}  // namespace

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
    return invalid("drawing: inserted objects are anchored to one or two cells");
  }
  if (!in_emu_range(options.row_off) || !in_emu_range(options.col_off) || !in_emu_range(cx) || !in_emu_range(cy) ||
      options.row >= Sheet::kMaxRows || options.col >= Sheet::kMaxCols) {
    return invalid("drawing: object offset, size or cell out of range");
  }
  std::size_t total_bytes = len;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    total_bytes += part.bytes.size();
  }
  if (wb.passthrough_parts().size() + wb.sheet_count() + parts + 2U > io::kMaxParts ||
      total_bytes > io::kMaxTotalExtractedBytes) {
    return invalid("drawing: package would exceed the reader's limits");
  }
  const Geometry geo(wb, wb.sheet(sheet));
  InsertBox box;
  box.from = AnchorPoint{options.row, options.col, options.row_off, options.col_off};
  if (!geo.advance(true, box.from.row, box.from.row_off, 0) || !geo.advance(false, box.from.col, box.from.col_off, 0)) {
    return invalid("drawing: anchor lies outside the sheet");
  }
  box.to = box.from;
  if (!geo.advance(true, box.to.row, box.to.row_off, cy) || !geo.advance(false, box.to.col, box.to.col_off, cx)) {
    return invalid("drawing: object extends past the sheet");
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
    return invalid("drawing: anchor could not be built");
  }
  const pugi::xml_node added = last ? last.next_sibling() : root.first_child();
  if (!added || added != root.last_child() || !anchor_cnvpr(added)) {
    return invalid("drawing: anchor could not be built");
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
    return invalid("drawing: image exceeds 32 MiB");
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
    return invalid("drawing: a graphic frame needs a width and a height");
  }
  if (!io::ooxml::is_safe_part_name(target_part_path) || rel_type.empty()) {
    return invalid("drawing: graphic frame target is not a part name");
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

Expected<void, Error> remove_drawing(Workbook& wb, std::size_t sheet, std::uint32_t object_id) {
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  const pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_object(root, object_id);
  if (!target) {
    return invalid(kNoObject);
  }
  return drop_objects(wb, path, doc, {top_of(root, target)});
}

void shift_drawing_anchors(Workbook& wb, std::size_t sheet, std::uint32_t index, std::uint32_t count, bool is_delete,
                           bool row_axis) {
  const std::string path = wb.sheet(sheet).drawing_rel_target();
  pugi::xml_document doc;
  if (path.empty() || count == 0 || !load_drawing_part(wb, path, doc)) {
    return;
  }
  const Geometry geo(wb, wb.sheet(sheet));
  const pugi::xml_node root = doc.document_element();
  const auto line = [row_axis](const AnchorPoint& p) { return row_axis ? p.row : p.col; };
  const auto deleted = [&](const AnchorPoint& p) { return is_delete && line(p) >= index && line(p) - index < count; };
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
    if (obj.edit_as == EditAs::kTwoCell && deleted(obj.from) && deleted(obj.to)) {
      const pugi::xml_node top = top_of(root, anchor);
      if (std::find(removed.begin(), removed.end(), top) == removed.end()) {
        removed.push_back(top);
      }
      continue;
    }
    const AnchorPoint from = shift_point(obj.from, index, count, is_delete, row_axis);
    AnchorPoint to = shift_point(obj.to, index, count, is_delete, row_axis);
    const std::int64_t extent = row_axis ? obj.cy : obj.cx;
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
    static_cast<void>(drop_objects(wb, path, doc, removed));
  } else if (changed) {
    static_cast<void>(io::ooxml::store_part_dom(wb, path, doc));
  }
}

std::int64_t sheet_offset_emu(const Workbook& wb, std::size_t sheet, bool row_axis, std::uint32_t line) {
  return Geometry(wb, wb.sheet(sheet)).offset(row_axis, line);
}

Expected<void, Error> set_drawing_name(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                       std::string_view name) {
  if (name.empty()) {
    return invalid("drawing: an object name cannot be empty");
  }
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  const pugi::xml_node target = find_object(doc.document_element(), object_id);
  if (!target) {
    return invalid(kNoObject);
  }
  pugi::xml_node cnvpr = anchor_cnvpr(target);
  if (!cnvpr) {
    return invalid(kNoObject);
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
  if (anchor.anchor_kind == AnchorKind::kAbsolute) {
    return invalid("drawing: objects are anchored to one or two cells");
  }
  pugi::xml_document doc;
  if (path.empty()) {
    return invalid(kNoObject);
  }
  RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
  const pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_object(root, object_id);
  if (!target) {
    return invalid(kNoObject);
  }
  const Geometry geo(wb, wb.sheet(sheet));
  const DrawingObject obj = read_drawing_object(target, {});
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
  if (!in_emu_range(anchor.row_off) || !in_emu_range(anchor.col_off) || !in_emu_range(p.cx) || !in_emu_range(p.cy) ||
      anchor.row >= Sheet::kMaxRows || anchor.col >= Sheet::kMaxCols) {
    return invalid("drawing: object offset, size or cell out of range");
  }
  p.from = AnchorPoint{anchor.row, anchor.col, anchor.row_off, anchor.col_off};
  const bool two = p.kind == AnchorKind::kTwoCell;
  if (obj.anchor_kind == p.kind && (!two || obj.edit_as == p.edit_as) && same_point(obj.from, p.from) &&
      cur_cx == p.cx && cur_cy == p.cy) {
    return Expected<void, Error>::Ok();
  }

  if (!geo.advance(true, p.from.row, p.from.row_off, 0) || !geo.advance(false, p.from.col, p.from.col_off, 0)) {
    return invalid("drawing: anchor lies outside the sheet");
  }
  const bool swap = swaps_box(object_xfrm(target));
  p.box_w = swap ? p.cy : p.cx;
  p.box_h = swap ? p.cx : p.cy;
  p.to = p.from;
  if (!geo.advance(true, p.to.row, p.to.row_off, p.box_h) || !geo.advance(false, p.to.col, p.to.col_off, p.box_w)) {
    return invalid("drawing: object extends past the sheet");
  }
  // `a:off` is the unrotated rectangle sharing the box's centre.
  p.off_x = geo.offset(false, p.from.col) + p.from.col_off + (p.box_w - p.cx) / 2;
  p.off_y = geo.offset(true, p.from.row) + p.from.row_off + (p.box_h - p.cy) / 2;
  const pugi::xml_node top = top_of(root, target);
  for (const pugi::xml_node& other : drawing_anchors(root, true)) {
    const DrawingObject candidate = read_drawing_object(other, {});
    if (top_of(root, other) == top && is_editable(candidate.kind) && candidate.object_id == object_id) {
      place_anchor(other, p);
    }
  }
  return io::ooxml::store_part_dom(wb, path, doc);
}

Expected<void, Error> set_drawing_z_order(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                          std::uint32_t index) {
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_object(root, object_id);
  if (!target) {
    return invalid(kNoObject);
  }
  const std::vector<pugi::xml_node> items = drawing_anchors(root, false);
  if (index >= items.size()) {
    return invalid("drawing: z-order position out of range");
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

}  // namespace formulon
