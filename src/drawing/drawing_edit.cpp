#include "drawing/drawing_edit.h"

#include <algorithm>
#include <cmath>
#include <string_view>
#include <utility>

#include "drawing/image_header.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_defs.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "print/sheet_geometry.h"
#include "sheet.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr const char* kCtDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";
constexpr const char* kRelImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
constexpr const char* kNsXdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
constexpr const char* kNsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
constexpr const char* kNsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
constexpr const char* kEmptyRels =
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>";
constexpr std::int64_t kEmuPerPoint = 12700;
constexpr std::int64_t kEmuPerPixel = 9525;
constexpr std::int64_t kEmuLimit = std::int64_t{1} << 31;

Error invalid(const char* message) {
  return make_error(FormulonErrorCode::kInvalidArgument, message);
}

Error unparseable(const char* message, std::string_view detail) {
  return make_error(FormulonErrorCode::kIoDrawingUnparseable, message, std::string(detail));
}

// Cell extents in EMU under the Windows display geometry.
class Geometry {
 public:
  Geometry(const Workbook& wb, const Sheet& sheet)
      : sheet_(sheet), model_(print::resolve_column_width_model(wb.styles(), print::GeometryMode::kDisplay)) {}

  std::int64_t size(bool row_axis, std::uint32_t index) const {
    const double pt = row_axis ? print::effective_row_height_pt(sheet_, index)
                               : print::effective_column_width_pt(sheet_, index, model_);
    return std::llround(pt * kEmuPerPoint);
  }

  std::int64_t offset(bool row_axis, std::uint32_t index) const {
    const double pt = row_axis ? print::row_offset_pt(sheet_, index) : print::column_offset_pt(sheet_, index, model_);
    return std::llround(pt * kEmuPerPoint);
  }

  // Moves the marker (`index`, `off`) `extent` EMU further along the axis.
  // Returns false when the result lies past the sheet's last line.
  bool advance(bool row_axis, std::uint32_t& index, std::int64_t& off, std::int64_t extent) const {
    const std::uint32_t limit = row_axis ? Sheet::kMaxRows : Sheet::kMaxCols;
    std::int64_t rem = off + extent;
    for (std::int64_t s = size(row_axis, index); rem >= s; s = size(row_axis, index)) {
      if (index + 1 >= limit) {
        off = rem;
        return rem <= s;
      }
      rem -= s;
      ++index;
    }
    off = rem;
    return true;
  }

 private:
  const Sheet& sheet_;
  print::ColumnWidthModel model_;
};

template <typename Fn>
void for_each_element(const pugi::xml_node& node, const Fn& fn) {
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element) {
      fn(child);
      for_each_element(child, fn);
    }
  }
}

bool has_retained_drawing(const Sheet& sheet) {
  return std::any_of(sheet.unknown_relationships().begin(), sheet.unknown_relationships().end(),
                     [](const UnknownRelationship& rel) { return rel.type == io::kRelDrawing; });
}

Expected<void, Error> load_drawing(const Workbook& wb, std::string_view path, pugi::xml_document& doc) {
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path);
  if (part == nullptr) {
    return unparseable("drawing: part is missing", path);
  }
  return parse_drawing_part(part->bytes, doc);
}

Expected<std::vector<DrawingRel>, Error> load_rels(const std::vector<PassthroughPart>& parts, std::string_view path) {
  const std::string rels_path = io::ooxml::rels_path_for_part(path);
  for (const PassthroughPart& part : parts) {
    if (part.path == rels_path) {
      return parse_part_rels(part.bytes, io::ooxml::dir_of(path));
    }
  }
  return std::vector<DrawingRel>{};
}

// The path of sheet `sheet`'s typed-editable drawing, or empty when it has
// none. Refuses a drawing retained from an `.xlsb` package.
Expected<std::string, Error> editable_drawing_path(const Workbook& wb, std::size_t sheet) {
  if (sheet >= wb.sheet_count()) {
    return invalid("drawing: sheet index out of range");
  }
  if (has_retained_drawing(wb.sheet(sheet))) {
    return unparseable("drawing: drawing retained from an xlsb package is not editable", wb.sheet(sheet).name());
  }
  return wb.sheet(sheet).drawing_rel_target();
}

// First `<stem><N>.<suffix>` with no part named `<stem><N>.*`.
std::string unused_part_path(const Workbook& wb, const std::string& stem, const char* suffix) {
  for (std::uint32_t n = 1;; ++n) {
    const std::string base = stem + std::to_string(n) + ".";
    const bool taken = std::any_of(wb.passthrough_parts().begin(), wb.passthrough_parts().end(),
                                   [&base](const PassthroughPart& part) { return part.path.rfind(base, 0) == 0; });
    if (!taken) {
      return base + suffix;
    }
  }
}

bool referenced(const Workbook& wb, const std::vector<PassthroughPart>& parts, const std::string& path) {
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const Sheet& sheet = wb.sheet(i);
    if (sheet.drawing_rel_target() == path) {
      return true;
    }
    for (const UnknownRelationship& rel : sheet.unknown_relationships()) {
      if (!rel.target_external && rel.target == path) {
        return true;
      }
    }
  }
  for (const UnknownRelationship& rel : wb.unknown_workbook_rels()) {
    if (!rel.target_external && (rel.target == path || "xl/" + rel.target == path)) {
      return true;
    }
  }
  for (const PassthroughPart& part : parts) {
    if (part.path.size() < 5 || part.path.compare(part.path.size() - 5, 5, ".rels") != 0) {
      continue;
    }
    auto rels = parse_part_rels(part.bytes, io::ooxml::dir_of(io::ooxml::dir_of(part.path)));
    if (!rels) {
      return true;  // An unreadable rels part might reference it.
    }
    for (const DrawingRel& rel : rels.value()) {
      if (!rel.external && rel.target == path) {
        return true;
      }
    }
  }
  return false;
}

// Drops each candidate part nothing references any more, with its rels part,
// and then the targets only that rels part referenced.
void remove_orphans(Workbook& wb, std::vector<std::string> candidates) {
  std::vector<PassthroughPart> parts = wb.passthrough_parts();
  const auto find = [&parts](const std::string& path) {
    return std::find_if(parts.begin(), parts.end(), [&path](const PassthroughPart& p) { return p.path == path; });
  };
  bool removed = false;
  while (!candidates.empty()) {
    const std::string path = std::move(candidates.back());
    candidates.pop_back();
    auto it = find(path);
    if (it == parts.end() || referenced(wb, parts, path)) {
      continue;
    }
    parts.erase(it);
    removed = true;
    auto rels_it = find(io::ooxml::rels_path_for_part(path));
    if (rels_it != parts.end()) {
      if (auto rels = parse_part_rels(rels_it->bytes, io::ooxml::dir_of(path))) {
        for (DrawingRel& rel : rels.value()) {
          if (!rel.external) {
            candidates.push_back(std::move(rel.target));
          }
        }
      }
      parts.erase(rels_it);
    }
  }
  if (removed) {
    wb.set_passthrough_parts(std::move(parts));
  }
}

// Removes the top-level drawing elements `tops`, then the relationships no
// remaining element names and the parts they leave unreferenced.
Expected<void, Error> drop_objects(Workbook& wb, const std::string& path, pugi::xml_document& doc,
                                   const std::vector<pugi::xml_node>& tops) {
  ASSIGN_OR_RETURN(std::vector<DrawingRel> rels, load_rels(wb.passthrough_parts(), path));
  for (const pugi::xml_node& top : tops) {
    top.parent().remove_child(top);
  }
  std::vector<std::string> used;
  for_each_element(doc, [&used](const pugi::xml_node& node) {
    for (pugi::xml_attribute attr = node.first_attribute(); attr; attr = attr.next_attribute()) {
      const std::string_view name = attr.name();
      if (name.find(':') != std::string_view::npos && name.rfind("xmlns:", 0) != 0) {
        used.emplace_back(attr.value());
      }
    }
  });
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  const std::string rels_path = io::ooxml::rels_path_for_part(path);
  pugi::xml_document rels_doc;
  if (rels.empty() || !io::ooxml::load_part_dom(wb, rels_path, "drawing", rels_doc)) {
    return Expected<void, Error>::Ok();
  }
  std::vector<std::string> candidates;
  pugi::xml_node root = rels_doc.child("Relationships");
  for (pugi::xml_node rel = root.child("Relationship"); rel;) {
    pugi::xml_node next = rel.next_sibling("Relationship");
    const std::string id = rel.attribute("Id").value();
    if (std::find(used.begin(), used.end(), id) == used.end()) {
      for (const DrawingRel& entry : rels) {
        if (entry.id == id && !entry.external) {
          candidates.push_back(entry.target);
        }
      }
      root.remove_child(rel);
    }
    rel = next;
  }
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, rels_path, rels_doc));
  remove_orphans(wb, std::move(candidates));
  return Expected<void, Error>::Ok();
}

std::string marker_xml(const char* tag, const AnchorPoint& p) {
  return std::string("<xdr:") + tag + "><xdr:col>" + std::to_string(p.col) + "</xdr:col><xdr:colOff>" +
         std::to_string(p.col_off) + "</xdr:colOff><xdr:row>" + std::to_string(p.row) + "</xdr:row><xdr:rowOff>" +
         std::to_string(p.row_off) + "</xdr:rowOff></xdr:" + tag + ">";
}

void ensure_attr(pugi::xml_node node, const char* name, const char* value) {
  if (!node.attribute(name)) {
    node.append_attribute(name).set_value(value);
  }
}

// Creates an empty part at `path` so `store_part_dom` can fill it.
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
    const std::uint32_t limit = row_axis ? Sheet::kMaxRows : Sheet::kMaxCols;
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

Expected<std::vector<DrawingObject>, Error> list_drawing_objects(const Workbook& wb, std::size_t sheet) {
  if (sheet >= wb.sheet_count()) {
    return invalid("drawing: sheet index out of range");
  }
  std::string path = wb.sheet(sheet).drawing_rel_target();
  for (const UnknownRelationship& rel : wb.sheet(sheet).unknown_relationships()) {
    if (path.empty() && rel.type == io::kRelDrawing && !rel.target_external) {
      path = rel.target;
    }
  }
  if (path.empty()) {
    return std::vector<DrawingObject>{};
  }
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
  ASSIGN_OR_RETURN(std::vector<DrawingRel> rels, load_rels(wb.passthrough_parts(), path));
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
  if (options.anchor_kind == AnchorKind::kAbsolute) {
    return invalid("drawing: images are anchored to one or two cells");
  }
  const std::int64_t cx = options.width_emu != 0 ? options.width_emu : info.px_width * kEmuPerPixel;
  const std::int64_t cy = options.height_emu != 0 ? options.height_emu : info.px_height * kEmuPerPixel;
  if (!in_emu_range(options.row_off) || !in_emu_range(options.col_off) || !in_emu_range(cx) || !in_emu_range(cy) ||
      options.row >= Sheet::kMaxRows || options.col >= Sheet::kMaxCols) {
    return invalid("drawing: image offset, size or cell out of range");
  }
  std::size_t total_bytes = len;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    total_bytes += part.bytes.size();
  }
  if (wb.passthrough_parts().size() + wb.sheet_count() + 3U > io::kMaxParts ||
      total_bytes > io::kMaxTotalExtractedBytes) {
    return invalid("drawing: package would exceed the reader's limits");
  }

  // Canonical markers: the offset lies inside its cell and `to` is `from`
  // plus the image size under the display geometry.
  const Geometry geo(wb, wb.sheet(sheet));
  AnchorPoint from{options.row, options.col, options.row_off, options.col_off};
  if (!geo.advance(true, from.row, from.row_off, 0) || !geo.advance(false, from.col, from.col_off, 0)) {
    return invalid("drawing: anchor lies outside the sheet");
  }
  AnchorPoint to = from;
  if (!geo.advance(true, to.row, to.row_off, cy) || !geo.advance(false, to.col, to.col_off, cx)) {
    return invalid("drawing: image extends past the sheet");
  }

  const bool fresh = path.empty();
  if (fresh) {
    path = unused_part_path(wb, "xl/drawings/drawing", "xml");
  }
  pugi::xml_document doc;
  if (fresh) {
    doc.append_child("xdr:wsDr");
  } else {
    RETURN_IF_ERROR(load_drawing(wb, path, doc));
  }
  ASSIGN_OR_RETURN(std::vector<DrawingRel> rels, load_rels(wb.passthrough_parts(), path));
  std::uint32_t id = 0;
  for_each_element(doc, [&id](const pugi::xml_node& node) {
    if (local_name(node) == "cNvPr") {
      id = std::max(id, node.attribute("id").as_uint());
    }
  });
  id = std::max(id, 1U) + 1U;  // Excel numbers drawing objects from 2.
  std::string rid;
  for (std::uint32_t n = 1; rid.empty(); ++n) {
    const std::string candidate = "rId" + std::to_string(n);
    if (std::none_of(rels.begin(), rels.end(), [&candidate](const DrawingRel& rel) { return rel.id == candidate; })) {
      rid = candidate;
    }
  }
  const std::string media_path = unused_part_path(wb, "xl/media/image", image_extension(info.format));

  std::string xml;
  if (options.anchor_kind == AnchorKind::kOneCell) {
    xml = "<xdr:oneCellAnchor>" + marker_xml("from", from) + "<xdr:ext cx=\"" + std::to_string(cx) + "\" cy=\"" +
          std::to_string(cy) + "\"/>";
  } else {
    xml = "<xdr:twoCellAnchor";
    xml += options.edit_as == EditAs::kOneCell    ? " editAs=\"oneCell\""
           : options.edit_as == EditAs::kAbsolute ? " editAs=\"absolute\""
                                                  : "";
    xml += ">" + marker_xml("from", from) + marker_xml("to", to);
  }
  xml += "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" + std::to_string(id) +
         "\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip "
         "xmlns:r=\"" +
         kNsR + "\" r:embed=\"" + rid +
         "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x=\"" +
         std::to_string(geo.offset(false, from.col) + from.col_off) + "\" y=\"" +
         std::to_string(geo.offset(true, from.row) + from.row_off) + "\"/><a:ext cx=\"" + std::to_string(cx) +
         "\" cy=\"" + std::to_string(cy) +
         "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>" +
         (options.anchor_kind == AnchorKind::kOneCell ? "</xdr:oneCellAnchor>" : "</xdr:twoCellAnchor>");
  pugi::xml_node root = doc.document_element();
  if (!root.append_buffer(xml.data(), xml.size())) {
    return invalid("drawing: anchor could not be built");
  }
  ensure_attr(root, "xmlns:xdr", kNsXdr);
  ensure_attr(root, "xmlns:a", kNsA);
  pugi::xml_node cnvpr = anchor_cnvpr(root.last_child());
  cnvpr.append_attribute("name").set_value(
      (options.name.empty() ? "Picture " + std::to_string(id) : options.name).c_str());
  if (!options.descr.empty()) {
    cnvpr.append_attribute("descr").set_value(options.descr.c_str());
  }

  const std::string rels_path = io::ooxml::rels_path_for_part(path);
  pugi::xml_document rels_doc;
  if (io::ooxml::find_passthrough_part(wb, rels_path) != nullptr) {
    RETURN_IF_ERROR(io::ooxml::load_part_dom(wb, rels_path, "drawing", rels_doc));
  } else {
    rels_doc.load_string(kEmptyRels);
  }
  pugi::xml_node rel = rels_doc.child("Relationships").append_child("Relationship");
  rel.append_attribute("Id").set_value(rid.c_str());
  rel.append_attribute("Type").set_value(kRelImage);
  rel.append_attribute("Target").set_value(relative_target(io::ooxml::dir_of(path), media_path).c_str());

  RETURN_IF_ERROR(wb.add_passthrough_part(
      PassthroughPart(media_path, std::string(), std::vector<std::uint8_t>(bytes, bytes + len))));
  ensure_default_content_type(wb, image_extension(info.format), image_content_type(info.format));
  RETURN_IF_ERROR(add_empty_part(wb, path, kCtDrawing));
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  RETURN_IF_ERROR(add_empty_part(wb, rels_path, ""));
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, rels_path, rels_doc));
  if (fresh) {
    wb.sheet(sheet).set_drawing_rel_target(path);
  }
  return id;
}

Expected<void, Error> remove_image(Workbook& wb, std::size_t sheet, std::uint32_t object_id) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if (path.empty()) {
    return invalid("drawing: no picture with that id");
  }
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
  const pugi::xml_node root = doc.document_element();
  for (const pugi::xml_node& anchor : drawing_anchors(root, false)) {
    const DrawingObject obj = read_drawing_object(anchor, {});
    if (obj.kind == DrawingObjectKind::kPicture && obj.object_id == object_id) {
      pugi::xml_node top = anchor;
      while (top.parent() != root) {
        top = top.parent();
      }
      return drop_objects(wb, path, doc, {top});
    }
  }
  return invalid("drawing: no picture with that id");
}

void shift_drawing_anchors(Workbook& wb, std::size_t sheet, std::uint32_t index, std::uint32_t count, bool is_delete,
                           bool row_axis) {
  const std::string path = wb.sheet(sheet).drawing_rel_target();
  pugi::xml_document doc;
  if (path.empty() || count == 0 || !load_drawing(wb, path, doc)) {
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
      pugi::xml_node top = anchor;
      while (top.parent() != root) {
        top = top.parent();
      }
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

}  // namespace formulon
