#include "drawing/drawing_edit.h"

#include <algorithm>
#include <cmath>
#include <string_view>
#include <utility>

#include "drawing/image_header.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "print/sheet_geometry.h"
#include "sheet.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
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

  // EMU from marker `from` to marker `to` along the axis, summed line by line
  // as `advance` walks them.
  std::int64_t span(bool row_axis, const AnchorPoint& from, const AnchorPoint& to) const {
    const std::uint32_t first = row_axis ? from.row : from.col;
    const std::uint32_t last = row_axis ? to.row : to.col;
    std::int64_t extent = (row_axis ? to.row_off : to.col_off) - (row_axis ? from.row_off : from.col_off);
    for (std::uint32_t line = first; line < last; ++line) {
      extent += size(row_axis, line);
    }
    return extent;
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

// The path of sheet `sheet`'s typed-editable drawing, or empty when it has
// none. Refuses a drawing retained from an `.xlsb` package.
Expected<std::string, Error> editable_drawing_path(const Workbook& wb, std::size_t sheet) {
  if (sheet >= wb.sheet_count()) {
    return invalid("drawing: sheet index out of range");
  }
  if (has_retained_drawing(wb.sheet(sheet))) {
    return unparseable("drawing: drawing retained from an xlsb package is not editable", wb.sheet(sheet).name());
  }
  if (wb.sheet(sheet).is_opaque_ooxml_sheet()) {
    return invalid("drawing: only a worksheet's drawing is editable");
  }
  return wb.sheet(sheet).drawing_rel_target();
}

constexpr const char* kNoObject = "drawing: no picture or chart with that id";

// Loads sheet `sheet`'s editable drawing into `doc` and returns its path; a
// sheet with no drawing holds no object.
Expected<std::string, Error> load_object_drawing(const Workbook& wb, std::size_t sheet, pugi::xml_document& doc) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if (path.empty()) {
    return invalid(kNoObject);
  }
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
  return path;
}

// The id for a new object: one past the largest `cNvPr/@id` in `doc`, and at
// least 2 (Excel numbers drawing objects from 2).
std::uint32_t next_object_id(const pugi::xml_document& doc) {
  std::uint32_t id = 0;
  for_each_element(doc, [&id](const pugi::xml_node& node) {
    if (local_name(node) == "cNvPr") {
      id = std::max(id, node.attribute("id").as_uint());
    }
  });
  return std::max(id, 1U) + 1U;
}

// Values of the prefixed, non-`xmlns:` attributes of `node` and its subtree:
// the relationship ids it may name.
std::vector<std::string> relationship_refs(const pugi::xml_node& node) {
  std::vector<std::string> used;
  const auto collect = [&used](const pugi::xml_node& element) {
    for (pugi::xml_attribute attr = element.first_attribute(); attr; attr = attr.next_attribute()) {
      const std::string_view name = attr.name();
      if (name.find(':') != std::string_view::npos && name.rfind("xmlns:", 0) != 0) {
        used.emplace_back(attr.value());
      }
    }
  };
  collect(node);
  for_each_element(node, collect);
  return used;
}

// First `<stem><N>.<suffix>` with no part, and no path in `reserved`, named
// `<stem><N>.*`.
std::string unused_part_path(const Workbook& wb, const std::string& stem, const char* suffix,
                             const std::vector<std::string>& reserved = {}) {
  const auto prefixed = [](const std::string& path, const std::string& base) { return path.rfind(base, 0) == 0; };
  for (std::uint32_t n = 1;; ++n) {
    const std::string base = stem + std::to_string(n) + ".";
    const bool taken =
        std::any_of(wb.passthrough_parts().begin(), wb.passthrough_parts().end(),
                    [&](const PassthroughPart& part) { return prefixed(part.path, base); }) ||
        std::any_of(reserved.begin(), reserved.end(), [&](const std::string& path) { return prefixed(path, base); });
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
  return load_drawing(wb, path, doc);
}

// Removes the top-level drawing elements `tops`, then the relationships no
// remaining element names and the parts they leave unreferenced.
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
  pugi::xml_node root = rels_doc.child("Relationships");
  for (pugi::xml_node rel = root.child("Relationship"); rel;) {
    pugi::xml_node next = rel.next_sibling("Relationship");
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

// Version 1 captures no dependent parts; version 2 appends them.
constexpr std::uint32_t kSnapshotVersion = 2;
constexpr std::uint32_t kSnapshotVersionNoParts = 1;
constexpr char kSnapshotMagic[4] = {'F', 'M', 'I', 'S'};

// The `xdr:wsDr` child holding `node`.
pugi::xml_node top_of(const pugi::xml_node& root, pugi::xml_node node) {
  while (node.parent() != root) {
    node = node.parent();
  }
  return node;
}

// The kinds the drawing edits move, snapshot and remove.
bool is_editable(DrawingObjectKind kind) {
  return kind == DrawingObjectKind::kPicture || kind == DrawingObjectKind::kChart;
}

// The first listed picture or chart anchor with `object_id`, or a null node.
pugi::xml_node find_object(const pugi::xml_node& root, std::uint32_t object_id) {
  for (const pugi::xml_node& anchor : drawing_anchors(root, false)) {
    const DrawingObject obj = read_drawing_object(anchor, {});
    if (is_editable(obj.kind) && obj.object_id == object_id) {
      return anchor;
    }
  }
  return {};
}

// Where a top-level element moved to list position `index` goes: before the
// element holding that item once `moving` is out of the list, or at the end
// (a null node).
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
  const std::string_view name = node.name();
  const std::size_t colon = name.find(':');
  return colon == std::string_view::npos ? std::string() : std::string(name.substr(0, colon + 1));
}

void set_llong(pugi::xml_node node, const char* name, std::int64_t value) {
  pugi::xml_attribute attr = node.attribute(name);
  if (!attr) {
    attr = node.append_attribute(name);
  }
  attr.set_value(static_cast<long long>(value));
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
    set_llong(ext, "cx", p.box_w);
    set_llong(ext, "cy", p.box_h);
  }
  const pugi::xml_node xfrm = object_xfrm(anchor);
  if (const pugi::xml_node off = child_local(xfrm, "off")) {
    set_llong(off, "x", p.off_x);
    set_llong(off, "y", p.off_y);
  }
  if (const pugi::xml_node ext = child_local(xfrm, "ext")) {
    set_llong(ext, "cx", p.cx);
    set_llong(ext, "cy", p.cy);
  }
}

// Copies onto `element` the root's declarations of every prefix its subtree
// uses (in names or in `Requires`/`Ignorable` lists) that it does not declare.
void declare_root_namespaces(const pugi::xml_node& root, pugi::xml_node element) {
  std::vector<std::string> prefixes;
  const auto use = [&prefixes](std::string_view prefix) {
    if (prefix != "xml" && prefix != "xmlns" && std::find(prefixes.begin(), prefixes.end(), prefix) == prefixes.end()) {
      prefixes.emplace_back(prefix);
    }
  };
  const auto visit = [&use](const pugi::xml_node& node) {
    const std::string_view name = node.name();
    use(name.substr(0, name.find(':') == std::string_view::npos ? 0 : name.find(':')));
    for (pugi::xml_attribute attr = node.first_attribute(); attr; attr = attr.next_attribute()) {
      const std::string_view attr_name = attr.name();
      const std::size_t colon = attr_name.find(':');
      if (colon != std::string_view::npos) {
        use(attr_name.substr(0, colon));
      }
      const std::string_view local = colon == std::string_view::npos ? attr_name : attr_name.substr(colon + 1);
      if (local == "Requires" || local == "Ignorable") {
        const std::string_view list = attr.value();
        for (std::size_t pos = 0; pos < list.size();) {
          const std::size_t end = std::min(list.find(' ', pos), list.size());
          if (end > pos) {
            use(list.substr(pos, end - pos));
          }
          pos = end + 1;
        }
      }
    }
  };
  visit(element);
  for_each_element(element, visit);
  for (const std::string& prefix : prefixes) {
    const std::string decl = prefix.empty() ? "xmlns" : "xmlns:" + prefix;
    const pugi::xml_attribute from_root = root.attribute(decl.c_str());
    if (from_root && !element.attribute(decl.c_str())) {
      element.append_attribute(decl.c_str()).set_value(from_root.value());
    }
  }
}

void put_u32(std::vector<std::uint8_t>& out, std::uint32_t value) {
  for (int shift = 0; shift < 32; shift += 8) {
    out.push_back(static_cast<std::uint8_t>(value >> shift));
  }
}

bool put_blob(std::vector<std::uint8_t>& out, const void* data, std::size_t len) {
  if (len > UINT32_MAX) {
    return false;
  }
  put_u32(out, static_cast<std::uint32_t>(len));
  const auto* bytes = static_cast<const std::uint8_t*>(data);
  out.insert(out.end(), bytes, bytes + len);
  return true;
}

// Bounds-checked reader over `snapshot_drawing` bytes.
class SnapshotReader {
 public:
  SnapshotReader(const std::uint8_t* data, std::size_t len) : data_(data), left_(len) {}

  bool u32(std::uint32_t& value) {
    if (left_ < 4) {
      return false;
    }
    value = static_cast<std::uint32_t>(data_[0]) | static_cast<std::uint32_t>(data_[1]) << 8 |
            static_cast<std::uint32_t>(data_[2]) << 16 | static_cast<std::uint32_t>(data_[3]) << 24;
    data_ += 4;
    left_ -= 4;
    return true;
  }

  bool flag(bool& value) {
    std::uint32_t raw = 0;
    if (!u32(raw) || raw > 1) {
      return false;
    }
    value = raw == 1;
    return true;
  }

  template <typename Container>
  bool blob(Container& out) {
    std::uint32_t len = 0;
    if (!u32(len) || len > left_) {
      return false;
    }
    out.assign(data_, data_ + len);
    data_ += len;
    left_ -= len;
    return true;
  }

  bool done() const { return left_ == 0; }

 private:
  const std::uint8_t* data_;
  std::size_t left_;
};

struct SnapshotRel {
  std::string id;
  std::string type;
  std::string target;
  bool external = false;
  bool has_part = false;
  std::string part_path;
  std::string content_type;
  std::vector<std::uint8_t> bytes;
};

// A part reachable from the captured relationships' parts, or a rels part.
struct SnapshotPart {
  std::string path;
  std::string content_type;
  std::vector<std::uint8_t> bytes;
};

struct DrawingSnapshot {
  std::uint32_t index = 0;
  std::string xml;
  std::vector<SnapshotRel> rels;
  std::vector<SnapshotPart> parts;
};

Expected<DrawingSnapshot, Error> decode_snapshot(const std::uint8_t* bytes, std::size_t len) {
  const Error malformed = invalid("drawing: malformed drawing snapshot");
  if (bytes == nullptr || len < sizeof(kSnapshotMagic) ||
      !std::equal(kSnapshotMagic, kSnapshotMagic + sizeof(kSnapshotMagic), bytes)) {
    return malformed;
  }
  SnapshotReader in(bytes + sizeof(kSnapshotMagic), len - sizeof(kSnapshotMagic));
  DrawingSnapshot snap;
  std::uint32_t version = 0;
  std::uint32_t count = 0;
  if (!in.u32(version) || (version != kSnapshotVersion && version != kSnapshotVersionNoParts) || !in.u32(snap.index) ||
      !in.blob(snap.xml) || !in.u32(count)) {
    return malformed;
  }
  for (std::uint32_t i = 0; i < count; ++i) {
    SnapshotRel rel;
    if (!in.blob(rel.id) || !in.blob(rel.type) || !in.blob(rel.target) || !in.flag(rel.external) ||
        !in.flag(rel.has_part)) {
      return malformed;
    }
    if (rel.has_part && (rel.external || !in.blob(rel.part_path) || !in.blob(rel.content_type) || !in.blob(rel.bytes) ||
                         !io::ooxml::is_safe_part_name(rel.part_path))) {
      return malformed;
    }
    snap.rels.push_back(std::move(rel));
  }
  if (version == kSnapshotVersion) {
    if (!in.u32(count)) {
      return malformed;
    }
    for (std::uint32_t i = 0; i < count; ++i) {
      SnapshotPart part;
      if (!in.blob(part.path) || !in.blob(part.content_type) || !in.blob(part.bytes) ||
          !io::ooxml::is_safe_part_name(part.path)) {
        return malformed;
      }
      snap.parts.push_back(std::move(part));
    }
  }
  if (!in.done()) {
    return malformed;
  }
  return snap;
}

// A free path for a copy of `path`: `xl/media/image1.png` -> `xl/media/imageN.png`.
std::string unused_copy_path(const Workbook& wb, const std::string& path, const std::vector<std::string>& reserved) {
  const std::string ext = io::ooxml::extension_of_part(path);
  std::string stem = ext.empty() ? path : path.substr(0, path.size() - ext.size() - 1);
  while (!stem.empty() && stem.back() >= '0' && stem.back() <= '9') {
    stem.pop_back();
  }
  return unused_part_path(wb, stem, ext.empty() ? "bin" : ext.c_str(), reserved);
}

// Removes the `a16:creationId` extension (and an emptied `a:extLst`) from `cnvpr`.
void drop_creation_id(pugi::xml_node cnvpr) {
  pugi::xml_node ext_lst = child_local(cnvpr, "extLst");
  for (pugi::xml_node ext = ext_lst.first_child(); ext;) {
    const pugi::xml_node next = ext.next_sibling();
    if (child_local(ext, "creationId")) {
      ext_lst.remove_child(ext);
    }
    ext = next;
  }
  if (ext_lst && !child_local(ext_lst, "")) {
    cnvpr.remove_child(ext_lst);
  }
}

bool is_within(pugi::xml_node node, const pugi::xml_node& ancestor) {
  for (; node; node = node.parent()) {
    if (node == ancestor) {
      return true;
    }
  }
  return false;
}

bool is_chart_rel(std::string_view type) {
  return strings::ends_with(type, "/chart") || strings::ends_with(type, "/chartEx");
}

// Package parts reachable from the parts at `roots` through rels parts,
// `roots` included, in discovery order.
Expected<std::vector<std::string>, Error> reachable_parts(const Workbook& wb, std::vector<std::string> roots) {
  std::vector<std::string> seen;
  for (std::size_t i = 0; i < roots.size(); ++i) {
    const std::string path = roots[i];
    if (std::find(seen.begin(), seen.end(), path) != seen.end() ||
        io::ooxml::find_passthrough_part(wb, path) == nullptr) {
      continue;
    }
    seen.push_back(path);
    ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, path));
    for (const io::ooxml::DrawingRel& rel : rels) {
      if (!rel.external) {
        roots.push_back(rel.target);
      }
    }
  }
  return seen;
}

// Internal targets of the relationships `node` names in drawing rels `rels`.
std::vector<std::string> named_targets(const pugi::xml_node& node, const std::vector<io::ooxml::DrawingRel>& rels) {
  const std::vector<std::string> used = relationship_refs(node);
  std::vector<std::string> targets;
  for (const io::ooxml::DrawingRel& rel : rels) {
    if (!rel.external && std::find(used.begin(), used.end(), rel.id) != used.end()) {
      targets.push_back(rel.target);
    }
  }
  return targets;
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
  pugi::xml_node node = rels_doc.child("Relationships").append_child("Relationship");
  node.append_attribute("Id").set_value(id.c_str());
  node.append_attribute("Type").set_value(std::string(type).c_str());
  node.append_attribute("Target").set_value(target.c_str());
  if (external) {
    node.append_attribute("TargetMode").set_value("External");
  }
}

Expected<void, Error> put_part(Workbook& wb, const std::string& path, const std::string& content_type,
                               const std::vector<std::uint8_t>& bytes) {
  if (io::ooxml::find_passthrough_part(wb, path) != nullptr) {
    return wb.replace_passthrough_part(path, bytes);
  }
  return wb.add_passthrough_part(PassthroughPart(path, content_type, bytes));
}

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
  ensure_attr(root, "xmlns:xdr", kNsXdr);
  ensure_attr(root, "xmlns:a", kNsA);
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
  RETURN_IF_ERROR(add_empty_part(wb, path, kCtDrawing));
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  RETURN_IF_ERROR(add_empty_part(wb, rels_path, ""));
  RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, rels_path, rels_doc));
  if (fresh) {
    wb.sheet(sheet).set_drawing_rel_target(path);
  }
  return Expected<void, Error>::Ok();
}

// Where `restore_drawing` puts one captured part.
struct PlannedPart {
  std::string recorded;  // Path in the capture.
  std::string path;      // Path it is written to (or shared at).
  const std::string* content_type = nullptr;
  const std::vector<std::uint8_t>* bytes = nullptr;
  const std::vector<std::uint8_t>* rels = nullptr;  // Captured rels part, if any.
  bool chart = false;                               // Reached through a chart relationship.
  bool write = true;                                // False when an identical part is shared.
};

std::size_t plan_index(const std::vector<PlannedPart>& plan, const std::string& recorded) {
  for (std::size_t i = 0; i < plan.size(); ++i) {
    if (plan[i].recorded == recorded) {
      return i;
    }
  }
  return plan.size();
}

// The captured part graph, top-level parts first, then breadth-first through
// the captured rels parts.
Expected<std::vector<PlannedPart>, Error> plan_parts(const DrawingSnapshot& snap) {
  const Error malformed = invalid("drawing: malformed drawing snapshot");
  const auto captured = [&snap](const std::string& path) -> const SnapshotPart* {
    for (const SnapshotPart& part : snap.parts) {
      if (part.path == path) {
        return &part;
      }
    }
    return nullptr;
  };
  std::vector<PlannedPart> plan;
  for (const SnapshotRel& rel : snap.rels) {
    if (rel.has_part && plan_index(plan, rel.part_path) == plan.size()) {
      PlannedPart part;
      part.recorded = rel.part_path;
      part.content_type = &rel.content_type;
      part.bytes = &rel.bytes;
      part.chart = is_chart_rel(rel.type);
      plan.push_back(std::move(part));
    }
  }
  for (std::size_t i = 0; i < plan.size(); ++i) {
    const SnapshotPart* rels = captured(io::ooxml::rels_path_for_part(plan[i].recorded));
    if (rels == nullptr) {
      continue;
    }
    plan[i].rels = &rels->bytes;
    auto children = io::ooxml::parse_part_rels(rels->bytes, io::ooxml::dir_of(plan[i].recorded));
    if (!children) {
      return malformed;
    }
    const bool chart = plan[i].chart;
    for (const io::ooxml::DrawingRel& child : children.value()) {
      const SnapshotPart* part = child.external ? nullptr : captured(child.target);
      if (part != nullptr && plan_index(plan, child.target) == plan.size()) {
        PlannedPart next;
        next.recorded = child.target;
        next.content_type = &part->content_type;
        next.bytes = &part->bytes;
        next.chart = chart || is_chart_rel(child.type);
        plan.push_back(std::move(next));
      }
    }
  }
  return plan;
}

// Picks each planned part's path. `owned` lists the parts the replaced object
// reaches, which are rewritten in place.
void place_parts(const Workbook& wb, std::vector<PlannedPart>& plan, const std::vector<std::string>& owned,
                 bool new_id) {
  std::vector<std::string> reserved;
  for (PlannedPart& part : plan) {
    const PassthroughPart* existing = io::ooxml::find_passthrough_part(wb, part.recorded);
    const bool free = std::find(reserved.begin(), reserved.end(), part.recorded) == reserved.end();
    const bool ours = std::find(owned.begin(), owned.end(), part.recorded) != owned.end();
    if (new_id && part.chart) {
      part.path = unused_copy_path(wb, part.recorded, reserved);
    } else if (free && (existing == nullptr || ours)) {
      part.path = part.recorded;
    } else if (free && !part.chart && part.rels == nullptr && existing->bytes == *part.bytes &&
               io::ooxml::find_passthrough_part(wb, io::ooxml::rels_path_for_part(part.recorded)) == nullptr) {
      part.path = part.recorded;
      part.write = false;
    } else {
      part.path = unused_copy_path(wb, part.recorded, reserved);
    }
    reserved.push_back(part.path);
  }
}

// Writes the planned parts and their rels parts, re-pointing every captured
// relationship whose target moved. Appends to `orphans` the targets of rels
// parts it overwrote.
Expected<void, Error> write_parts(Workbook& wb, const std::vector<PlannedPart>& plan,
                                  std::vector<std::string>& orphans) {
  for (const PlannedPart& part : plan) {
    if (!part.write) {
      continue;
    }
    const bool existed = io::ooxml::find_passthrough_part(wb, part.path) != nullptr;
    RETURN_IF_ERROR(put_part(wb, part.path, *part.content_type, *part.bytes));
    const std::string ext = io::ooxml::extension_of_part(part.path);
    auto info = probe_image(part.bytes->data(), part.bytes->size());
    if (!existed && part.content_type->empty() && !ext.empty() && info) {
      ensure_default_content_type(wb, ext.c_str(), image_content_type(info.value().format));
    }
    if (part.rels == nullptr) {
      continue;
    }
    const std::string rels_path = io::ooxml::rels_path_for_part(part.path);
    if (auto old = load_part_rels(wb, part.path)) {
      for (const io::ooxml::DrawingRel& rel : old.value()) {
        if (!rel.external) {
          orphans.push_back(rel.target);
        }
      }
    }
    RETURN_IF_ERROR(put_part(wb, rels_path, "", *part.rels));
    pugi::xml_document rels_doc;
    RETURN_IF_ERROR(io::ooxml::load_part_dom(wb, rels_path, "drawing", rels_doc));
    bool moved = false;
    for (pugi::xml_node rel = rels_doc.child("Relationships").child("Relationship"); rel;
         rel = rel.next_sibling("Relationship")) {
      if (std::string_view(rel.attribute("TargetMode").value()) == "External") {
        continue;
      }
      auto target = io::ooxml::resolve_relative_path(io::ooxml::dir_of(part.recorded), rel.attribute("Target").value());
      const std::size_t child = target ? plan_index(plan, target.value()) : plan.size();
      if (child == plan.size() || (plan[child].path == plan[child].recorded && part.path == part.recorded)) {
        continue;
      }
      rel.attribute("Target").set_value(relative_target(io::ooxml::dir_of(part.path), plan[child].path).c_str());
      moved = true;
    }
    if (moved) {
      RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, rels_path, rels_doc));
    }
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<std::vector<DrawingObject>, Error> list_drawing_objects(const Workbook& wb, std::size_t sheet) {
  ASSIGN_OR_RETURN(std::string path, sheet_drawing_path(wb, sheet));
  if (path.empty()) {
    return std::vector<DrawingObject>{};
  }
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
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
  const std::uint32_t id = next_object_id(doc);
  const std::string rid = next_rel_id(rels);
  const std::string media_path = unused_part_path(wb, "xl/media/image", image_extension(info.format));
  const std::string xml =
      anchor_open_xml(options, box, cx, cy) + "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" + std::to_string(id) +
      "\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip "
      "xmlns:r=\"" +
      kNsR + "\" r:embed=\"" + rid + "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>" +
      xfrm_xml("a:xfrm", box, cx, cy) + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>" +
      anchor_close_xml(options);
  RETURN_IF_ERROR(
      append_anchor(doc, xml, options.name.empty() ? "Picture " + std::to_string(id) : options.name, options.descr));
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
  const std::uint32_t id = next_object_id(doc);
  const std::string rid = next_rel_id(rels);
  std::string data(graphic_data_xml);
  for (std::size_t at = data.find("{rId}"); at != std::string::npos; at = data.find("{rId}", at + rid.size())) {
    data.replace(at, 5, rid);
  }
  const std::string xml = anchor_open_xml(placement, box, cx, cy) +
                          "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" + std::to_string(id) +
                          "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>" + xfrm_xml("xdr:xfrm", box, cx, cy) +
                          "<a:graphic>" + data + "</a:graphic></xdr:graphicFrame>" + anchor_close_xml(placement);
  RETURN_IF_ERROR(append_anchor(doc, xml, placement.name.empty() ? "Chart " + std::to_string(id) : placement.name,
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
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
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

Expected<std::vector<std::uint8_t>, Error> snapshot_drawing(const Workbook& wb, std::size_t sheet,
                                                            std::uint32_t object_id) {
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  const pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_object(root, object_id);
  if (!target) {
    return invalid(kNoObject);
  }
  const pugi::xml_node top = top_of(root, target);
  const std::vector<pugi::xml_node> items = drawing_anchors(root, false);
  std::uint32_t index = 0;
  while (top_of(root, items[index]) != top) {
    ++index;
  }
  pugi::xml_document copy;
  const pugi::xml_node element = copy.append_copy(top);
  declare_root_namespaces(root, element);
  std::string xml;
  io::append_raw_xml(xml, element);

  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, path));
  pugi::xml_document rels_doc;
  if (!rels.empty()) {
    RETURN_IF_ERROR(io::ooxml::load_part_dom(wb, io::ooxml::rels_path_for_part(path), "drawing", rels_doc));
  }
  const std::vector<std::string> used = relationship_refs(top);
  std::vector<const io::ooxml::DrawingRel*> taken;
  for (const io::ooxml::DrawingRel& rel : rels) {
    if (std::find(used.begin(), used.end(), rel.id) != used.end()) {
      taken.push_back(&rel);
    }
  }

  std::vector<std::uint8_t> out(kSnapshotMagic, kSnapshotMagic + sizeof(kSnapshotMagic));
  put_u32(out, kSnapshotVersion);
  put_u32(out, index);
  bool fits = put_blob(out, xml.data(), xml.size());
  put_u32(out, static_cast<std::uint32_t>(taken.size()));
  std::vector<std::string> top_parts;
  for (const io::ooxml::DrawingRel* rel : taken) {
    const pugi::xml_node node =
        rels_doc.child("Relationships").find_child_by_attribute("Relationship", "Id", rel->id.c_str());
    const std::string target_text = node.attribute("Target").value();
    const PassthroughPart* part = rel->external ? nullptr : io::ooxml::find_passthrough_part(wb, rel->target);
    fits = fits && put_blob(out, rel->id.data(), rel->id.size()) && put_blob(out, rel->type.data(), rel->type.size()) &&
           put_blob(out, target_text.data(), target_text.size());
    put_u32(out, rel->external ? 1U : 0U);
    put_u32(out, part != nullptr ? 1U : 0U);
    if (part != nullptr) {
      top_parts.push_back(part->path);
      fits = fits && put_blob(out, part->path.data(), part->path.size()) &&
             put_blob(out, part->content_type.data(), part->content_type.size()) &&
             put_blob(out, part->bytes.data(), part->bytes.size());
    }
  }
  ASSIGN_OR_RETURN(std::vector<std::string> reached, reachable_parts(wb, top_parts));
  std::vector<const PassthroughPart*> dependents;
  for (const std::string& reached_path : reached) {
    if (std::find(top_parts.begin(), top_parts.end(), reached_path) == top_parts.end()) {
      dependents.push_back(io::ooxml::find_passthrough_part(wb, reached_path));
    }
    if (const PassthroughPart* part_rels =
            io::ooxml::find_passthrough_part(wb, io::ooxml::rels_path_for_part(reached_path))) {
      dependents.push_back(part_rels);
    }
  }
  put_u32(out, static_cast<std::uint32_t>(dependents.size()));
  for (const PassthroughPart* part : dependents) {
    fits = fits && put_blob(out, part->path.data(), part->path.size()) &&
           put_blob(out, part->content_type.data(), part->content_type.size()) &&
           put_blob(out, part->bytes.data(), part->bytes.size());
  }
  if (!fits) {
    return invalid("drawing: object is too large to snapshot");
  }
  return out;
}

Expected<std::uint32_t, Error> restore_drawing(Workbook& wb, std::size_t sheet, const std::uint8_t* bytes,
                                               std::size_t len, std::uint32_t flags) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if ((flags & ~kDrawingRestoreNewId) != 0) {
    return invalid("drawing: unknown drawing restore flags");
  }
  ASSIGN_OR_RETURN(DrawingSnapshot snap, decode_snapshot(bytes, len));
  ASSIGN_OR_RETURN(std::vector<PlannedPart> plan, plan_parts(snap));
  pugi::xml_document fragment;
  std::size_t elements = 0;
  if (io::load_xml_buffer(fragment, std::vector<std::uint8_t>(snap.xml.begin(), snap.xml.end()), "drawing",
                          "snapshot")) {
    for (pugi::xml_node child = fragment.first_child(); child; child = child.next_sibling()) {
      elements += child.type() == pugi::node_element ? 1U : 0U;
    }
  }
  const pugi::xml_node element = fragment.document_element();
  pugi::xml_node object;
  for (const pugi::xml_node& anchor : drawing_anchors(fragment, false)) {
    if (elements == 1 && !object && is_editable(read_drawing_object(anchor, {}).kind)) {
      object = anchor;
    }
  }
  if (!object) {
    return invalid("drawing: snapshot holds no picture or chart");
  }
  std::uint32_t id = read_drawing_object(object, {}).object_id;
  const bool new_id = (flags & kDrawingRestoreNewId) != 0;

  std::size_t total_bytes = 0;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    total_bytes += part.bytes.size();
  }
  for (const SnapshotRel& rel : snap.rels) {
    total_bytes += rel.bytes.size();
  }
  for (const SnapshotPart& part : snap.parts) {
    total_bytes += part.bytes.size();
  }
  if (wb.passthrough_parts().size() + wb.sheet_count() + snap.rels.size() + snap.parts.size() + 2U > io::kMaxParts ||
      total_bytes > io::kMaxTotalExtractedBytes) {
    return invalid("drawing: package would exceed the reader's limits");
  }

  const bool fresh = path.empty();
  pugi::xml_document doc;
  RETURN_IF_ERROR(prepare_drawing(wb, path, doc));
  pugi::xml_node root = doc.document_element();
  ensure_attr(root, "xmlns:xdr", kNsXdr);
  ensure_attr(root, "xmlns:a", kNsA);
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, path));

  pugi::xml_node replaced;
  if (new_id) {
    const std::uint32_t next = next_object_id(doc);
    const std::uint32_t old_id = id;
    for_each_element(fragment, [old_id, next](const pugi::xml_node& node) {
      if (local_name(node) == "cNvPr" && node.attribute("id").as_uint() == old_id) {
        node.attribute("id").set_value(next);
        drop_creation_id(node);
      }
    });
    id = next;
  } else {
    if (const pugi::xml_node same = find_object(root, id)) {
      replaced = top_of(root, same);
    }
    bool clash = false;
    for_each_element(root, [&](const pugi::xml_node& node) {
      const std::string_view holder = local_name(node.parent());
      clash = clash || (local_name(node) == "cNvPr" && node.attribute("id").as_uint() == id && holder != "nvPicPr" &&
                        holder != "nvGraphicFramePr" && !(replaced && is_within(node, replaced)));
    });
    if (clash) {
      return invalid("drawing: the id belongs to an object that is not a picture or chart");
    }
  }

  // Resolve every relationship and part path before anything in the package changes.
  std::vector<std::string> owned;
  if (replaced) {
    ASSIGN_OR_RETURN(owned, reachable_parts(wb, named_targets(replaced, rels)));
  }
  place_parts(wb, plan, owned, new_id);
  const std::string dir = io::ooxml::dir_of(path);
  std::vector<std::string> resolved(snap.rels.size());
  for (std::size_t i = 0; i < snap.rels.size(); ++i) {
    const SnapshotRel& rel = snap.rels[i];
    if (rel.external || rel.has_part) {
      resolved[i] = rel.external ? rel.target : plan[plan_index(plan, rel.part_path)].path;
      continue;
    }
    auto target = io::ooxml::resolve_relative_path(dir, rel.target);
    if (!target) {
      return invalid("drawing: snapshot target does not resolve");
    }
    resolved[i] = std::move(target.value());
  }

  const std::string rels_path = io::ooxml::rels_path_for_part(path);
  pugi::xml_document rels_doc;
  RETURN_IF_ERROR(load_rels_doc(wb, rels_path, rels_doc));
  std::vector<std::string> orphans;
  RETURN_IF_ERROR(write_parts(wb, plan, orphans));
  bool rels_changed = false;
  std::vector<std::pair<std::string, std::string>> renamed;
  for (std::size_t i = 0; i < snap.rels.size(); ++i) {
    const SnapshotRel& rel = snap.rels[i];
    const std::string& target = resolved[i];
    const auto same = std::find_if(rels.begin(), rels.end(), [&](const io::ooxml::DrawingRel& entry) {
      return entry.type == rel.type && entry.external == rel.external && entry.target == target;
    });
    std::string rid;
    if (same != rels.end()) {
      rid = same->id;
    } else {
      rid = next_rel_id(rels);
      append_rel(rels_doc, rid, rel.type, rel.external ? target : relative_target(dir, target), rel.external);
      rels.push_back(io::ooxml::DrawingRel{rid, rel.type, target, rel.external});
      rels_changed = true;
    }
    renamed.emplace_back(rel.id, rid);
  }
  const auto rename = [&renamed](const pugi::xml_node& node) {
    for (pugi::xml_attribute attr = node.first_attribute(); attr; attr = attr.next_attribute()) {
      const std::string_view name = attr.name();
      if (name.find(':') == std::string_view::npos || name.rfind("xmlns:", 0) == 0) {
        continue;
      }
      for (const auto& entry : renamed) {
        if (entry.first == attr.value()) {
          attr.set_value(entry.second.c_str());
          break;
        }
      }
    }
  };
  rename(element);
  for_each_element(element, rename);

  const pugi::xml_node before = new_id ? pugi::xml_node() : insertion_point(root, replaced, snap.index);
  pugi::xml_node inserted = before ? root.insert_copy_before(element, before) : root.append_copy(element);
  // Declarations the root already makes are dropped from the restored element.
  for (pugi::xml_attribute attr = inserted.first_attribute(); attr;) {
    const pugi::xml_attribute next = attr.next_attribute();
    const pugi::xml_attribute in_root = root.attribute(attr.name());
    if (std::string_view(attr.name()).rfind("xmlns", 0) == 0 && in_root &&
        std::string_view(in_root.value()) == attr.value()) {
      inserted.remove_attribute(attr);
    }
    attr = next;
  }

  RETURN_IF_ERROR(add_empty_part(wb, path, kCtDrawing));
  if (rels_changed) {
    RETURN_IF_ERROR(add_empty_part(wb, rels_path, ""));
    RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, rels_path, rels_doc));
  }
  if (replaced) {
    RETURN_IF_ERROR(drop_objects(wb, path, doc, {replaced}));
  } else {
    RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  }
  if (fresh) {
    wb.sheet(sheet).set_drawing_rel_target(path);
  }
  io::ooxml::remove_orphans(wb, std::move(orphans));
  return id;
}

}  // namespace formulon
