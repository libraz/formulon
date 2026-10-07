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

Expected<std::vector<io::ooxml::DrawingRel>, Error> load_rels(const std::vector<PassthroughPart>& parts,
                                                              std::string_view path) {
  const std::string rels_path = io::ooxml::rels_path_for_part(path);
  for (const PassthroughPart& part : parts) {
    if (part.path == rels_path) {
      return io::ooxml::parse_part_rels(part.bytes, io::ooxml::dir_of(path));
    }
  }
  return std::vector<io::ooxml::DrawingRel>{};
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

// Removes the top-level drawing elements `tops`, then the relationships no
// remaining element names and the parts they leave unreferenced.
Expected<void, Error> drop_objects(Workbook& wb, const std::string& path, pugi::xml_document& doc,
                                   const std::vector<pugi::xml_node>& tops) {
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_rels(wb.passthrough_parts(), path));
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

constexpr std::uint32_t kSnapshotVersion = 1;
constexpr char kSnapshotMagic[4] = {'F', 'M', 'I', 'S'};

// The `xdr:wsDr` child holding `node`.
pugi::xml_node top_of(const pugi::xml_node& root, pugi::xml_node node) {
  while (node.parent() != root) {
    node = node.parent();
  }
  return node;
}

// The first listed picture anchor with `object_id`, or a null node.
pugi::xml_node find_picture(const pugi::xml_node& root, std::uint32_t object_id) {
  for (const pugi::xml_node& anchor : drawing_anchors(root, false)) {
    const DrawingObject obj = read_drawing_object(anchor, {});
    if (obj.kind == DrawingObjectKind::kPicture && obj.object_id == object_id) {
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

pugi::xml_node pic_xfrm(const pugi::xml_node& anchor) {
  return child_local(child_local(anchor_content(anchor), "spPr"), "xfrm");
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

// A resolved `set_image_anchor` placement.
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
  const pugi::xml_node xfrm = pic_xfrm(anchor);
  if (const pugi::xml_node off = child_local(xfrm, "off")) {
    set_llong(off, "x", p.off_x);
    set_llong(off, "y", p.off_y);
  }
  if (const pugi::xml_node ext = child_local(xfrm, "ext")) {
    set_llong(ext, "cx", p.cx);
    set_llong(ext, "cy", p.cy);
  }
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

// Bounds-checked reader over `snapshot_image` bytes.
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

struct ImageSnapshot {
  std::uint32_t index = 0;
  std::string xml;
  std::vector<SnapshotRel> rels;
};

Expected<ImageSnapshot, Error> decode_snapshot(const std::uint8_t* bytes, std::size_t len) {
  const Error malformed = invalid("drawing: malformed image snapshot");
  if (bytes == nullptr || len < sizeof(kSnapshotMagic) ||
      !std::equal(kSnapshotMagic, kSnapshotMagic + sizeof(kSnapshotMagic), bytes)) {
    return malformed;
  }
  SnapshotReader in(bytes + sizeof(kSnapshotMagic), len - sizeof(kSnapshotMagic));
  ImageSnapshot snap;
  std::uint32_t version = 0;
  std::uint32_t count = 0;
  if (!in.u32(version) || version != kSnapshotVersion || !in.u32(snap.index) || !in.blob(snap.xml) || !in.u32(count)) {
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
  if (!in.done()) {
    return malformed;
  }
  return snap;
}

// A free path for a copy of `path`: `xl/media/image1.png` -> `xl/media/imageN.png`.
std::string unused_copy_path(const Workbook& wb, const std::string& path) {
  const std::string ext = io::ooxml::extension_of_part(path);
  std::string stem = ext.empty() ? path : path.substr(0, path.size() - ext.size() - 1);
  while (!stem.empty() && stem.back() >= '0' && stem.back() <= '9') {
    stem.pop_back();
  }
  return unused_part_path(wb, stem, ext.empty() ? "bin" : ext.c_str());
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
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_rels(wb.passthrough_parts(), path));
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
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_rels(wb.passthrough_parts(), path));
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
    if (std::none_of(rels.begin(), rels.end(),
                     [&candidate](const io::ooxml::DrawingRel& rel) { return rel.id == candidate; })) {
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

std::int64_t sheet_offset_emu(const Workbook& wb, std::size_t sheet, bool row_axis, std::uint32_t line) {
  return Geometry(wb, wb.sheet(sheet)).offset(row_axis, line);
}

Expected<void, Error> set_image_anchor(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                       const ImageAnchor& anchor) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if (anchor.anchor_kind == AnchorKind::kAbsolute) {
    return invalid("drawing: images are anchored to one or two cells");
  }
  pugi::xml_document doc;
  if (path.empty()) {
    return invalid("drawing: no picture with that id");
  }
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
  const pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_picture(root, object_id);
  if (!target) {
    return invalid("drawing: no picture with that id");
  }
  const DrawingObject obj = read_drawing_object(target, {});
  Placement p;
  p.kind = anchor.anchor_kind;
  p.edit_as = anchor.edit_as;
  p.cx = anchor.width_emu != 0 ? anchor.width_emu : obj.cx;
  p.cy = anchor.height_emu != 0 ? anchor.height_emu : obj.cy;
  if (!in_emu_range(anchor.row_off) || !in_emu_range(anchor.col_off) || !in_emu_range(p.cx) || !in_emu_range(p.cy) ||
      anchor.row >= Sheet::kMaxRows || anchor.col >= Sheet::kMaxCols) {
    return invalid("drawing: image offset, size or cell out of range");
  }
  p.from = AnchorPoint{anchor.row, anchor.col, anchor.row_off, anchor.col_off};
  const bool two = p.kind == AnchorKind::kTwoCell;
  if (obj.anchor_kind == p.kind && (!two || obj.edit_as == p.edit_as) && same_point(obj.from, p.from) &&
      obj.cx == p.cx && obj.cy == p.cy) {
    return Expected<void, Error>::Ok();
  }

  const Geometry geo(wb, wb.sheet(sheet));
  if (!geo.advance(true, p.from.row, p.from.row_off, 0) || !geo.advance(false, p.from.col, p.from.col_off, 0)) {
    return invalid("drawing: anchor lies outside the sheet");
  }
  const bool swap = swaps_box(pic_xfrm(target));
  p.box_w = swap ? p.cy : p.cx;
  p.box_h = swap ? p.cx : p.cy;
  p.to = p.from;
  if (!geo.advance(true, p.to.row, p.to.row_off, p.box_h) || !geo.advance(false, p.to.col, p.to.col_off, p.box_w)) {
    return invalid("drawing: image extends past the sheet");
  }
  // `a:off` is the unrotated rectangle sharing the box's centre.
  p.off_x = geo.offset(false, p.from.col) + p.from.col_off + (p.box_w - p.cx) / 2;
  p.off_y = geo.offset(true, p.from.row) + p.from.row_off + (p.box_h - p.cy) / 2;
  const pugi::xml_node top = top_of(root, target);
  for (const pugi::xml_node& other : drawing_anchors(root, true)) {
    const DrawingObject candidate = read_drawing_object(other, {});
    if (top_of(root, other) == top && candidate.kind == DrawingObjectKind::kPicture &&
        candidate.object_id == object_id) {
      place_anchor(other, p);
    }
  }
  return io::ooxml::store_part_dom(wb, path, doc);
}

Expected<void, Error> set_image_z_order(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::uint32_t index) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if (path.empty()) {
    return invalid("drawing: no picture with that id");
  }
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
  pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_picture(root, object_id);
  if (!target) {
    return invalid("drawing: no picture with that id");
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

Expected<std::vector<std::uint8_t>, Error> snapshot_image(const Workbook& wb, std::size_t sheet,
                                                          std::uint32_t object_id) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if (path.empty()) {
    return invalid("drawing: no picture with that id");
  }
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_drawing(wb, path, doc));
  const pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_picture(root, object_id);
  if (!target) {
    return invalid("drawing: no picture with that id");
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

  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_rels(wb.passthrough_parts(), path));
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
  for (const io::ooxml::DrawingRel* rel : taken) {
    const pugi::xml_node node =
        rels_doc.child("Relationships").find_child_by_attribute("Relationship", "Id", rel->id.c_str());
    const std::string target_text = node.attribute("Target").value();
    const PassthroughPart* part = rel->external ? nullptr : io::ooxml::find_passthrough_part(wb, rel->target);
    if (part != nullptr && io::ooxml::find_passthrough_part(wb, io::ooxml::rels_path_for_part(part->path)) != nullptr) {
      return unparseable("drawing: a part the picture references has relationships of its own", part->path);
    }
    fits = fits && put_blob(out, rel->id.data(), rel->id.size()) && put_blob(out, rel->type.data(), rel->type.size()) &&
           put_blob(out, target_text.data(), target_text.size());
    put_u32(out, rel->external ? 1U : 0U);
    put_u32(out, part != nullptr ? 1U : 0U);
    if (part != nullptr) {
      fits = fits && put_blob(out, part->path.data(), part->path.size()) &&
             put_blob(out, part->content_type.data(), part->content_type.size()) &&
             put_blob(out, part->bytes.data(), part->bytes.size());
    }
  }
  if (!fits) {
    return invalid("drawing: picture is too large to snapshot");
  }
  return out;
}

Expected<std::uint32_t, Error> restore_image(Workbook& wb, std::size_t sheet, const std::uint8_t* bytes,
                                             std::size_t len, std::uint32_t flags) {
  ASSIGN_OR_RETURN(std::string path, editable_drawing_path(wb, sheet));
  if ((flags & ~kImageRestoreNewId) != 0) {
    return invalid("drawing: unknown image restore flags");
  }
  ASSIGN_OR_RETURN(ImageSnapshot snap, decode_snapshot(bytes, len));
  pugi::xml_document fragment;
  std::size_t elements = 0;
  if (io::load_xml_buffer(fragment, std::vector<std::uint8_t>(snap.xml.begin(), snap.xml.end()), "drawing",
                          "snapshot")) {
    for (pugi::xml_node child = fragment.first_child(); child; child = child.next_sibling()) {
      elements += child.type() == pugi::node_element ? 1U : 0U;
    }
  }
  const pugi::xml_node element = fragment.document_element();
  pugi::xml_node picture;
  for (const pugi::xml_node& anchor : drawing_anchors(fragment, false)) {
    if (elements == 1 && !picture && read_drawing_object(anchor, {}).kind == DrawingObjectKind::kPicture) {
      picture = anchor;
    }
  }
  if (!picture) {
    return invalid("drawing: image snapshot holds no picture");
  }
  std::uint32_t id = read_drawing_object(picture, {}).object_id;
  const bool new_id = (flags & kImageRestoreNewId) != 0;

  std::size_t total_bytes = 0;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    total_bytes += part.bytes.size();
  }
  for (const SnapshotRel& rel : snap.rels) {
    total_bytes += rel.bytes.size();
  }
  if (wb.passthrough_parts().size() + wb.sheet_count() + snap.rels.size() + 2U > io::kMaxParts ||
      total_bytes > io::kMaxTotalExtractedBytes) {
    return invalid("drawing: package would exceed the reader's limits");
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
  pugi::xml_node root = doc.document_element();
  ensure_attr(root, "xmlns:xdr", kNsXdr);
  ensure_attr(root, "xmlns:a", kNsA);
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_rels(wb.passthrough_parts(), path));

  pugi::xml_node replaced;
  if (new_id) {
    std::uint32_t next = 0;
    for_each_element(doc, [&next](const pugi::xml_node& node) {
      if (local_name(node) == "cNvPr") {
        next = std::max(next, node.attribute("id").as_uint());
      }
    });
    next = std::max(next, 1U) + 1U;
    const std::uint32_t old_id = id;
    for_each_element(fragment, [old_id, next](const pugi::xml_node& node) {
      if (local_name(node) == "cNvPr" && node.attribute("id").as_uint() == old_id) {
        node.attribute("id").set_value(next);
        drop_creation_id(node);
      }
    });
    id = next;
  } else {
    if (const pugi::xml_node same = find_picture(root, id)) {
      replaced = top_of(root, same);
    }
    bool clash = false;
    for_each_element(root, [&](const pugi::xml_node& node) {
      clash = clash || (local_name(node) == "cNvPr" && node.attribute("id").as_uint() == id &&
                        local_name(node.parent()) != "nvPicPr" && !(replaced && is_within(node, replaced)));
    });
    if (clash) {
      return invalid("drawing: the id belongs to an object that is not a picture");
    }
  }

  // Resolve every relationship before anything in the package changes.
  const std::string dir = io::ooxml::dir_of(path);
  std::vector<std::string> resolved(snap.rels.size());
  for (std::size_t i = 0; i < snap.rels.size(); ++i) {
    const SnapshotRel& rel = snap.rels[i];
    if (rel.external || rel.has_part) {
      resolved[i] = rel.external ? rel.target : rel.part_path;
      continue;
    }
    auto target = io::ooxml::resolve_relative_path(dir, rel.target);
    if (!target) {
      return invalid("drawing: image snapshot target does not resolve");
    }
    resolved[i] = std::move(target.value());
  }

  const std::string rels_path = io::ooxml::rels_path_for_part(path);
  pugi::xml_document rels_doc;
  if (io::ooxml::find_passthrough_part(wb, rels_path) != nullptr) {
    RETURN_IF_ERROR(io::ooxml::load_part_dom(wb, rels_path, "drawing", rels_doc));
  } else {
    rels_doc.load_string(kEmptyRels);
  }
  bool rels_changed = false;
  std::vector<std::pair<std::string, std::string>> renamed;
  for (std::size_t i = 0; i < snap.rels.size(); ++i) {
    const SnapshotRel& rel = snap.rels[i];
    std::string target = resolved[i];
    if (rel.has_part) {
      const PassthroughPart* existing = io::ooxml::find_passthrough_part(wb, target);
      if (existing != nullptr && existing->bytes != rel.bytes) {
        target = unused_copy_path(wb, target);
        existing = nullptr;
      }
      if (existing == nullptr) {
        RETURN_IF_ERROR(wb.add_passthrough_part(PassthroughPart(target, rel.content_type, rel.bytes)));
        const std::string ext = io::ooxml::extension_of_part(target);
        auto info = probe_image(rel.bytes.data(), rel.bytes.size());
        if (rel.content_type.empty() && !ext.empty() && info) {
          ensure_default_content_type(wb, ext.c_str(), image_content_type(info.value().format));
        }
      }
    }
    const auto same = std::find_if(rels.begin(), rels.end(), [&](const io::ooxml::DrawingRel& entry) {
      return entry.type == rel.type && entry.external == rel.external && entry.target == target;
    });
    std::string rid;
    if (same != rels.end()) {
      rid = same->id;
    } else {
      for (std::uint32_t n = 1; rid.empty(); ++n) {
        const std::string candidate = "rId" + std::to_string(n);
        if (std::none_of(rels.begin(), rels.end(),
                         [&candidate](const io::ooxml::DrawingRel& entry) { return entry.id == candidate; })) {
          rid = candidate;
        }
      }
      pugi::xml_node node = rels_doc.child("Relationships").append_child("Relationship");
      node.append_attribute("Id").set_value(rid.c_str());
      node.append_attribute("Type").set_value(rel.type.c_str());
      node.append_attribute("Target").set_value(rel.external ? target.c_str() : relative_target(dir, target).c_str());
      if (rel.external) {
        node.append_attribute("TargetMode").set_value("External");
      }
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
  return id;
}

}  // namespace formulon
