#include <algorithm>
#include <string_view>
#include <utility>
#include <vector>

#include "chart/chart_ns.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_edit_internal.h"
#include "drawing/drawing_geometry.h"
#include "drawing/image_header.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml/relationship_xml.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {

using namespace drawing_edit;

namespace {

// Version 1 captures no dependent parts; version 2 appends them.
constexpr std::uint32_t kSnapshotVersion = 2;

constexpr std::uint32_t kSnapshotVersionNoParts = 1;

constexpr char kSnapshotMagic[4] = {'F', 'M', 'I', 'S'};

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
  bool has_part_graph = false;
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
  snap.has_part_graph = version == kSnapshotVersion;
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

Expected<void, Error> put_part(Workbook& wb, const std::string& path, const std::string& content_type,
                               const std::vector<std::uint8_t>& bytes) {
  if (io::ooxml::find_passthrough_part(wb, path) != nullptr) {
    return wb.replace_passthrough_part(path, bytes);
  }
  return wb.add_passthrough_part(PassthroughPart(path, content_type, bytes));
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
// parts it replaces or removes.
Expected<void, Error> write_parts(Workbook& wb, const std::vector<PlannedPart>& plan, std::vector<std::string>& orphans,
                                  bool has_part_graph) {
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
    if (part.rels == nullptr && !has_part_graph) {
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
    if (part.rels == nullptr) {
      if (has_part_graph && io::ooxml::find_passthrough_part(wb, rels_path) != nullptr) {
        std::vector<std::uint8_t> erase(wb.passthrough_parts().size(), 0U);
        for (std::size_t i = 0; i < wb.passthrough_parts().size(); ++i) {
          if (wb.passthrough_parts()[i].path == rels_path) {
            erase[i] = 1U;
          }
        }
        RETURN_IF_ERROR(wb.erase_passthrough_parts(std::move(erase)));
      }
      continue;
    }
    RETURN_IF_ERROR(put_part(wb, rels_path, "", *part.rels));
    pugi::xml_document rels_doc;
    RETURN_IF_ERROR(io::ooxml::load_part_dom(wb, rels_path, "drawing", rels_doc));
    bool moved = false;
    const pugi::xml_node root = io::ooxml::relationship_root(rels_doc);
    if (!root) {
      return unparseable("drawing: rels part has an invalid Relationships root", rels_path);
    }
    for (pugi::xml_node rel = root.first_child(); rel; rel = rel.next_sibling()) {
      if (!io::ooxml::is_relationship_element(rel, root)) {
        continue;
      }
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
    const pugi::xml_node rels_root = io::ooxml::relationship_root(rels_doc);
    pugi::xml_node node;
    for (pugi::xml_node child = rels_root.first_child(); child; child = child.next_sibling()) {
      if (io::ooxml::is_relationship_element(child, rels_root) &&
          std::string_view(child.attribute("Id").value()) == rel->id) {
        node = child;
        break;
      }
    }
    if (!node) {
      return unparseable("drawing: relationship is missing from rels part", rel->id);
    }
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
  io::ensure_attr(root, "xmlns:xdr", chart_ns::kNsXdr);
  io::ensure_attr(root, "xmlns:a", chart_ns::kNsA);
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, path));

  pugi::xml_node replaced;
  if (new_id) {
    ASSIGN_OR_RETURN(const std::uint32_t next, next_object_id(doc));
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
  RETURN_IF_ERROR(write_parts(wb, plan, orphans, snap.has_part_graph));
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
      if (name.find(':') == std::string_view::npos || strings::starts_with(name, "xmlns:")) {
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
    if (strings::starts_with(attr.name(), "xmlns") && in_root && std::string_view(in_root.value()) == attr.value()) {
      inserted.remove_attribute(attr);
    }
    attr = next;
  }

  RETURN_IF_ERROR(add_empty_part(wb, path, chart_ns::kCtDrawing));
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
