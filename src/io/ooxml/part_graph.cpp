#include "io/ooxml/part_graph.h"

#include <algorithm>
#include <cstring>
#include <utility>

#include "io/ooxml/package_validator.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace ooxml {
namespace {

Error unparseable(std::string_view what, std::string_view detail) {
  return make_error(FormulonErrorCode::kIoDrawingUnparseable, std::string("drawing: ") + std::string(what),
                    std::string(detail));
}

}  // namespace

Expected<std::vector<DrawingRel>, Error> parse_part_rels(const std::vector<std::uint8_t>& bytes,
                                                         std::string_view owner_dir) {
  pugi::xml_document doc;
  if (!load_xml_buffer(doc, bytes, "drawing", "rels")) {
    return unparseable("rels part is not well-formed XML", owner_dir);
  }
  std::vector<DrawingRel> rels;
  for (pugi::xml_node rel = doc.child("Relationships").child("Relationship"); rel;
       rel = rel.next_sibling("Relationship")) {
    DrawingRel entry;
    entry.id = rel.attribute("Id").value();
    entry.type = rel.attribute("Type").value();
    entry.external = std::strcmp(rel.attribute("TargetMode").value(), "External") == 0;
    entry.target = rel.attribute("Target").value();
    if (!entry.external) {
      auto resolved = resolve_relative_path(owner_dir, entry.target);
      if (!resolved) {
        return unparseable("rels target does not resolve", entry.target);
      }
      entry.target = std::move(resolved.value());
    }
    rels.push_back(std::move(entry));
  }
  return rels;
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
    if (!strings::ends_with(part.path, ".rels")) {
      continue;
    }
    auto rels = parse_part_rels(part.bytes, dir_of(dir_of(part.path)));
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
    auto rels_it = find(rels_path_for_part(path));
    if (rels_it != parts.end()) {
      if (auto rels = parse_part_rels(rels_it->bytes, dir_of(path))) {
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

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
