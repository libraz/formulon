#include "io/ooxml/part_graph.h"

#include <cstddef>
#include <cstring>
#include <unordered_map>
#include <unordered_set>
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
  const std::vector<PassthroughPart>& parts = wb.passthrough_parts();
  if (candidates.empty() || parts.empty()) {
    return;
  }
  const std::size_t part_count = parts.size();
  std::vector<std::uint8_t> tombstones(part_count, 0U);
  struct PathIndex {
    std::vector<std::size_t> indices;
    std::size_t next_live = 0;
  };
  std::unordered_map<std::string, std::size_t> path_indices;
  std::vector<PathIndex> paths;
  paths.reserve(part_count);
  path_indices.reserve(part_count);
  for (std::size_t i = 0; i < part_count; ++i) {
    const auto [it, inserted] = path_indices.try_emplace(parts[i].path, paths.size());
    if (inserted) {
      paths.emplace_back();
    }
    paths[it->second].indices.push_back(i);
  }

  // A malformed rels part leaves the graph incomplete, so nothing is removed.
  std::vector<std::vector<std::string>> outgoing(part_count);
  std::unordered_map<std::string, std::size_t> inbound;
  inbound.reserve(part_count);
  for (std::size_t i = 0; i < part_count; ++i) {
    if (!strings::ends_with(parts[i].path, ".rels")) {
      continue;
    }
    auto rels = parse_part_rels(parts[i].bytes, dir_of(dir_of(parts[i].path)));
    if (!rels) {
      return;
    }
    std::vector<std::string>& targets = outgoing[i];
    targets.reserve(rels.value().size());
    for (const DrawingRel& rel : rels.value()) {
      if (rel.external) {
        continue;
      }
      targets.push_back(rel.target);
      ++inbound[rel.target];
    }
  }

  std::unordered_set<std::string> model_roots;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const Sheet& sheet = wb.sheet(i);
    model_roots.insert(sheet.drawing_rel_target());
    for (const UnknownRelationship& rel : sheet.unknown_relationships()) {
      if (!rel.target_external) {
        model_roots.insert(rel.target);
      }
    }
  }
  for (const UnknownRelationship& rel : wb.unknown_workbook_rels()) {
    if (rel.target_external) {
      continue;
    }
    model_roots.insert(rel.target);
    model_roots.insert("xl/" + rel.target);
  }

  const auto first_live = [&path_indices, &paths, &tombstones](const std::string& path) {
    const auto found = path_indices.find(path);
    if (found == path_indices.end()) {
      return tombstones.size();
    }
    PathIndex& index = paths[found->second];
    while (index.next_live < index.indices.size() && tombstones[index.indices[index.next_live]] != 0U) {
      ++index.next_live;
    }
    return index.next_live == index.indices.size() ? tombstones.size() : index.indices[index.next_live];
  };
  const auto decrement_outgoing = [&inbound](const std::vector<std::string>& targets) {
    for (const std::string& target : targets) {
      const auto found = inbound.find(target);
      if (found != inbound.end() && found->second != 0U) {
        --found->second;
      }
    }
  };
  const auto tombstone = [&tombstones](std::size_t index) {
    if (tombstones[index] != 0U) {
      return;
    }
    tombstones[index] = 1U;
  };

  bool removed_primary = false;
  while (!candidates.empty()) {
    const std::string path = std::move(candidates.back());
    candidates.pop_back();
    const std::size_t primary = first_live(path);
    if (primary == part_count || model_roots.find(path) != model_roots.end()) {
      continue;
    }
    const auto inbound_for_path = inbound.find(path);
    if (inbound_for_path != inbound.end() && inbound_for_path->second != 0U) {
      continue;
    }

    removed_primary = true;
    tombstone(primary);
    if (strings::ends_with(path, ".rels")) {
      decrement_outgoing(outgoing[primary]);
    }

    const std::size_t rels_index = first_live(rels_path_for_part(path));
    if (rels_index != part_count) {
      // LIFO: the last relationship is processed first.
      for (const std::string& target : outgoing[rels_index]) {
        candidates.push_back(target);
      }
      tombstone(rels_index);
      decrement_outgoing(outgoing[rels_index]);
    }
  }
  if (removed_primary) {
    (void)wb.erase_passthrough_parts(std::move(tombstones));
  }
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
