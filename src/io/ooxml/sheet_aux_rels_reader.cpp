
#include "io/ooxml/sheet_aux_rels_reader.h"

#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml/package_validator.h"
#include "io/ooxml/rels_walker.h"
#include "io/ooxml_defs.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace ooxml {

namespace {

/// Collects every relationship of `rel_type` in the sheet's rels part,
/// resolving each `Target` against `sheet_dir`. Other rel types
/// (printerSettings, drawings, comments, ...) are handled by sibling helpers.
Expected<std::vector<std::string>, Error> load_targets_of_type(const ZipReader& zip, std::string_view sheet_rels_path,
                                                               std::string_view sheet_dir, std::string_view rel_type) {
  std::vector<std::string> targets;
  auto visit_status = visit_relationship_nodes(zip, sheet_rels_path, "sheet rels",
                                               [&](const pugi::xml_node& rel) -> Expected<void, Error> {
                                                 const std::string_view type = rel.attribute("Type").value();
                                                 const std::string_view target = rel.attribute("Target").value();
                                                 if (target.empty()) {
                                                   return Expected<void, Error>::Ok();
                                                 }
                                                 if (type == rel_type) {
                                                   auto resolved = resolve_relative_path(sheet_dir, target);
                                                   if (!resolved) {
                                                     return resolved.error();
                                                   }
                                                   targets.push_back(std::move(resolved).value());
                                                 }
                                                 return Expected<void, Error>::Ok();
                                               });
  if (!visit_status) {
    return visit_status.error();
  }
  return targets;
}

}  // namespace

Expected<std::vector<std::string>, Error> load_sheet_table_targets(const ZipReader& zip,
                                                                   std::string_view sheet_rels_path,
                                                                   std::string_view sheet_dir) {
  return load_targets_of_type(zip, sheet_rels_path, sheet_dir, kRelTable);
}

Expected<SheetAuxRels, Error> load_sheet_aux_rels(const ZipReader& zip, std::string_view sheet_rels_path,
                                                  std::string_view sheet_dir,
                                                  std::string_view legacy_drawing_body_rid) {
  SheetAuxRels out;
  // A sheet may declare up to two `kRelVmlDrawing` relationships (comment
  // geometry and header/footer image); collect every candidate here and
  // resolve which one is the modelled comment-VML slot after the walk,
  // once `out.comments_path` and every candidate id/target is known.
  struct VmlCandidate {
    std::string id;
    std::string path;
  };
  std::vector<VmlCandidate> vml_candidates;
  auto visit_status = visit_relationship_nodes(
      zip, sheet_rels_path, "sheet rels", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        const std::string_view type = rel.attribute("Type").value();
        const std::string_view target = rel.attribute("Target").value();
        if (target.empty()) {
          return Expected<void, Error>::Ok();
        }
        if (type == kRelHyperlink) {
          const std::string id = rel.attribute("Id").value();
          if (id.empty()) {
            return Expected<void, Error>::Ok();
          }
          // Hyperlink targets are external URLs — preserve them exactly as
          // the OOXML producer wrote them (no relative-path resolution).
          out.hyperlink_rid_to_target.emplace(id, std::string(target));
        } else if (type == kRelComments) {
          auto resolved = resolve_relative_path(sheet_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          out.comments_path = std::move(resolved).value();
        } else if (type == kRelVmlDrawing) {
          auto resolved = resolve_relative_path(sheet_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          vml_candidates.push_back(VmlCandidate{std::string(rel.attribute("Id").value()), std::move(resolved).value()});
        } else if (type == kRelPrinterSettings) {
          auto resolved = resolve_relative_path(sheet_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          out.printer_settings_rid.assign(rel.attribute("Id").value());
          out.printer_settings_path = std::move(resolved).value();
        } else if (type == kRelDrawing) {
          auto resolved = resolve_relative_path(sheet_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          out.drawing_path = std::move(resolved).value();
        } else if (type != kRelTable && type != kRelPivotTable) {
          UnknownRelationship entry;
          entry.id = rel.attribute("Id").value();
          entry.type = std::string(type);
          entry.target_external = std::string_view(rel.attribute("TargetMode").value()) == "External";
          if (entry.target_external) {
            entry.target = std::string(target);
          } else {
            auto resolved = resolve_relative_path(sheet_dir, target);
            if (!resolved) {
              return resolved.error();
            }
            entry.target = std::move(resolved).value();
          }
          if (!entry.id.empty()) {
            out.unknown_rels.push_back(std::move(entry));
          }
        }
        return Expected<void, Error>::Ok();
      });
  if (!visit_status) {
    return visit_status.error();
  }
  // Select the one `kRelVmlDrawing` candidate that models comment
  // geometry: prefer the id the worksheet body's `<legacyDrawing>`
  // element names, since that is the unambiguous signal a producer
  // gives us. When nothing matches (no `<legacyDrawing>` element, or its
  // id names no relationship in this file) but the sheet does have
  // comments, fall back to the first candidate in document order so a
  // single-VML sheet still round-trips its comment geometry. Every
  // other candidate is preserved verbatim in `unknown_rels` instead of
  // being silently dropped.
  std::size_t selected = vml_candidates.size();
  if (!legacy_drawing_body_rid.empty()) {
    for (std::size_t i = 0; i < vml_candidates.size(); ++i) {
      if (vml_candidates[i].id == legacy_drawing_body_rid) {
        selected = i;
        break;
      }
    }
  }
  if (selected == vml_candidates.size() && !out.comments_path.empty() && !vml_candidates.empty()) {
    selected = 0;
  }
  for (std::size_t i = 0; i < vml_candidates.size(); ++i) {
    if (i == selected) {
      out.vml_path = vml_candidates[i].path;
      continue;
    }
    if (vml_candidates[i].id.empty()) {
      continue;
    }
    UnknownRelationship entry;
    entry.id = vml_candidates[i].id;
    entry.type = std::string(kRelVmlDrawing);
    entry.target = vml_candidates[i].path;
    entry.target_external = false;
    out.unknown_rels.push_back(std::move(entry));
  }
  return out;
}

Expected<std::vector<std::string>, Error> load_sheet_pivot_table_targets(const ZipReader& zip,
                                                                         std::string_view sheet_rels_path,
                                                                         std::string_view sheet_dir) {
  return load_targets_of_type(zip, sheet_rels_path, sheet_dir, kRelPivotTable);
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
