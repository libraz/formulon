
#include "io/ooxml/workbook_rels_reader.h"

#include <string>
#include <string_view>
#include <utility>

#include "io/ooxml/package_validator.h"
#include "io/ooxml/rels_walker.h"
#include "io/ooxml_defs.h"
#include "io/unknown_relationship.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace ooxml {

Expected<WorkbookRels, Error> load_workbook_rels(const ZipReader& zip, std::string_view workbook_path) {
  WorkbookRels rels;
  const std::string rels_path = rels_path_for_part(workbook_path);
  if (!zip.has_entry(rels_path)) {
    // Excel always emits this; treat absence as a broken package.
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "workbook rels: part not found",
                      "context=ooxml_reader rels_path=" + rels_path);
  }

  const std::string base_dir = dir_of(workbook_path);
  auto visit_status = visit_relationship_nodes(
      zip, rels_path, "workbook rels", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        const std::string_view type = attr_str(rel, "Type");
        const std::string_view target = attr_str(rel, "Target");
        if (target.empty()) {
          return Expected<void, Error>::Ok();
        }
        // The workbook `<sheets>` list can point at worksheets as well as
        // chart/dialog/macro (and future extension) sheet types. Every
        // in-package relationship whose type ends in "sheet" is retained
        // here; the orchestrator parses worksheets and preserves all other
        // types as opaque sheet entries instead of rejecting the workbook.
        const bool is_sheet_type =
            type == kRelWorksheet || (type.size() >= 5U && type.substr(type.size() - 5U) == "sheet");
        if (is_sheet_type) {
          const std::string id(attr_str(rel, "Id"));
          if (id.empty()) {
            return Expected<void, Error>::Ok();
          }
          auto resolved = resolve_relative_path(base_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          WorkbookRels::SheetTarget sheet_target;
          sheet_target.path = std::move(resolved).value();
          sheet_target.relationship_type.assign(type);
          rels.sheet_targets.emplace(id, std::move(sheet_target));
        } else if (type == kRelSharedStrings) {
          // Last writer wins on duplicates (Excel never emits more than one,
          // but defending against malformed inputs costs almost nothing).
          auto resolved = resolve_relative_path(base_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          rels.sst_path = std::move(resolved).value();
        } else if (type == kRelStyles) {
          auto resolved = resolve_relative_path(base_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          rels.styles_path = std::move(resolved).value();
        } else if (type == kRelPivotCacheDefinition) {
          const std::string id(attr_str(rel, "Id"));
          if (id.empty()) {
            return Expected<void, Error>::Ok();
          }
          auto resolved = resolve_relative_path(base_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          rels.pivot_cache_definition_paths_by_rid.emplace(id, std::move(resolved).value());
        } else if (type == kRelExternalLink) {
          const std::string id(attr_str(rel, "Id"));
          if (id.empty()) {
            return Expected<void, Error>::Ok();
          }
          auto resolved = resolve_relative_path(base_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          rels.external_link_paths_by_rid.emplace(id, std::move(resolved).value());
        } else {
          // Unrecognised Type URI: capture verbatim so the writer can
          // re-emit the entry. Without this, the matching part (theme,
          // calcChain, vbaProject, customXml, ...) survives in
          // passthrough but becomes an orphan in the relationship
          // graph and Excel opens the file in "needs repair" mode.
          UnknownRelationship entry;
          entry.id.assign(attr_str(rel, "Id"));
          entry.type.assign(type);
          const bool external = attr_str(rel, "TargetMode") == "External";
          entry.target_external = external;
          if (external) {
            entry.target.assign(target);
          } else {
            auto resolved = resolve_relative_path(base_dir, target);
            if (!resolved) {
              return resolved.error();
            }
            entry.target = std::move(resolved).value();
          }
          rels.unknown_rels.push_back(std::move(entry));
        }
        return Expected<void, Error>::Ok();
      });
  if (!visit_status) {
    return visit_status.error();
  }
  return rels;
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
