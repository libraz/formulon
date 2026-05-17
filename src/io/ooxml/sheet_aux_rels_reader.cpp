// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

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

Expected<std::vector<std::string>, Error> load_sheet_table_targets(const ZipReader& zip,
                                                                   std::string_view sheet_rels_path,
                                                                   std::string_view sheet_dir) {
  std::vector<std::string> targets;
  auto visit_status = visit_relationship_nodes(zip, sheet_rels_path, "sheet rels",
                                               [&](const pugi::xml_node& rel) -> Expected<void, Error> {
                                                 const std::string_view type = rel.attribute("Type").value();
                                                 const std::string_view target = rel.attribute("Target").value();
                                                 if (target.empty()) {
                                                   return Expected<void, Error>::Ok();
                                                 }
                                                 if (type == kRelTable) {
                                                   auto resolved = resolve_relative_path(sheet_dir, target);
                                                   if (!resolved) {
                                                     return resolved.error();
                                                   }
                                                   targets.push_back(std::move(resolved).value());
                                                 }
                                                 // Other rel types (printerSettings, drawings, comments, ...)
                                                 // are handled by sibling helpers.
                                                 return Expected<void, Error>::Ok();
                                               });
  if (!visit_status) {
    return visit_status.error();
  }
  return targets;
}

Expected<SheetAuxRels, Error> load_sheet_aux_rels(const ZipReader& zip, std::string_view sheet_rels_path,
                                                  std::string_view sheet_dir) {
  SheetAuxRels out;
  auto visit_status = visit_relationship_nodes(zip, sheet_rels_path, "sheet rels",
                                               [&](const pugi::xml_node& rel) -> Expected<void, Error> {
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
                                                   out.vml_path = std::move(resolved).value();
                                                 } else if (type == kRelPrinterSettings) {
                                                   auto resolved = resolve_relative_path(sheet_dir, target);
                                                   if (!resolved) {
                                                     return resolved.error();
                                                   }
                                                   out.printer_settings_rid.assign(rel.attribute("Id").value());
                                                   out.printer_settings_path = std::move(resolved).value();
                                                 }
                                                 return Expected<void, Error>::Ok();
                                               });
  if (!visit_status) {
    return visit_status.error();
  }
  return out;
}

Expected<std::vector<std::string>, Error> load_sheet_pivot_table_targets(const ZipReader& zip,
                                                                         std::string_view sheet_rels_path,
                                                                         std::string_view sheet_dir) {
  std::vector<std::string> targets;
  auto visit_status = visit_relationship_nodes(zip, sheet_rels_path, "sheet rels",
                                               [&](const pugi::xml_node& rel) -> Expected<void, Error> {
                                                 const std::string_view type = rel.attribute("Type").value();
                                                 const std::string_view target = rel.attribute("Target").value();
                                                 if (target.empty()) {
                                                   return Expected<void, Error>::Ok();
                                                 }
                                                 if (type == kRelPivotTable) {
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

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
