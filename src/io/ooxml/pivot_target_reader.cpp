// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "io/ooxml/pivot_target_reader.h"

#include <string>
#include <string_view>
#include <utility>

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

Expected<std::string, Error> load_pivot_cache_records_target(const ZipReader& zip, std::string_view definition_path) {
  const std::string rels_path = rels_path_for_part(definition_path);
  if (!zip.has_entry(rels_path)) {
    return std::string{};
  }
  const std::string base_dir = dir_of(definition_path);
  std::string records_target;
  auto visit_status = visit_relationship_nodes(zip, rels_path, "pivotCache rels",
                                               [&](const pugi::xml_node& rel) -> Expected<void, Error> {
                                                 const std::string_view type = rel.attribute("Type").value();
                                                 const std::string_view target = rel.attribute("Target").value();
                                                 if (target.empty()) {
                                                   return Expected<void, Error>::Ok();
                                                 }
                                                 if (type == kRelPivotCacheRecords && records_target.empty()) {
                                                   auto resolved = resolve_relative_path(base_dir, target);
                                                   if (!resolved) {
                                                     return resolved.error();
                                                   }
                                                   records_target = std::move(resolved).value();
                                                 }
                                                 return Expected<void, Error>::Ok();
                                               });
  if (!visit_status) {
    return visit_status.error();
  }
  return records_target;
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
