
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
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace ooxml {

Expected<std::string, Error> load_pivot_cache_records_target(const ZipReader& zip, std::string_view definition_path) {
  const std::string rels_path = rels_path_for_part(definition_path);
  if (!zip.has_entry(rels_path)) {
    return std::string{};
  }
  const std::string base_dir = dir_of(definition_path);
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_relationship_document(zip, rels_path, "pivotCache rels", "ooxml_reader", &doc));
  ASSIGN_OR_RETURN(auto targets, collect_relationship_targets(doc, base_dir, kRelPivotCacheRecords));
  // The last records relationship wins when a malformed definition declares several.
  return targets.empty() ? std::string{} : std::move(targets.back());
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
