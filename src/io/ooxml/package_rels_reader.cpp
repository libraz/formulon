#include "io/ooxml/package_rels_reader.h"

#include <string>
#include <utility>

#include "io/ooxml/package_validator.h"
#include "io/ooxml/relationship_xml.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace ooxml {

Expected<std::vector<UnknownRelationship>, Error> read_unknown_package_rels(const std::vector<std::uint8_t>& bytes,
                                                                            std::string_view reader_module) {
  const auto context = [&](std::string_view suffix) {
    std::string ctx("context=");
    ctx.append(reader_module);
    ctx.append(" part=_rels/.rels");
    ctx.append(suffix);
    return ctx;
  };
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, bytes, reader_module, "package-level rels"));
  const pugi::xml_node root = relationship_root(doc);
  if (!root) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "package-level rels: missing <Relationships>",
                      context({}));
  }
  const bool legacy_root = is_legacy_relationships_root(root);
  std::vector<UnknownRelationship> result;
  for (const pugi::xml_node rel : root.children()) {
    if (!is_relationship_element(rel, legacy_root)) {
      continue;
    }
    const std::string_view type(rel.attribute("Type").value());
    if (type == kRelOfficeDocument || type == kRelCoreProperties || type == kRelExtendedProperties ||
        type == kRelCustomProperties) {
      continue;
    }
    const bool external = std::string_view(rel.attribute("TargetMode").value()) == "External";
    std::string target(rel.attribute("Target").value());
    if (type.empty() || target.empty()) {
      continue;
    }
    if (!external) {
      if (target.front() == '/') {
        target.erase(0, 1);
      }
      if (!is_safe_part_name(target)) {
        return make_error(FormulonErrorCode::kIoZipSlip, "package relationship target escapes package root",
                          context(" target=" + target));
      }
    }
    result.push_back(
        UnknownRelationship{std::string(rel.attribute("Id").value()), std::string(type), std::move(target), external});
  }
  return result;
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
