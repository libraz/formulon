//
// Shared loader and walker for OOXML `.rels` files. It validates the common
// `<Relationships><Relationship .../></Relationships>` envelope and lets
// consumers revisit one parsed document without extracting it again.
//
// Internal helper for the OOXML reader. Not exposed beyond `src/io/`.

#ifndef FORMULON_IO_OOXML_RELS_WALKER_H_
#define FORMULON_IO_OOXML_RELS_WALKER_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml/package_validator.h"
#include "io/ooxml/relationship_xml.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace ooxml {

/// Reads and validates one relationship document into `out`. `label` and
/// `context` retain the diagnostics used by the ZIP-backed visitor below.
inline Expected<void, Error> load_relationship_document(const ZipReader& zip, std::string_view rels_path,
                                                        std::string_view label, std::string_view context,
                                                        pugi::xml_document* out) {
  auto rels_bytes_or = zip.read_entry(rels_path);
  if (!rels_bytes_or) {
    return std::move(rels_bytes_or.error());
  }
  const std::vector<std::uint8_t>& rels_bytes = rels_bytes_or.value();

  pugi::xml_document& doc = *out;
  pugi::xml_parse_result parse =
      doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=");
    ctx.append(context);
    ctx.append(" part=");
    ctx.append(rels_path);
    ctx.append(" desc=");
    ctx.append(parse.description());
    std::string message(label);
    message.append(": pugixml parse failed");
    return make_error(FormulonErrorCode::kIoXmlParse, std::move(message), std::move(ctx));
  }
  pugi::xml_node root = relationship_root(doc);
  if (!root) {
    std::string ctx("context=");
    ctx.append(context);
    ctx.append(" part=");
    ctx.append(rels_path);
    std::string message(label);
    message.append(": missing <Relationships>");
    return make_error(FormulonErrorCode::kIoRelationshipBroken, std::move(message), std::move(ctx));
  }
  return Expected<void, Error>::Ok();
}

/// Visits validated `<Relationship>` children in document order. The
/// document is validated by `load_relationship_document` before it reaches
/// this overload; a callback error short-circuits the walk.
template <typename Fn>
Expected<void, Error> visit_relationship_nodes(const pugi::xml_document& doc, Fn&& fn) {
  const pugi::xml_node root = relationship_root(doc);
  const bool legacy_root = is_legacy_relationships_root(root);
  for (const pugi::xml_node rel : root.children()) {
    if (!is_relationship_element(rel, legacy_root)) {
      continue;
    }
    auto status = fn(rel);
    if (!status) {
      return std::move(status.error());
    }
  }
  return Expected<void, Error>::Ok();
}

/// Loads and validates `rels_path`, then visits its relationship children.
/// The callback's return type must be `Expected<void, Error>` (or implicitly
/// convertible); a failed status short-circuits the walk.
template <typename Fn>
Expected<void, Error> visit_relationship_nodes(const ZipReader& zip, std::string_view rels_path, std::string_view label,
                                               std::string_view context, Fn&& fn) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_relationship_document(zip, rels_path, label, context, &doc));
  return visit_relationship_nodes(doc, std::forward<Fn>(fn));
}

/// `visit_relationship_nodes` for the OOXML reader, which is the common
/// caller and always reports `context=ooxml_reader`.
template <typename Fn>
Expected<void, Error> visit_relationship_nodes(const ZipReader& zip, std::string_view rels_path, std::string_view label,
                                               Fn&& fn) {
  return visit_relationship_nodes(zip, rels_path, label, "ooxml_reader", std::forward<Fn>(fn));
}

/// Resolves every relationship of `rel_type` in `doc` against `base_dir`, in
/// document order. Relationships without a `Target` are skipped.
inline Expected<std::vector<std::string>, Error> collect_relationship_targets(const pugi::xml_document& doc,
                                                                              std::string_view base_dir,
                                                                              std::string_view rel_type) {
  std::vector<std::string> targets;
  RETURN_IF_ERROR(visit_relationship_nodes(doc, [&](const pugi::xml_node& rel) -> Expected<void, Error> {
    const std::string_view target = rel.attribute("Target").value();
    if (!target.empty() && std::string_view(rel.attribute("Type").value()) == rel_type) {
      ASSIGN_OR_RETURN(auto resolved, resolve_relative_path(base_dir, target));
      targets.push_back(std::move(resolved));
    }
    return Expected<void, Error>::Ok();
  }));
  return targets;
}

/// Captures `rel` verbatim as an `UnknownRelationship`: an in-package target
/// is resolved against `base_dir`, an external one is kept as written.
inline Expected<UnknownRelationship, Error> make_unknown_relationship(const pugi::xml_node& rel,
                                                                      std::string_view base_dir) {
  UnknownRelationship entry;
  entry.id = rel.attribute("Id").value();
  entry.type = rel.attribute("Type").value();
  entry.target_external = std::string_view(rel.attribute("TargetMode").value()) == "External";
  if (entry.target_external) {
    entry.target = rel.attribute("Target").value();
  } else {
    ASSIGN_OR_RETURN(entry.target, resolve_relative_path(base_dir, rel.attribute("Target").value()));
  }
  return entry;
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_RELS_WALKER_H_
