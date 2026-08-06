//
// Shared walker for OOXML `.rels` files. Every per-part rels reader
// (workbook rels, sheet aux rels, pivot cache rels, external-link rels)
// loads the same shape — `<Relationships><Relationship .../></Relationships>`
// — and dispatches on `Type=` to interpret each entry.
// `visit_relationship_nodes` centralises the load + parse +
// envelope-validation step so each consumer site reduces to a single
// per-entry lambda.
//
// Internal helper for the OOXML reader. Not exposed beyond `src/io/`.

#ifndef FORMULON_IO_OOXML_RELS_WALKER_H_
#define FORMULON_IO_OOXML_RELS_WALKER_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace ooxml {

/// Loads `rels_path` from `zip`, parses it as a `<Relationships>` document,
/// and invokes `fn(node)` for each `<Relationship>` child in document
/// order. `label` is the human-readable name used in error messages
/// (e.g. `"workbook rels"`, `"sheet rels"`).
///
/// Returns a parse / structure error on malformed input (`kIoXmlParse` /
/// `kIoRelationshipBroken`) or the first error surfaced by `fn`. The
/// callback's return type must be `Expected<void, Error>` (or implicitly
/// convertible); a failed status short-circuits the walk.
template <typename Fn>
Expected<void, Error> visit_relationship_nodes(const ZipReader& zip, std::string_view rels_path, std::string_view label,
                                               Fn&& fn) {
  auto rels_bytes_or = zip.read_entry(rels_path);
  if (!rels_bytes_or) {
    return rels_bytes_or.error();
  }
  const std::vector<std::uint8_t>& rels_bytes = rels_bytes_or.value();

  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=ooxml_reader part=");
    ctx.append(rels_path);
    ctx.append(" desc=");
    ctx.append(parse.description());
    std::string message(label);
    message.append(": pugixml parse failed");
    return make_error(FormulonErrorCode::kIoXmlParse, std::move(message), std::move(ctx));
  }
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    std::string ctx("context=ooxml_reader part=");
    ctx.append(rels_path);
    std::string message(label);
    message.append(": missing <Relationships>");
    return make_error(FormulonErrorCode::kIoRelationshipBroken, std::move(message), std::move(ctx));
  }

  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    auto status = fn(rel);
    if (!status) {
      return status.error();
    }
  }
  return Expected<void, Error>::Ok();
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_RELS_WALKER_H_
