#ifndef FORMULON_IO_OOXML_RELATIONSHIP_XML_H_
#define FORMULON_IO_OOXML_RELATIONSHIP_XML_H_

#include <string>
#include <string_view>

#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"

namespace formulon::io::ooxml {

namespace relationship_xml_detail {

inline constexpr std::string_view kNamespace = kNsPackageRelationships;

inline bool matches(const pugi::xml_node& node, std::string_view local, bool allow_legacy) {
  if (node.type() != pugi::node_element) {
    return false;
  }
  const std::string_view name(node.name());
  const std::size_t colon = name.find(':');
  if (colon == 0U) {
    return false;
  }
  const std::string_view prefix = qname_prefix(name);
  if (local_name(node) != local) {
    return false;
  }
  const std::string_view uri = find_namespace_binding(node, prefix);
  return uri == kNamespace || (allow_legacy && colon == std::string_view::npos && uri.empty());
}

}  // namespace relationship_xml_detail

inline pugi::xml_node relationship_root(const pugi::xml_document& doc) {
  const pugi::xml_node root = doc.document_element();
  return relationship_xml_detail::matches(root, "Relationships", true) ? root : pugi::xml_node{};
}

/// True when `root` is an unprefixed, unbound `<Relationships>`, whose
/// children are then accepted without a namespace.
inline bool is_legacy_relationships_root(const pugi::xml_node& root) {
  return std::string_view(root.name()) == "Relationships" && find_namespace_binding(root, "").empty();
}

/// `is_relationship_element` with the root's legacy status already computed
/// by `is_legacy_relationships_root`, for loops over many children.
inline bool is_relationship_element(const pugi::xml_node& node, bool legacy_root) {
  return relationship_xml_detail::matches(node, "Relationship", legacy_root);
}

inline bool is_relationship_element(const pugi::xml_node& node, const pugi::xml_node& root) {
  return is_relationship_element(node, is_legacy_relationships_root(root));
}

inline bool has_unbound_relationship_prefix(const pugi::xml_node& node) {
  if (node.type() != pugi::node_element) {
    return false;
  }
  const std::string_view name(node.name());
  const std::size_t colon = name.find(':');
  if (colon == std::string_view::npos || name.substr(0, colon) == "xml") {
    return false;
  }
  return colon == 0U || find_namespace_binding(node, name.substr(0, colon)).empty();
}

inline std::string relationship_child_qname(const pugi::xml_node& root) {
  return std::string(qname_prefix_colon(root.name())) + "Relationship";
}

}  // namespace formulon::io::ooxml

#endif  // FORMULON_IO_OOXML_RELATIONSHIP_XML_H_
