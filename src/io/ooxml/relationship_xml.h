#ifndef FORMULON_IO_OOXML_RELATIONSHIP_XML_H_
#define FORMULON_IO_OOXML_RELATIONSHIP_XML_H_

#include <string>
#include <string_view>

#include "io/xml_utils.h"
#include "pugixml.hpp"

namespace formulon::io::ooxml {

namespace relationship_xml_detail {

inline constexpr std::string_view kNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

inline bool matches(const pugi::xml_node& node, std::string_view local, bool allow_legacy) {
  if (node.type() != pugi::node_element) {
    return false;
  }
  const std::string_view name(node.name());
  const std::size_t colon = name.find(':');
  if (colon == 0U) {
    return false;
  }
  const std::string_view prefix = colon == std::string_view::npos ? std::string_view{} : name.substr(0, colon);
  if ((colon == std::string_view::npos ? name : name.substr(colon + 1U)) != local) {
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

inline bool is_relationship_element(const pugi::xml_node& node, const pugi::xml_node& root) {
  const bool legacy = std::string_view(root.name()) == "Relationships" && find_namespace_binding(root, "").empty();
  return relationship_xml_detail::matches(node, "Relationship", legacy);
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
  const std::string_view name(root.name());
  const std::size_t colon = name.find(':');
  return colon == std::string_view::npos ? std::string("Relationship")
                                         : std::string(name.substr(0, colon + 1U)) + "Relationship";
}

}  // namespace formulon::io::ooxml

#endif  // FORMULON_IO_OOXML_RELATIONSHIP_XML_H_
