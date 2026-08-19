
#include "c_api/parts/xml_fragment.h"

#include <cstddef>
#include <string>
#include <string_view>

#include "pugixml.hpp"

namespace formulon {
namespace c_api {
namespace parts {

FragmentValidation validate_single_element_fragment(std::string_view fragment, std::string_view expected_name) {
  // The full parse mode keeps declarations, comments, processing
  // instructions and doctypes in the document tree so they can be
  // rejected as top-level siblings instead of being silently discarded.
  // The caller retains the original bytes; this parser owns its copy.
  pugi::xml_document doc;
  const pugi::xml_parse_result parse =
      doc.load_buffer(fragment.data(), fragment.size(), pugi::parse_full, pugi::encoding_utf8);
  if (!parse) {
    return {false, "pugixml=parse_failed description=" + std::string(parse.description()) +
                       " offset=" + std::to_string(parse.offset)};
  }

  pugi::xml_node root;
  std::size_t top_level_count = 0;
  for (pugi::xml_node node : doc.children()) {
    ++top_level_count;
    if (node.type() != pugi::node_element) {
      return {false, "pugixml=top_level_node_rejected type=" + std::to_string(static_cast<int>(node.type())) +
                         " index=" + std::to_string(top_level_count - 1U)};
    }
    if (root) {
      return {false, "pugixml=multiple_top_level_elements count_at_rejection=" + std::to_string(top_level_count)};
    }
    root = node;
  }
  if (!root) {
    return {false, "pugixml=no_top_level_element"};
  }
  if (std::string_view(root.name()) != expected_name) {
    return {false, "pugixml=wrong_root name=" + std::string(root.name()) + " expected=" + std::string(expected_name)};
  }
  return {true, "pugixml=ok top_level_elements=1"};
}

bool fragment_within_size_limit(std::string_view fragment, std::size_t limit_bytes) {
  return fragment.size() <= limit_bytes;
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon
