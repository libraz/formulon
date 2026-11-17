#include "io/ooxml/drawing_guides.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "io/ooxml/preset_geometry.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"

namespace formulon {
namespace drawing_guides {
namespace {

bool contains(const std::vector<std::uint8_t>& bytes, std::string_view text) {
  return std::search(bytes.begin(), bytes.end(), text.begin(), text.end()) != bytes.end();
}

}  // namespace

std::optional<PassthroughPart> without_undefined_guides(const PassthroughPart& part, std::string_view content_type,
                                                        std::size_t* refused) {
  const std::vector<std::uint8_t>& bytes = part.bytes;
  // Most parts state no adjust guide at all; they are not parsed.
  if (content_type != io::kCtDrawing || !contains(bytes, "avLst") ||
      !(contains(bytes, ":gd ") || contains(bytes, "<gd "))) {
    return std::nullopt;
  }
  pugi::xml_document doc;
  if (!io::load_xml_buffer(doc, bytes, "drawing", part.path)) {
    return std::nullopt;
  }
  std::vector<pugi::xml_node> doomed;
  std::size_t counted = 0;
  doc.document_element().find_node([&doomed, &counted](const pugi::xml_node& geom) {
    if (geom.type() != pugi::node_element || io::local_name(geom) != "prstGeom") {
      return false;
    }
    const std::optional<std::size_t> preset = preset_geometry::index_of(geom.attribute("prst").value());
    if (!preset) {
      return false;
    }
    const std::vector<preset_geometry::AdjustGuide> defined = preset_geometry::adjust_guides(*preset);
    std::vector<std::string_view> seen;
    for (const pugi::xml_node& gd : io::children_local(io::child_local(geom, "avLst"), "gd")) {
      const std::string_view name = gd.attribute("name").value();
      const bool known = std::any_of(defined.begin(), defined.end(),
                                     [name](const preset_geometry::AdjustGuide& g) { return g.name == name; });
      const bool repeated = std::find(seen.begin(), seen.end(), name) != seen.end();
      seen.push_back(name);
      if (known && !repeated) {
        continue;
      }
      doomed.push_back(gd);
      counted += defined.empty() ? 0U : 1U;
    }
    return false;
  });
  if (doomed.empty()) {
    return std::nullopt;
  }
  for (const pugi::xml_node& gd : doomed) {
    gd.parent().remove_child(gd);
  }
  *refused += counted;
  std::string xml(io::kXmlDecl);
  io::append_raw_xml(xml, doc.document_element());
  PassthroughPart out = part;
  out.bytes.assign(xml.begin(), xml.end());
  return out;
}

}  // namespace drawing_guides
}  // namespace formulon
