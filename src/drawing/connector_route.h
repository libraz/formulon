#ifndef FORMULON_DRAWING_CONNECTOR_ROUTE_H_
#define FORMULON_DRAWING_CONNECTOR_ROUTE_H_

// Connectors glued to shapes (`stCxn` / `endCxn`) following an edit of the shapes they are glued to.

#include <cstddef>
#include <cstdint>
#include <vector>

#include "pugixml.hpp"

namespace formulon {

class Workbook;

namespace drawing_edit {

/// One glued end of a connector: the connection site it names, in EMU in the connector's own frame.
struct GluedEnd {
  pugi::xml_node connector;  ///< The `cxnSp`.
  bool start = true;         ///< `stCxn`, else `endCxn`.
  std::int64_t x = 0;
  std::int64_t y = 0;
};

/// Every connector end under `root` glued to a shape or picture of the drawing at a site its preset has.
std::vector<GluedEnd> glued_ends(const pugi::xml_node& root);

/// Follows an edit of the drawing `root` of sheet `sheet` as Excel does, given the glued ends before it.
/// A connector glued at both ends, one of whose sites moved, is routed again between its two sites; one glued at
/// a single end moves by that end's displacement. An end whose site the edit left in place moves nothing.
void follow_glued_shapes(const Workbook& wb, std::size_t sheet, const pugi::xml_node& root,
                         const std::vector<GluedEnd>& before);

}  // namespace drawing_edit
}  // namespace formulon

#endif  // FORMULON_DRAWING_CONNECTOR_ROUTE_H_
