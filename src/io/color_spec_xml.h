//
// Read and write OOXML colour elements (rgb, theme and tint, indexed, auto) as ColorSpec.

#ifndef FORMULON_IO_COLOR_SPEC_XML_H_
#define FORMULON_IO_COLOR_SPEC_XML_H_

#include <cstdint>
#include <string>

#include "pugixml.hpp"
#include "styles.h"

namespace formulon {
namespace io {

/// Reads the colour attributes of a `<color>`-like element (`color`, `fgColor`,
/// `bgColor`, ...). Precedence is `rgb`, `theme` (+ `tint`), `indexed`, `auto`.
/// Returns `kNone` when the node is absent or carries none of them.
ColorSpec read_color_spec(const pugi::xml_node& node);

/// Appends the colour attributes (with a leading space) for `spec`. A `kNone`
/// spec emits `rgb="<fallback_argb>"`; every other kind is authoritative.
void append_color_spec_attrs(std::string& out, const ColorSpec& spec, std::uint32_t fallback_argb);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_COLOR_SPEC_XML_H_
