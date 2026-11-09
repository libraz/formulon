//
// Retention of unmodelled children of an OOXML root element.

#ifndef FORMULON_IO_OOXML_RAW_CHILD_SWEEP_H_
#define FORMULON_IO_OOXML_RAW_CHILD_SWEEP_H_

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet_passthrough.h"
#include "utils/index_sort.h"

namespace formulon {
namespace io {
namespace ooxml {

/// Captures every child of `root` that `is_modelled` rejects into `out`, tagged
/// with its schema slot (`slot_of`, `slot_count` for a name outside the
/// sequence).
///
/// The sweep is by exclusion, so an element no release has modelled yet
/// survives a round trip. A name outside the sequence takes the slot of the
/// child before it, keeping it between the same two siblings. The result is
/// stable-sorted by slot so the writer can emit it with a single forward
/// cursor, which also repairs a source whose children were not in schema order.
template <typename SlotOf, typename IsModelled>
void CaptureUnmodelledChildren(const pugi::xml_node& root, SlotOf slot_of, std::size_t slot_count,
                               IsModelled is_modelled, WorksheetRawExtensions& out) {
  out.clear();
  std::size_t previous_slot = 0;
  for (pugi::xml_node child = root.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    const std::string_view name = child.name();
    const std::size_t slot = slot_of(name);
    if (slot != slot_count) {
      previous_slot = slot;
    }
    if (is_modelled(name)) {
      continue;
    }
    out.push_back(
        WorksheetRawChild{static_cast<std::uint32_t>(slot == slot_count ? previous_slot : slot), raw_xml(child)});
  }
  sort_by_index(out, [](const WorksheetRawChild& lhs, const WorksheetRawChild& rhs) { return lhs.slot < rhs.slot; });
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_RAW_CHILD_SWEEP_H_
