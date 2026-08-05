//
// Implementation of the Ptg dispatch table. The table is a flat
// `constexpr std::array` sorted by `base_byte` so a binary search lands
// the row in O(log N) without dragging in `<map>` / `<unordered_map>`.
//
// Class-marked Ptgs (the `0x20`/`0x40`/`0x60` trio in [MS-XLSB] §2.5.97)
// are stored once at their Reference-class base byte; `class_from_byte`
// + `is_class_marked` recover the class bits.

#include "io/xlsb/ptg.h"

#include <algorithm>
#include <cstdint>

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Compile-time check that the table is sorted by `base_byte`. Catches
/// re-orderings during edits before the runtime binary search can
/// silently mis-locate a row.
constexpr bool TableIsSortedByBaseByte() {
  for (std::size_t i = 1; i < kPtgInfoCount; ++i) {
    if (kPtgInfoTable[i - 1].base_byte >= kPtgInfoTable[i].base_byte) {
      return false;
    }
  }
  return true;
}

static_assert(TableIsSortedByBaseByte(), "kPtgInfoTable must be sorted by base_byte");

}  // namespace

const PtgInfo* lookup_ptg(std::uint8_t base_byte) {
  // std::lower_bound on the sorted base_byte column. The dispatch table
  // is small (~50 rows), so the branch-misprediction story doesn't
  // matter; readability wins.
  auto it = std::lower_bound(kPtgInfoTable.begin(), kPtgInfoTable.end(), base_byte,
                             [](const PtgInfo& info, std::uint8_t b) { return info.base_byte < b; });
  if (it == kPtgInfoTable.end() || it->base_byte != base_byte) {
    return nullptr;
  }
  return &(*it);
}

const PtgInfo* lookup_ptg_from_wire(std::uint8_t first_byte) {
  // Strip class bits only when the byte falls in the class-marked
  // operand range (0x20..0x7F). Operator Ptgs (0x01..0x1F) and the
  // 0xE0..0xFF extension band are not class-marked, so we look them
  // up verbatim.
  if (first_byte >= 0x20 && first_byte <= 0x7F) {
    return lookup_ptg(static_cast<std::uint8_t>(first_byte & 0x1F) | 0x20);
  }
  return lookup_ptg(first_byte);
}

PtgClass class_from_byte(std::uint8_t first_byte) {
  // Class bits are positions 5..6 of the first byte. The decode is:
  //   0x00 -> Reference
  //   0x20 -> (Reference class encoded; legacy convention places the
  //           Reference variant at this base, with Value/Array offset by
  //           +0x20/+0x40 respectively).
  //   0x40 -> Value
  //   0x60 -> Array
  // For unclassed bytes (0x01..0x1F) the upper bits are zero and we
  // return Reference, which is a harmless default — callers gate on
  // `is_class_marked` first when the distinction matters.
  switch (first_byte & 0x60) {
    case 0x00:
    case 0x20:
      return PtgClass::Reference;
    case 0x40:
      return PtgClass::Value;
    case 0x60:
      return PtgClass::Array;
    default:
      return PtgClass::Reference;
  }
}

bool is_class_marked(PtgKind kind) {
  switch (kind) {
    case PtgKind::Array:
    case PtgKind::Func:
    case PtgKind::FuncVar:
    case PtgKind::Name:
    case PtgKind::Ref:
    case PtgKind::Area:
    case PtgKind::MemArea:
    case PtgKind::MemErr:
    case PtgKind::MemNoMem:
    case PtgKind::MemFunc:
    case PtgKind::RefErr:
    case PtgKind::AreaErr:
    case PtgKind::RefN:
    case PtgKind::AreaN:
    case PtgKind::MemAreaN:
    case PtgKind::MemNoMemN:
    case PtgKind::NameX:
    case PtgKind::Ref3d:
    case PtgKind::Area3d:
    case PtgKind::RefErr3d:
    case PtgKind::AreaErr3d:
      return true;
    default:
      return false;
  }
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
