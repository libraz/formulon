// Record types of the generated preset geometry table (preset_geometry_table.inc). The table is data only; the
// evaluator in preset_geometry.cpp reads it.

#ifndef FORMULON_IO_OOXML_PRESET_GEOMETRY_TABLE_H_
#define FORMULON_IO_OOXML_PRESET_GEOMETRY_TABLE_H_

#include <cstdint>
#include <string_view>

namespace formulon {
namespace preset_geometry {

/// A shared operand: `kind` 0 is the constant `arg`; 1, 2, 3 and 4 are w, h, ss and ls divided by `arg`.
struct PresetValue {
  std::uint8_t kind;
  std::int32_t arg;
};

/// An adjust guide: an index into `kPresetAdjustNames` and its default value.
struct PresetAdjust {
  std::uint8_t name;
  std::int32_t value;
};

/// One preset: its adjust guides (`kPresetAdjusts[adjust_first ..]`), and at `kPresetCode[code_offset]` its
/// `guide_count` guide instructions followed by `site_count` connection sites (x, y, then angle operand).
struct PresetEntry {
  std::uint16_t adjust_first;
  std::uint8_t adjust_count;
  std::uint16_t code_offset;
  std::uint8_t guide_count;
  std::uint8_t site_count;
};

}  // namespace preset_geometry
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_PRESET_GEOMETRY_TABLE_H_
