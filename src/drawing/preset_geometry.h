// The DrawingML preset shapes (ECMA-376 Part 1, 20.1.10.56 ST_ShapeType, and the geometries annex): their names,
// their adjust guides with defaults, and their connection sites evaluated on a shape's size and adjust values.

#ifndef FORMULON_DRAWING_PRESET_GEOMETRY_H_
#define FORMULON_DRAWING_PRESET_GEOMETRY_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
#include <vector>

namespace formulon {
namespace preset_geometry {

/// The position of `prst` among the preset names (byte order), or unset for a name that is not a preset.
std::optional<std::size_t> index_of(std::string_view prst);

/// The number of presets, and the name at `index`.
std::size_t preset_count();
std::string_view name_at(std::size_t index);

/// An adjust guide of a preset: its name and default value, in the preset's raw DrawingML units.
struct AdjustGuide {
  std::string_view name;
  std::int32_t value = 0;
};

/// The adjust guides of preset `index`, in the order the preset lists them; empty for a preset without any.
std::vector<AdjustGuide> adjust_guides(std::size_t index);

/// A point in a shape's own frame, in EMU from its top-left corner before rotation and flips.
struct Point {
  double x = 0.0;
  double y = 0.0;
};

/// The connection sites of preset `index` on a `w` x `h` EMU shape, in the preset's site order (a file's `idx`).
/// `stated` overrides the defaults of the guides it names; a name the preset does not define is ignored.
std::vector<Point> connection_sites(std::size_t index, double w, double h, const std::vector<AdjustGuide>& stated);

}  // namespace preset_geometry
}  // namespace formulon

#endif  // FORMULON_DRAWING_PRESET_GEOMETRY_H_
