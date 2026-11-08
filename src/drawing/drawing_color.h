// DrawingML colour elements (`a:srgbClr`, `a:schemeClr`, `a:sysClr`) as a `ColorSpec`.
//
// A scheme colour whose only transforms are the `lumMod` / `lumOff` pair Excel's palette writes
// reads as a Theme `ColorSpec` (theme index and tint of a cell colour); every other colour reads
// as the rendered Rgb. Theme indices follow the cell numbering: 0 lt1/bg1, 1 dk1/tx1, 2 lt2/bg2,
// 3 dk2/tx2, 4-9 accent1-6, 10 hlink, 11 folHlink.

#ifndef FORMULON_DRAWING_DRAWING_COLOR_H_
#define FORMULON_DRAWING_DRAWING_COLOR_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#include "pugixml.hpp"
#include "styles.h"

namespace formulon {

class Workbook;

namespace drawing_color {

/// Theme index of a scheme colour name (`bg1`/`lt1` -> 0, `accent1` -> 4, ...).
std::optional<std::uint32_t> scheme_index(std::string_view name);

/// True when a `c:clrMapOvr` element remaps `bg1` / `tx1` / `bg2` / `tx2`.
bool remaps_color_map(const pugi::xml_node& clr_map_ovr);

/// The Theme form of a `a:schemeClr` element: set only when its transforms are exactly the
/// `lumMod` / `lumOff` / `alpha` shapes a tint maps to. `remapped` withholds `bg1` / `tx1` /
/// `bg2` / `tx2`.
std::optional<ColorSpec> read_theme(const pugi::xml_node& clr, bool remapped);

/// The rendered colour of a `srgbClr` / `schemeClr` / `sysClr` element as an opaque Rgb
/// (`lumMod` / `lumOff` applied as a tint); unset for any other element.
std::optional<ColorSpec> read_rendered(const Workbook& wb, const pugi::xml_node& clr);

/// Theme when representable, else the rendered Rgb.
std::optional<ColorSpec> read_color(const Workbook& wb, const pugi::xml_node& clr, bool remapped);

/// The `a:alpha/@val` of a colour element; empty when absent.
std::string alpha_of(const pugi::xml_node& clr);

/// True when both colours write the same element (Rgb ignores the alpha byte; a tint compares
/// at its stored precision).
bool same_color(const ColorSpec& a, const ColorSpec& b);

/// Appends the colour element for `color` (Rgb or Theme) to `parent`, carrying `alpha` over when
/// non-empty. `prefix` is the DrawingML namespace prefix with its colon.
void write_color(pugi::xml_node parent, const std::string& prefix, const ColorSpec& color, const std::string& alpha);

}  // namespace drawing_color
}  // namespace formulon

#endif  // FORMULON_DRAWING_DRAWING_COLOR_H_
