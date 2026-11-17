// DrawingML colour elements (`a:srgbClr`, `a:schemeClr`, `a:sysClr`) as a `DrawingColor`: the base
// colour, its ordered transform children (`a:alpha` aside) and the colour Excel renders for them.
//
// Theme indices follow the cell numbering: 0 lt1/bg1, 1 dk1/tx1, 2 lt2/bg2, 3 dk2/tx2, 4-9
// accent1-6, 10 hlink, 11 folHlink. A Theme base with a tint is the input shorthand for its
// `lumMod` / `lumOff` pair; a read never has one.

#ifndef FORMULON_DRAWING_DRAWING_COLOR_H_
#define FORMULON_DRAWING_DRAWING_COLOR_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "pugixml.hpp"
#include "styles.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// One colour transform child (`a:lumMod`, `a:shade`, `a:comp`, ...), by local name.
struct ColorTransform {
  std::string name;
  std::optional<std::int32_t> value;  ///< `val`; unset for an element without one.

  bool operator==(const ColorTransform& o) const { return name == o.name && value == o.value; }
};

/// A DrawingML colour: Rgb or Theme base plus its transforms in document order.
struct DrawingColor {
  ColorSpec base;
  std::vector<ColorTransform> transforms;
  /// Opaque 0xFFRRGGBB as Excel renders the colour; unset when a transform is one the engine
  /// does not model. Computed on read, ignored on input.
  std::optional<std::uint32_t> resolved_argb;
  /// Opacity Excel renders, 0..100000 (100000 opaque); unset while an alpha transform is present or
  /// `resolved_argb` is unset. Computed on read, ignored on input.
  std::optional<std::uint32_t> resolved_alpha;
};

/// The colour and opacity a base and its transforms render as.
struct ResolvedPaint {
  std::optional<std::uint32_t> argb;   ///< Opaque 0xFFRRGGBB.
  std::optional<std::uint32_t> alpha;  ///< 0..100000.
};

namespace drawing_color {

/// Most transforms a colour may carry; a stored colour with more reads as no colour.
// @limit: engine
inline constexpr std::size_t kMaxTransforms = 64;

/// Theme index of a scheme colour name (`bg1`/`lt1` -> 0, `accent1` -> 4, ...).
std::optional<std::uint32_t> scheme_index(std::string_view name);

/// True when a `c:clrMapOvr` element remaps `bg1` / `tx1` / `bg2` / `tx2`.
bool remaps_color_map(const pugi::xml_node& clr_map_ovr);

/// The base and transforms (`alpha` included, in place) of a `srgbClr` / `schemeClr` / `sysClr` element with
/// its rendered colour and opacity;
/// unset for any other element, a scheme name outside the twelve, a `val` that is not an int32,
/// a transform outside the colour's namespace, or more than `kMaxTransforms` transforms. A `bg1` /
/// `tx1` / `bg2` / `tx2` that `clr_map_ovr` remaps reads as the Rgb of the slot it maps to.
std::optional<DrawingColor> read_drawing_color(const Workbook& wb, const pugi::xml_node& clr,
                                               const pugi::xml_node& clr_map_ovr);

/// The base and transforms of a colour element as written, without `c:clrMapOvr` or rendering; unset as
/// `read_drawing_color`.
std::optional<DrawingColor> read_written_color(const pugi::xml_node& clr);

/// The colour and opacity Excel renders for `base_argb` under `transforms`, applied in order. `alpha`
/// leaves the colour as it is. The colour is unset when a transform is not modelled (`alphaMod`,
/// `alphaOff`, an unknown name) or lacks the value it needs; the opacity is 100000 without an alpha
/// transform and unset with one, or when the colour is unset.
ResolvedPaint resolve_transforms(std::uint32_t base_argb, const std::vector<ColorTransform>& transforms);

/// `color` with a Theme tint expanded into its `lumMod` / `lumOff` transforms.
DrawingColor expand_tint(const DrawingColor& color);

/// Checks a colour given as input: an Rgb or Theme base, no tint beside transforms, transform names
/// `[A-Za-z][A-Za-z0-9]*` of at most 32 characters, at most `kMaxTransforms` of them.
Expected<void, Error> validate_input(const DrawingColor& color);

/// True when both colours write the same element (Rgb ignores the alpha byte; tints compare in their
/// expanded form).
bool same_drawing_color(const DrawingColor& a, const DrawingColor& b);

/// Appends the colour element for `color` to `parent`, then `alpha` when non-empty.
void write_drawing_color(pugi::xml_node parent, const std::string& prefix, const DrawingColor& color,
                         const std::string& alpha);

/// The Theme form of a `a:schemeClr` element: set only when its transforms are exactly the
/// `lumMod` / `lumOff` / `alpha` shapes a tint maps to. `remapped` withholds `bg1` / `tx1` /
/// `bg2` / `tx2`.
std::optional<ColorSpec> read_theme(const pugi::xml_node& clr, bool remapped);

/// The rendered colour of a `srgbClr` / `schemeClr` / `sysClr` element as an opaque Rgb
/// (`lumMod` / `lumOff` applied as a tint); unset for any other element.
std::optional<ColorSpec> read_rendered(const Workbook& wb, const pugi::xml_node& clr);

/// Theme when representable, else the rendered Rgb.
std::optional<ColorSpec> read_color(const Workbook& wb, const pugi::xml_node& clr, bool remapped);

/// The colour `write_color` would write back as the same element: the Theme form, or the Rgb of an
/// `a:srgbClr` whose only transform is `alpha`; unset for a colour `read_color` only approximates.
std::optional<ColorSpec> read_exact(const pugi::xml_node& clr, bool remapped);

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
