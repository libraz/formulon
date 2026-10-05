//
// Colour resolution: theme, tint, indexed and auto colours to effective ARGB.
//
// Theme colours use the `theme="N"` order of a spreadsheet colour (0=lt1,
// 1=dk1, 2=lt2, 3=dk2, 4..11 as in the scheme), tinted with integer HLS
// arithmetic (HLSMAX=240, floor on the new luminance) that reproduces Excel's
// rendered RGB.

#ifndef FORMULON_COLOR_RESOLVE_H_
#define FORMULON_COLOR_RESOLVE_H_

#include <cstdint>
#include <vector>

#include "styles.h"
#include "theme.h"

namespace formulon {

class Workbook;

/// Where a colour is used; decides what an automatic colour means.
enum class ColorContext : std::uint8_t {
  kFont = 0,
  kFillForeground = 1,
  kFillBackground = 2,
  kBorder = 3,
};

/// How a colour was resolved.
enum class ColorResolution : std::uint8_t {
  kExact = 0,             ///< Literal RGB, or theme / palette colour resolved from real data.
  kDefaultTheme = 1,      ///< Theme colour resolved against the default theme (no theme part).
  kIndexOutOfRange = 2,   ///< Theme or palette index beyond the table; black is returned.
  kThemeUnparseable = 3,  ///< Theme colour resolved against the default theme (part unparseable).
  kAutoContext = 4,       ///< Automatic / system colour (or no colour) chosen by context.
};

struct ResolvedColor {
  std::uint32_t argb = 0xFF000000U;
  ColorResolution resolution = ColorResolution::kExact;
};

/// `<indexedColors>` override entries (AARRGGBB); empty selects the built-in
/// palette. Entry `i` overrides palette index `i` when `i < size()`.
using IndexedPalette = std::vector<std::uint32_t>;

/// Built-in palette colour for `index` (0..63) as 0xFFRRGGBB.
std::uint32_t default_indexed_color(std::uint32_t index);

/// Applies a tint (-1..1) to an opaque RGB colour.
std::uint32_t apply_tint(std::uint32_t argb, double tint);

ResolvedColor resolve_color(const ColorSpec& spec, const Theme& theme, ThemeSource source,
                            const IndexedPalette& palette, ColorContext context);

/// Resolves against the workbook's theme and `<indexedColors>`.
ResolvedColor resolve_color(const Workbook& wb, const ColorSpec& spec, ColorContext context);

}  // namespace formulon

#endif  // FORMULON_COLOR_RESOLVE_H_
