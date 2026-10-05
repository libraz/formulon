//
// Workbook theme model: colour scheme and major/minor fonts of the theme part.
//
// The theme lives in the workbook's passthrough parts and is parsed on every
// call, never cached. Edits patch only the `a:clrScheme` / `a:fontScheme`
// elements of the existing part, so the unmodelled remainder survives.

#ifndef FORMULON_THEME_H_
#define FORMULON_THEME_H_

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// Number of colours in a theme colour scheme.
constexpr std::size_t kThemeColorCount = 12;

/// Colour scheme order, as written in `a:clrScheme`: dk1, lt1, dk2, lt2,
/// accent1..accent6, hlink, folHlink. Note this is not the order of the
/// `theme="N"` attribute of a spreadsheet colour (see `color_resolve.h`).
using ThemeColors = std::array<std::uint32_t, kThemeColorCount>;  // 0xFFRRGGBB

/// The four font faces of a theme. `*_ea` is `a:ea` when it names a face,
/// otherwise the `a:font script="Jpan"` face.
struct ThemeFonts {
  std::string major_latin;
  std::string major_ea;
  std::string minor_latin;
  std::string minor_ea;
};

struct Theme {
  ThemeColors colors{};
  ThemeFonts fonts;
};

/// Where a loaded theme came from.
enum class ThemeSource : std::uint8_t {
  kPart = 0,         ///< Parsed from the workbook's theme part.
  kDefault = 1,      ///< No theme part: the default theme is reported.
  kUnparseable = 2,  ///< A theme part exists but could not be parsed: the default theme is reported.
};

struct LoadedTheme {
  Theme theme;
  ThemeSource source = ThemeSource::kDefault;
};

/// The Office 2013-2022 theme (Calibri / Calibri Light, Yu Gothic for East Asian).
const Theme& default_theme();

/// Parses the workbook's theme part. Never fails: an absent or unparseable
/// part yields `default_theme()` with the matching `ThemeSource`.
LoadedTheme load_theme(const Workbook& wb);

/// Sets the 12 scheme colours. When the workbook has no theme part, a minimal
/// valid one is generated first, registered with its content type and the
/// workbook relationship. Returns the parse error when an existing part is
/// not a parseable theme (the part is then left untouched).
Expected<void, Error> set_theme_colors(Workbook& wb, const ThemeColors& colors);

/// Sets the four theme fonts; same generation and failure rules as
/// `set_theme_colors`.
Expected<void, Error> set_theme_fonts(Workbook& wb, const ThemeFonts& fonts);

}  // namespace formulon

#endif  // FORMULON_THEME_H_
