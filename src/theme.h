//
// Workbook theme model: colour scheme and major/minor fonts of the theme part.
//
// `Workbook::load_theme` and the theme setters read and patch the part
// through `io/theme_part.h`.

#ifndef FORMULON_THEME_H_
#define FORMULON_THEME_H_

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>

namespace formulon {

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

/// The current Excel Office theme (Aptos Display / Aptos Narrow, Yu Gothic for East Asian).
const Theme& default_theme();

}  // namespace formulon

#endif  // FORMULON_THEME_H_
