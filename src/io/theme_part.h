//
// Reader and in-place editor for the workbook theme part (DrawingML
// `a:theme`), which lives in the workbook's passthrough parts. Edits patch
// only the `a:clrScheme` / `a:fontScheme` elements of the existing part, so
// the unmodelled remainder survives.

#ifndef FORMULON_IO_THEME_PART_H_
#define FORMULON_IO_THEME_PART_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "theme.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

namespace io {

/// Content type of the theme part.
inline constexpr const char* kThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";

/// Package path of the theme part: the workbook relationship's target when it
/// names one, otherwise the conventional path.
std::string theme_part_path(const Workbook& wb);

/// A minimal valid theme (colour scheme, font scheme, format scheme with three
/// entries per list) carrying the default colours, fonts and script fonts.
std::string build_default_theme_xml();

/// Whether `bytes` is well-formed XML holding a complete colour scheme.
bool is_parseable_theme(const std::vector<std::uint8_t>& bytes);

/// Marks every formula dirty when the theme's font faces or script faces now
/// differ from `before` (Normal-font metrics read them).
void mark_formulas_dirty_if_fonts_changed(Workbook& wb, const Theme& before);

/// Implements `Workbook::load_theme`.
LoadedTheme load_theme(const Workbook& wb);

/// Implements `Workbook::set_theme_colors`.
Expected<void, Error> set_theme_colors(Workbook& wb, const ThemeColors& colors);

/// Implements `Workbook::set_theme_fonts`.
Expected<void, Error> set_theme_fonts(Workbook& wb, const ThemeFonts& fonts);

/// Implements `Workbook::set_theme_script_font`.
Expected<void, Error> set_theme_script_font(Workbook& wb, bool major, std::string_view script,
                                            std::string_view typeface);

/// Implements `Workbook::reset_theme`.
Expected<void, Error> reset_theme(Workbook& wb);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_THEME_PART_H_
