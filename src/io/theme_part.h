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

#include "excel_profile.h"
#include "passthrough_part.h"
#include "styles.h"
#include "theme.h"
#include "unknown_relationship.h"
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

/// Gives `font` the name of the theme's face for `locale`'s script when its scheme is minor or major: the script entry
/// for a locale with one (ja, ko, zh, th), else the latin face. A missing or empty entry for the script leaves the
/// name, as does a font without a scheme. Only the name changes. Idempotent.
void resolve_scheme_font_name(FontRecord& font, const Theme& theme, ExcelLocale locale);

/// `resolve_scheme_font_name` for `font` under `wb`'s theme and profile; the theme is read only for a scheme font.
void resolve_scheme_font_name(FontRecord& font, const Workbook& wb);

/// `resolve_scheme_font_name` for every font of `wb` (the style table and the fonts note runs carry), under its theme
/// and profile. Run after each event that can change the result: a load, a font write, a theme edit, a profile change.
void resolve_scheme_font_names(Workbook& wb);

/// Implements `Workbook::load_theme`.
LoadedTheme load_theme(const Workbook& wb);

/// Removes `defaultThemeVersion` from `<workbookPr>`, as Excel does after it
/// changes a theme: Mac Excel opens a book naming that version with the built-in
/// theme in place of the part. An empty fragment becomes `<workbookPr/>` so the
/// xlsb writer does not write its default version back.
void drop_default_theme_version(Workbook& wb);

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
