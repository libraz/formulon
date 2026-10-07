//
// Reader and in-place editor for the workbook theme part (DrawingML
// `a:theme`), which lives in the workbook's passthrough parts. Edits patch
// only the `a:clrScheme` / `a:fontScheme` elements of the existing part, so
// the unmodelled remainder survives.

#ifndef FORMULON_IO_THEME_PART_H_
#define FORMULON_IO_THEME_PART_H_

#include "theme.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

namespace io {

/// Implements `Workbook::load_theme`.
LoadedTheme load_theme(const Workbook& wb);

/// Implements `Workbook::set_theme_colors`.
Expected<void, Error> set_theme_colors(Workbook& wb, const ThemeColors& colors);

/// Implements `Workbook::set_theme_fonts`.
Expected<void, Error> set_theme_fonts(Workbook& wb, const ThemeFonts& fonts);

/// Implements `Workbook::reset_theme`.
Expected<void, Error> reset_theme(Workbook& wb);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_THEME_PART_H_
