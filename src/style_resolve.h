//
// Effective cell style: the formatting a cell shows after row, column and cell XF precedence.
//
// The style is read from the selected `cellXfs` record directly (font, fill,
// border and number-format indices as stored); named-style inheritance is not
// recomputed. Conditional formatting is not applied.

#ifndef FORMULON_STYLE_RESOLVE_H_
#define FORMULON_STYLE_RESOLVE_H_

#include <array>
#include <cstdint>
#include <string>

#include "color_resolve.h"

namespace formulon {

class Workbook;
class Sheet;

/// Which level supplied the XF.
enum class StyleSource : std::uint8_t {
  kCell = 0,
  kRow = 1,
  kColumn = 2,
  kDefault = 3,
};

struct EffectiveStyle {
  std::uint32_t xf_index = 0;
  StyleSource source = StyleSource::kDefault;
  std::uint32_t font_index = 0;
  std::uint32_t fill_index = 0;
  std::uint32_t border_index = 0;
  ResolvedColor font_color;
  std::uint8_t fill_pattern = 0;
  ResolvedColor fill_foreground;
  ResolvedColor fill_background;
  /// Left, right, top, bottom, diagonal.
  std::array<ResolvedColor, 5> border_colors;
  std::string num_fmt_code;
  bool locked = true;
  bool hidden = false;
};

/// XF precedence: an existing cell, then its row (when it carries a style),
/// then its column, then xf 0. An xf index beyond `cellXfs` falls back to 0.
EffectiveStyle effective_style(const Workbook& wb, const Sheet& sheet, std::uint32_t row, std::uint32_t col);

}  // namespace formulon

#endif  // FORMULON_STYLE_RESOLVE_H_
