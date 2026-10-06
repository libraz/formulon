//
// Effective cell style: the formatting a cell shows after row, column and cell XF precedence.
//
// The style is read from the selected `cellXfs` record directly (font, fill,
// border and number-format indices as stored); named-style inheritance is not
// recomputed. Conditional formatting is not applied. A storage slot that was
// created only as an internal blank gap is not an owning cell for precedence.

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

/// The lightweight result shared by style inspection and display rendering.
/// `xf_index` is the selected raw cellXfs index and is not range-checked
/// here; a consumer that reads the styles table (e.g. `effective_style`)
/// clamps an index beyond it to xf 0 while preserving `source`.
struct EffectiveXf {
  std::uint32_t xf_index = 0;
  StyleSource source = StyleSource::kDefault;
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

/// Selects an XF using cell, row, column, default precedence. A cell owns the
/// coordinate when it has a formula, a non-blank cached value, a non-zero XF,
/// or an explicit XF assignment (including XF 0); an implicitly materialized
/// blank gap is skipped. Row style is effective only when the row carries
/// `customFormat`. The returned source uses the `StyleSource` values, and
/// an absent override returns xf 0.
EffectiveXf select_effective_xf(const Sheet& sheet, std::uint32_t row, std::uint32_t col);

/// XF precedence: an owning cell, then its row (when it carries a style),
/// then its column, then xf 0. An xf index beyond `cellXfs` falls back to 0.
EffectiveStyle effective_style(const Workbook& wb, const Sheet& sheet, std::uint32_t row, std::uint32_t col);

}  // namespace formulon

#endif  // FORMULON_STYLE_RESOLVE_H_
