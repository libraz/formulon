#include "style_resolve.h"

#include <algorithm>
#include <optional>
#include <string_view>

#include "cell.h"
#include "sheet.h"
#include "styles.h"
#include "workbook.h"

namespace formulon {
namespace {

struct XfChoice {
  std::uint32_t xf_index = 0;
  StyleSource source = StyleSource::kDefault;
};

XfChoice choose_xf(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  if (const Cell* cell = sheet.cell_at(row, col)) {
    return {cell->xf_index, StyleSource::kCell};
  }
  const SheetLayout& layout = sheet.layout();
  const auto row_it = std::find_if(layout.row_overrides.begin(), layout.row_overrides.end(),
                                   [row](const RowLayout& r) { return r.row == row && r.has_style; });
  if (row_it != layout.row_overrides.end()) {
    return {row_it->style_xf, StyleSource::kRow};
  }
  const auto col_it = std::find_if(layout.columns.begin(), layout.columns.end(), [col](const ColumnLayout& c) {
    return c.first <= col && col <= c.last && c.has_style;
  });
  if (col_it != layout.columns.end()) {
    return {col_it->style_xf, StyleSource::kColumn};
  }
  return {0U, StyleSource::kDefault};
}

/// A record built programmatically may set only its `*_argb` sibling; that
/// literal is then the colour, matching what the writer emits.
ColorSpec with_literal_fallback(const ColorSpec& spec, std::uint32_t argb, std::uint32_t none_value) {
  if (spec.kind != ColorSpec::Kind::kNone || argb == none_value) {
    return spec;
  }
  ColorSpec literal;
  literal.kind = ColorSpec::Kind::kRgb;
  literal.rgb = argb;
  return literal;
}

}  // namespace

EffectiveStyle effective_style(const Workbook& wb, const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  const StylesTable& styles = wb.styles();
  const XfChoice choice = choose_xf(sheet, row, col);
  EffectiveStyle out;
  out.source = choice.source;
  out.xf_index = choice.xf_index < styles.cell_xfs.size() ? choice.xf_index : 0U;
  const CellXf xf = styles.cell_xfs.empty() ? CellXf{} : styles.cell_xfs[out.xf_index];

  const LoadedTheme theme = load_theme(wb);
  const auto resolve = [&](const ColorSpec& spec, ColorContext context) {
    return resolve_color(spec, theme.theme, theme.source, styles.indexed_colors, context);
  };

  out.font_index = xf.font_index < styles.fonts.size() ? xf.font_index : 0U;
  out.fill_index = xf.fill_index < styles.fills.size() ? xf.fill_index : 0U;
  out.border_index = xf.border_index < styles.borders.size() ? xf.border_index : 0U;
  if (out.font_index < styles.fonts.size()) {
    out.font_color = resolve(
        with_literal_fallback(styles.fonts[out.font_index].color, styles.fonts[out.font_index].color_argb, 0xFF000000U),
        ColorContext::kFont);
  } else {
    out.font_color = resolve(ColorSpec{}, ColorContext::kFont);
  }
  if (out.fill_index < styles.fills.size()) {
    const FillRecord& fill = styles.fills[out.fill_index];
    out.fill_pattern = fill.pattern;
    out.fill_foreground = resolve(with_literal_fallback(fill.fg, fill.fg_argb, 0U), ColorContext::kFillForeground);
    out.fill_background = resolve(with_literal_fallback(fill.bg, fill.bg_argb, 0U), ColorContext::kFillBackground);
  } else {
    out.fill_foreground = resolve(ColorSpec{}, ColorContext::kFillForeground);
    out.fill_background = resolve(ColorSpec{}, ColorContext::kFillBackground);
  }
  const BorderRecord border =
      out.border_index < styles.borders.size() ? styles.borders[out.border_index] : BorderRecord{};
  const BorderSide* sides[5] = {&border.left, &border.right, &border.top, &border.bottom, &border.diagonal};
  for (std::size_t i = 0; i < 5; ++i) {
    out.border_colors[i] =
        resolve(with_literal_fallback(sides[i]->color, sides[i]->color_argb, 0U), ColorContext::kBorder);
  }
  const std::optional<std::string_view> code = effective_num_fmt(styles, xf.num_fmt_id);
  out.num_fmt_code = code ? std::string(*code) : std::string("General");
  if (xf.has_protection) {
    out.locked = xf.locked;
    out.hidden = xf.hidden;
  }
  return out;
}

}  // namespace formulon
