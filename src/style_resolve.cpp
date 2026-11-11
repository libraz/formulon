#include "style_resolve.h"

#include <algorithm>
#include <cassert>
#include <optional>
#include <string_view>
#include <utility>

#include "cell.h"
#include "row_cells.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "styles.h"
#include "theme.h"
#include "workbook.h"

namespace formulon {
bool cell_owns_style(const Cell& cell) noexcept {
  return cell.has_explicit_xf || cell.xf_index != 0U || !cell.formula_text.empty() || !cell.cached_value.is_blank();
}

EffectiveXf select_effective_xf(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  if (const std::optional<StyleSlot> slot = sheet.cell_style_at(row, col)) {
    return {slot->xf, StyleSource::kCell};
  }
  return select_axis_xf(sheet.layout(), row, col);
}

std::vector<AxisSegment> axis_segments(const SheetLayout& layout, std::uint32_t row, std::uint32_t first,
                                       std::uint64_t end) {
  std::vector<AxisSegment> out;
  if (end <= first) {
    return out;
  }
  if (const RowLayout* row_entry = find_row_override(layout, row); row_entry != nullptr && row_entry->has_style) {
    out.push_back({first, end, row_entry->style_xf});
    return out;
  }
  const auto by_first = [](const ColumnLayout& lhs, const ColumnLayout& rhs) { return lhs.first < rhs.first; };
  assert(std::is_sorted(layout.columns.begin(), layout.columns.end(), by_first));
  (void)by_first;
  // The first span that can reach `first`: spans are disjoint, so only the one
  // starting at or before it can straddle it.
  auto it = std::upper_bound(layout.columns.begin(), layout.columns.end(), first,
                             [](std::uint32_t key, const ColumnLayout& c) { return key < c.first; });
  if (it != layout.columns.begin()) {
    --it;
  }
  for (; it != layout.columns.end() && it->first < end; ++it) {
    if (!it->has_style || it->last < first) {
      continue;
    }
    out.push_back({std::max(it->first, first), std::min<std::uint64_t>(static_cast<std::uint64_t>(it->last) + 1U, end),
                   it->style_xf});
  }
  return out;
}

std::vector<AxisSegment> explicit_default_observable(const SheetLayout& layout, std::uint32_t row, std::uint32_t first,
                                                     std::uint64_t end) {
  std::vector<AxisSegment> pieces = axis_segments(layout, row, first, end);
  pieces.erase(std::remove_if(pieces.begin(), pieces.end(), [](const AxisSegment& seg) { return seg.xf == 0U; }),
               pieces.end());
  return pieces;
}

namespace {

/// Columns of non-zero column style across the sheet.
std::uint64_t styled_column_count(const SheetLayout& layout) {
  std::uint64_t count = 0;
  for (const ColumnLayout& span : layout.columns) {
    if (span.has_style && span.style_xf != 0U) {
      count += static_cast<std::uint64_t>(span.last) - span.first + 1U;
    }
  }
  return count;
}

/// Whether the row's runs turn into an xf-0 row format (see `row_emit_plan`).
bool converts_to_row_format(const SheetLayout& layout, std::uint32_t row, const RowCells& cells,
                            bool row_has_phantoms) {
  if (row_has_phantoms || !cells.empty()) {
    return false;
  }
  if (const RowLayout* entry = find_row_override(layout, row); entry != nullptr && entry->has_style) {
    return false;
  }
  std::uint64_t covered = 0;
  for (const StyleRun& run : cells.runs()) {
    if (run.xf != 0U) {
      return false;
    }
    const std::uint64_t end = static_cast<std::uint64_t>(run.last) + 1U;
    std::uint64_t at = run.first;
    for (const AxisSegment& seg : axis_segments(layout, row, run.first, end)) {
      if (seg.xf == 0U || seg.first != at) {
        return false;
      }
      at = seg.end;
    }
    if (at != end) {
      return false;
    }
    covered += end - run.first;
  }
  return covered >= 3U && covered == styled_column_count(layout);
}

}  // namespace

RowEmitPlan row_emit_plan(const Sheet& sheet, std::uint32_t row, const RowCells& cells, bool row_has_phantoms) {
  RowEmitPlan plan;
  const std::vector<StyleRun>& runs = cells.runs();
  if (runs.empty()) {
    return plan;
  }
  const SheetLayout& layout = sheet.layout();
  if (converts_to_row_format(layout, row, cells, row_has_phantoms)) {
    plan.as_row_format = true;
    return plan;
  }
  plan.runs.reserve(runs.size());
  const auto keep = [&plan](std::uint64_t first, std::uint64_t end, std::uint32_t xf) {
    if (first >= end) {
      return;
    }
    if (!plan.runs.empty() && plan.runs.back().end == first && plan.runs.back().xf == xf) {
      plan.runs.back().end = end;
      return;
    }
    plan.runs.push_back({static_cast<std::uint32_t>(first), end, xf});
  };
  for (const StyleRun& run : runs) {
    const std::uint64_t end = static_cast<std::uint64_t>(run.last) + 1U;
    std::uint64_t at = run.first;
    // A piece is written only where its xf differs from the inherited one;
    // columns no axis piece covers inherit xf 0.
    for (const AxisSegment& seg : axis_segments(layout, row, run.first, end)) {
      if (run.xf != 0U) {
        keep(at, seg.first, run.xf);
      }
      if (run.xf != seg.xf) {
        keep(seg.first, seg.end, run.xf);
      }
      at = seg.end;
    }
    if (run.xf != 0U) {
      keep(at, end, run.xf);
    }
  }
  return plan;
}

EffectiveXf select_axis_xf(const SheetLayout& layout, std::uint32_t row, std::uint32_t col) {
  if (const RowLayout* row_entry = find_row_override(layout, row); row_entry != nullptr && row_entry->has_style) {
    return {row_entry->style_xf, StyleSource::kRow};
  }
  const auto col_it = std::find_if(layout.columns.begin(), layout.columns.end(), [col](const ColumnLayout& c) {
    return c.first <= col && col <= c.last && c.has_style;
  });
  if (col_it != layout.columns.end()) {
    return {col_it->style_xf, StyleSource::kColumn};
  }
  return {0U, StyleSource::kDefault};
}

EffectiveStyle effective_style(const Workbook& wb, const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  const StylesTable& styles = wb.styles();
  const EffectiveXf choice = select_effective_xf(sheet, row, col);
  EffectiveStyle out;
  out.source = choice.source;
  out.xf_index = choice.xf_index < styles.cell_xfs.size() ? choice.xf_index : 0U;
  const CellXf xf = styles.cell_xfs.empty() ? CellXf{} : styles.cell_xfs[out.xf_index];

  std::optional<LoadedTheme> theme;
  const auto resolve = [&](const ColorSpec& spec, ColorContext context) {
    if (spec.kind != ColorSpec::Kind::kTheme) {
      return resolve_color(spec, default_theme(), ThemeSource::kDefault, styles.indexed_colors, context);
    }
    if (!theme) {
      theme.emplace(wb.load_theme());
    }
    return resolve_color(spec, theme->theme, theme->source, styles.indexed_colors, context);
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
  const std::optional<std::string_view> code = effective_num_fmt(styles, xf.num_fmt_id, wb.excel_profile());
  out.num_fmt_code = code ? std::string(*code) : std::string("General");
  if (xf.has_protection) {
    out.locked = xf.locked;
    out.hidden = xf.hidden;
  }
  return out;
}

}  // namespace formulon
