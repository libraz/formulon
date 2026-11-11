// Workbook-level hyperlink edits: the list change plus the cell formats Excel
// writes with it. Adding a link gives its cells the built-in Hyperlink cell
// style; removing one resets its whole range to cell format 0. File readers
// write `Sheet::mutable_hyperlinks()` directly, so a load stays style-neutral.

#include <cstddef>
#include <cstdint>
#include <mutex>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "column_width.h"
#include "eval/coerce.h"
#include "eval/recalc_engine.h"
#include "new_workbook_defaults.h"
#include "sheet.h"
#include "style_region.h"
#include "styles.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/utf8_length.h"
#include "value.h"
#include "workbook.h"
#include "workbook_sheet_mutation.h"

namespace formulon {

namespace {

constexpr std::uint32_t kHyperlinkBuiltinId = 8U;
constexpr std::uint32_t kNormalBuiltinId = 0U;
/// `FontRecord::underline` ordinal of a single underline (`<u/>`).
constexpr std::uint8_t kSingleUnderline = 1U;
/// Theme colour slot of a hyperlink (`hlink`).
constexpr std::uint32_t kHyperlinkThemeColor = 10U;

StyleRect rect_of(const Hyperlink& link) {
  return StyleRect{link.row, link.col, link.last_row, link.last_col};
}

/// Refuses, before anything changes, a range `remap_range_style` would refuse
/// for its size: more than `kMaxStyledRangeCells` cells that are not a whole
/// column or row.
Expected<void, Error> check_restyle_area(const StyleRect& rect) {
  const bool all_rows = rect.first_row == 0U && rect.last_row == Sheet::kMaxRows - 1U;
  const bool all_cols = rect.first_col == 0U && rect.last_col == Sheet::kMaxCols - 1U;
  const std::uint64_t area = (static_cast<std::uint64_t>(rect.last_row - rect.first_row) + 1U) *
                             (static_cast<std::uint64_t>(rect.last_col - rect.first_col) + 1U);
  if (!all_rows && !all_cols && area > kMaxStyledRangeCells) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "rect covers too many cells to format",
                      "cells=" + std::to_string(area));
  }
  return Expected<void, Error>::Ok();
}

/// The `cell_style_xfs` index of the Normal style, by builtin id.
std::uint32_t normal_style_xf(const StylesTable& styles) {
  for (const CellStyleRecord& style : styles.cell_styles) {
    if (style.builtin_id == kNormalBuiltinId && style.xf_id < styles.cell_style_xfs.size()) {
      return style.xf_id;
    }
  }
  return 0U;
}

/// The `cell_style_xfs` index of the built-in Hyperlink style. An existing
/// record is found by builtin id, whatever its name; otherwise the style is
/// added as a clone of Normal whose font is the Normal font underlined in the
/// hyperlink theme colour. The font is appended, never inserted, so font ids
/// already referenced verbatim (`phoneticPr`) keep their meaning.
std::uint32_t hyperlink_style_xf(StylesTable& styles, std::string_view name) {
  for (const CellStyleRecord& style : styles.cell_styles) {
    if (style.builtin_id == kHyperlinkBuiltinId && style.xf_id < styles.cell_style_xfs.size()) {
      return style.xf_id;
    }
  }
  if (styles.cell_style_xfs.empty()) {
    styles.cell_style_xfs.emplace_back();
  }
  const FontRecord* normal_font = resolve_normal_font(styles);
  FontRecord font = normal_font != nullptr ? *normal_font : FontRecord{};
  font.underline = kSingleUnderline;
  font.color = ColorSpec{};
  font.color.kind = ColorSpec::Kind::kTheme;
  font.color.theme = kHyperlinkThemeColor;
  styles.fonts.push_back(std::move(font));

  CellXf xf = styles.cell_style_xfs[normal_style_xf(styles)];
  xf.font_index = static_cast<std::uint32_t>(styles.fonts.size() - 1U);
  xf.xf_id = 0U;
  xf.apply_number_format = false;
  xf.apply_font = false;
  xf.apply_fill = false;
  xf.apply_border = false;
  xf.apply_alignment = false;
  xf.apply_protection = false;
  styles.cell_style_xfs.push_back(xf);
  const auto index = static_cast<std::uint32_t>(styles.cell_style_xfs.size() - 1U);

  CellStyleRecord record;
  record.name = std::string(name);
  record.xf_id = index;
  record.builtin_id = kHyperlinkBuiltinId;
  styles.cell_styles.push_back(std::move(record));
  return index;
}

/// `base` under the Hyperlink style: its font and parent style replaced, every
/// other part kept. A number format, fill or border that differs from the
/// style's is marked applied, as Excel marks it.
CellXf hyperlink_cell_xf(const StylesTable& styles, std::uint32_t base, std::uint32_t style_index) {
  CellXf xf = base < styles.cell_xfs.size() ? styles.cell_xfs[base] : CellXf{};
  const CellXf& style = styles.cell_style_xfs[style_index];
  xf.font_index = style.font_index;
  xf.xf_id = style_index;
  xf.apply_font = false;
  xf.apply_number_format = xf.apply_number_format || xf.num_fmt_id != style.num_fmt_id;
  xf.apply_fill = xf.apply_fill || xf.fill_index != style.fill_index;
  xf.apply_border = xf.apply_border || xf.border_index != style.border_index;
  return xf;
}

/// Whether OM Add would write the link's text into `(row, col)`: the cell
/// shows nothing (no value, formula or spilled value) and is not a cell a
/// merge hides.
bool takes_link_text(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = sheet.cell_at(row, col);
  if (cell != nullptr && (!cell->formula_text.empty() || !cell->cached_value.is_blank())) {
    return false;
  }
  if (!sheet.resolve_cell_value(row, col).is_blank()) {
    return false;
  }
  for (const MergeRange& merge : sheet.merges()) {
    if (merge.first_row <= row && row <= merge.last_row && merge.first_col <= col && col <= merge.last_col) {
      return merge.first_row == row && merge.first_col == col;
    }
  }
  return true;
}

/// The text Excel writes into an empty top-left cell (display, else target,
/// else location) and the refusal naming that field when it is over the cap.
struct LinkText {
  std::string_view text;
  const char* over_cap;
};

LinkText link_text(const Hyperlink& link) {
  if (!link.display.empty()) {
    return {link.display, "display must be at most 32767 UTF-16 units"};
  }
  if (!link.target.empty()) {
    return {link.target, "target must be at most 32767 UTF-16 units"};
  }
  return {link.location, "location must be at most 32767 UTF-16 units"};
}

}  // namespace

Expected<void, Error> Workbook::set_hyperlink(std::size_t sheet_index, Hyperlink link) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, sheets_.size()));
  if (!Sheet::rect_in_grid(link.row, link.col, link.last_row, link.last_col) || link.row > link.last_row ||
      link.col > link.last_col) {
    return make_error(FormulonErrorCode::kInvalidArgument, "rect is inverted or outside the grid");
  }
  const LinkText chosen = link_text(link);
  const std::string_view text = chosen.text;
  if (text.size() > eval::kExcelTextCapUnits && utf16_units_in(text) > eval::kExcelTextCapUnits) {
    return make_error(FormulonErrorCode::kInvalidArgument, chosen.over_cap);
  }
  const StyleRect rect = rect_of(link);
  RETURN_IF_ERROR(check_restyle_area(rect));

  Sheet& sheet = sheets_[sheet_index];
  const bool write_text = !text.empty() && takes_link_text(sheet, link.row, link.col);
  if (write_text) {
    RETURN_IF_ERROR(set_cell_text(sheet_index, link.row, link.col, text));
  }

  const std::size_t font_count = styles_.fonts.size();
  const std::size_t style_xf_count = styles_.cell_style_xfs.size();
  const std::size_t style_count = styles_.cell_styles.size();
  const std::size_t xf_count = styles_.cell_xfs.size();
  const std::uint32_t style_index =
      hyperlink_style_xf(styles_, new_workbook_defaults(excel_profile_.locale).hyperlink_style_name);
  std::vector<XfRemap> map;
  for (const std::uint32_t base : formulon::range_style_bases(sheet, rect, {})) {
    map.push_back({base, intern_cell_xf(styles_, hyperlink_cell_xf(styles_, base, style_index))});
  }
  if (auto styled = remap_range_style(sheet_index, rect, std::move(map)); !styled) {
    styles_.fonts.resize(font_count);
    styles_.cell_style_xfs.resize(style_xf_count);
    styles_.cell_styles.resize(style_count);
    styles_.cell_xfs.resize(xf_count);
    if (write_text) {
      // The cell was empty before; put it back so a refusal changes nothing.
      (void)set_cell_value(sheet_index, link.row, link.col, Value::blank());
    }
    return styled;
  }

  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  sheet.set_hyperlink(std::move(link));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::reset_hyperlink_ranges(std::size_t sheet_index, const std::vector<StyleRect>& rects) {
  for (const StyleRect& rect : rects) {
    std::vector<XfRemap> map;
    for (const std::uint32_t base : formulon::range_style_bases(sheets_[sheet_index], rect, {})) {
      map.push_back({base, 0U});
    }
    RETURN_IF_ERROR(remap_range_style(sheet_index, rect, std::move(map)));
  }
  return Expected<void, Error>::Ok();
}

Expected<std::size_t, Error> Workbook::remove_hyperlinks_intersecting(std::size_t sheet_index, MergeRange range) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, sheets_.size()));
  if (!Sheet::rect_in_grid(range.first_row, range.first_col, range.last_row, range.last_col) ||
      range.first_row > range.last_row || range.first_col > range.last_col) {
    return make_error(FormulonErrorCode::kInvalidArgument, "range is inverted or outside the grid");
  }
  Sheet& sheet = sheets_[sheet_index];
  std::vector<StyleRect> resets;
  for (const Hyperlink& link : sheet.hyperlinks()) {
    if (link.intersects(range.first_row, range.first_col, range.last_row, range.last_col)) {
      RETURN_IF_ERROR(check_restyle_area(rect_of(link)));
      resets.push_back(rect_of(link));
    }
  }
  if (resets.empty()) {
    return std::size_t{0};
  }
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    sheet.remove_hyperlinks_intersecting(range);
  }
  RETURN_IF_ERROR(reset_hyperlink_ranges(sheet_index, resets));
  return resets.size();
}

Expected<void, Error> Workbook::clear_hyperlinks(std::size_t sheet_index) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, sheets_.size()));
  Sheet& sheet = sheets_[sheet_index];
  std::vector<StyleRect> resets;
  for (const Hyperlink& link : sheet.hyperlinks()) {
    RETURN_IF_ERROR(check_restyle_area(rect_of(link)));
    resets.push_back(rect_of(link));
  }
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    sheet.mutable_hyperlinks().clear();
  }
  return reset_hyperlink_ranges(sheet_index, resets);
}

}  // namespace formulon
