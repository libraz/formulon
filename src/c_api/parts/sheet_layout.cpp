//
// C ABI - sheet layout (columns, rows, range styles, format defaults).
//
// Bridges `Sheet::layout()` over the stable C ABI.
// The `set_*` mutators upsert into the underlying `SheetLayout` vectors
// via small private helpers so concurrent column-span edits stay
// internally consistent (e.g. setting the width on a span that
// overlaps an existing entry replaces only the overlapping portion).

#include <algorithm>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "column_width.h"
#include "print/sheet_geometry.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "style_region.h"
#include "utils/error.h"
#include "utils/number_text.h"
#include "workbook.h"

using formulon::drop_stateless_columns;
using formulon::overlay_column_span;
using formulon::c_api::parts::check_cell_xf_index;
using formulon::c_api::parts::check_column_span;
using formulon::c_api::parts::check_finite_non_negative;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_limit;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_row_span;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::copy_owned_bytes;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::kMaxOutlineLevel;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::set_last_error;

namespace {

// Excel's widest column, in stored character units, and its tallest row, in points.
constexpr double kMaxColumnWidthChars = 255.0;
constexpr double kMaxRowHeightPt = 409.0;

// A finite, non-negative `value` no larger than `max`.
fm_status_t check_bounded(double value, double max, const char* field) {
  if (auto rc = check_finite_non_negative(value, field); rc != 0) {
    return rc;
  }
  if (value > max) {
    return invalid_argument("value is above the limit", std::string(field) + "=" + formulon::fixed_string(value, 6));
  }
  return 0;
}

fm_status_t check_sheet_column_span(const fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return check_column_span(first, last);
}

fm_status_t check_sheet_row_span(const fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return check_row_span(first, last);
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_column_layout_count(const fm_workbook_t* wb, size_t sheet_index,
                                                        size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_count = wb->workbook().sheet(sheet_index).layout().columns.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_column_layout_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                     fm_column_layout_info_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& cols = wb->workbook().sheet(sheet_index).layout().columns;
  if (auto rc = check_index(idx, cols.size(), "index"); rc != 0) {
    return rc;
  }
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_column_layout_info_t{};
  out->struct_size = struct_size;
  out->first = cols[idx].first;
  out->last = cols[idx].last;
  out->width_chars = cols[idx].width;
  out->hidden = cols[idx].hidden ? 1 : 0;
  out->outline_level = cols[idx].outline_level;
  out->collapsed = cols[idx].collapsed ? 1 : 0;
  out->has_width_chars = formulon::HasExplicitColumnWidth(cols[idx]) ? 1 : 0;
  out->has_xf_index = cols[idx].has_style ? 1 : 0;
  out->xf_index = cols[idx].style_xf;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_layout_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_count = wb->workbook().sheet(sheet_index).layout().row_overrides.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_layout_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                  fm_row_layout_info_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& rows = wb->workbook().sheet(sheet_index).layout().row_overrides;
  if (auto rc = check_index(idx, rows.size(), "index"); rc != 0) {
    return rc;
  }
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_row_layout_info_t{};
  out->struct_size = struct_size;
  out->row = rows[idx].row;
  out->height_pt = rows[idx].height;
  out->hidden = rows[idx].hidden ? 1 : 0;
  out->outline_level = rows[idx].outline_level;
  out->collapsed = rows[idx].collapsed ? 1 : 0;
  out->has_xf_index = rows[idx].has_style ? 1 : 0;
  out->xf_index = rows[idx].style_xf;
  out->has_height_pt = rows[idx].has_height ? 1 : 0;
  out->custom_height = rows[idx].custom_height ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                 double width) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  if (auto rc = check_bounded(width, kMaxColumnWidthChars, "width_chars"); rc != 0) {
    return rc;
  }
  overlay_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last,
                      [width](formulon::ColumnLayout& entry) {
                        entry.width = width;
                        entry.has_width = true;
                      });
  return 0;
}

extern "C" fm_status_t fm_sheet_reset_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first,
                                                   uint32_t last) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  formulon::SheetLayout& layout = wb->workbook().sheet(sheet_index).mutable_layout();
  overlay_column_span(layout, first, last, [](formulon::ColumnLayout& entry) {
    entry.width = 0.0;
    entry.has_width = false;
  });
  drop_stateless_columns(layout);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                  int32_t hidden) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  overlay_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last,
                      [hidden](formulon::ColumnLayout& entry) { entry.hidden = (hidden != 0); });
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                   uint32_t level) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  if (auto rc = check_limit(level, kMaxOutlineLevel + 1U, "level"); rc != 0) {
    return rc;
  }
  const auto outline_level = static_cast<std::uint8_t>(level);
  overlay_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last,
                      [outline_level](formulon::ColumnLayout& entry) { entry.outline_level = outline_level; });
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_collapsed(fm_workbook_t* wb, size_t sheet_index, uint32_t first,
                                                     uint32_t last, int32_t collapsed) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  formulon::SheetLayout& layout = wb->workbook().sheet(sheet_index).mutable_layout();
  overlay_column_span(layout, first, last,
                      [collapsed](formulon::ColumnLayout& entry) { entry.collapsed = (collapsed != 0); });
  if (collapsed == 0) {
    drop_stateless_columns(layout);
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t first,
                                                    uint32_t last, uint32_t xf_index) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_xf_index(wb, xf_index); rc != 0) {
    return rc;
  }
  formulon::set_column_style_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last, true, xf_index);
  return 0;
}

extern "C" fm_status_t fm_sheet_reset_column_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t first,
                                                      uint32_t last) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  formulon::set_column_style_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last, false, 0U);
  return 0;
}

namespace {

// Applies `clear` to the existing overrides of `[first, last]` and drops each
// one once it carries nothing else. Rows without an override are left alone.
void clear_row_flags(formulon::Sheet& sheet, uint32_t first, uint32_t last, void (*clear)(formulon::RowLayout&)) {
  formulon::SheetLayout& layout = sheet.mutable_layout();
  const bool default_hidden = sheet.format_defaults().zero_height;
  const auto begin = formulon::row_override_lower_bound(layout, first);
  const auto end = formulon::row_override_lower_bound(layout, last + 1U);
  auto kept = begin;
  for (auto it = begin; it != end; ++it) {
    clear(*it);
    if (!formulon::should_erase_row_override(*it, default_hidden, true)) {
      *kept++ = *it;
    }
  }
  layout.row_overrides.erase(kept, end);
}

// Applies `apply(entry, arg)` to the override of every row in `[first, last]`,
// creating the missing ones with the sheet's default visibility. The vector is
// rebuilt once, so a span costs O(overrides + span), not one insertion per row.
void overlay_row_span(formulon::Sheet& sheet, uint32_t first, uint32_t last, void (*apply)(formulon::RowLayout&, void*),
                      void* arg) {
  const bool default_hidden = sheet.format_defaults().zero_height;
  formulon::SheetLayout& layout = sheet.mutable_layout();
  std::vector<formulon::RowLayout>& rows = layout.row_overrides;
  const auto begin = formulon::row_override_lower_bound(layout, first);
  const auto end = formulon::row_override_lower_bound(layout, last + 1U);
  std::vector<formulon::RowLayout> merged;
  merged.reserve(rows.size() - static_cast<size_t>(end - begin) + (static_cast<size_t>(last - first) + 1U));
  merged.insert(merged.end(), rows.begin(), begin);
  auto it = begin;
  for (uint32_t row = first;; ++row) {
    if (it != end && it->row == row) {
      merged.push_back(*it);
      ++it;
    } else {
      formulon::RowLayout fresh;
      fresh.row = row;
      fresh.hidden = default_hidden;
      merged.push_back(fresh);
    }
    apply(merged.back(), arg);
    if (row == last) {
      break;
    }
  }
  merged.insert(merged.end(), end, rows.end());
  rows = std::move(merged);
}

struct RowHeightEdit {
  double height;
  bool custom;
};

// Stores `height` on every row of `[first, last]`, as a manual override when
// `custom` and as an automatic height's cached value otherwise; every other
// row attribute is kept.
fm_status_t store_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last, double height,
                             bool custom) {
  clear_last_error();
  if (auto rc = check_sheet_row_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  if (auto rc = check_bounded(height, kMaxRowHeightPt, "height_pt"); rc != 0) {
    return rc;
  }
  RowHeightEdit edit{height, custom};
  overlay_row_span(
      wb->workbook().sheet(sheet_index), first, last,
      [](formulon::RowLayout& entry, void* arg) {
        const auto* e = static_cast<const RowHeightEdit*>(arg);
        entry.height = e->height;
        entry.has_height = true;
        entry.custom_height = e->custom;
      },
      &edit);
  return 0;
}

struct RowHiddenEdit {
  bool hidden;
  bool changed;
};

}  // namespace

extern "C" fm_status_t fm_sheet_set_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                               double height) {
  return store_row_height(wb, sheet_index, first, last, height, true);
}

extern "C" fm_status_t fm_sheet_set_row_auto_height(fm_workbook_t* wb, size_t sheet_index, uint32_t first,
                                                    uint32_t last, double height) {
  return store_row_height(wb, sheet_index, first, last, height, false);
}

extern "C" fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                               int32_t hidden) {
  clear_last_error();
  if (auto rc = check_sheet_row_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  RowHiddenEdit edit{hidden != 0, false};
  overlay_row_span(
      wb->workbook().sheet(sheet_index), first, last,
      [](formulon::RowLayout& entry, void* arg) {
        auto* e = static_cast<RowHiddenEdit*>(arg);
        e->changed = e->changed || entry.hidden != e->hidden;
        entry.hidden = e->hidden;
      },
      &edit);
  // SUBTOTAL(101..111, ...) and AGGREGATE's hidden-row-skipping option read
  // this row's hidden state without that state appearing anywhere in the
  // dep graph's cell/range edges, so the ordinary dirty propagation a
  // cell-value write triggers cannot reach them.
  if (edit.changed) {
    wb->workbook().mark_row_visibility_dependents_dirty();
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                uint32_t level) {
  clear_last_error();
  if (auto rc = check_sheet_row_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  if (auto rc = check_limit(level, kMaxOutlineLevel + 1U, "level"); rc != 0) {
    return rc;
  }
  std::uint8_t outline_level = static_cast<std::uint8_t>(level);
  overlay_row_span(
      wb->workbook().sheet(sheet_index), first, last,
      [](formulon::RowLayout& entry, void* arg) { entry.outline_level = *static_cast<const std::uint8_t*>(arg); },
      &outline_level);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_collapsed(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                  int32_t collapsed) {
  clear_last_error();
  if (auto rc = check_sheet_row_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  if (collapsed != 0) {
    overlay_row_span(sheet, first, last, [](formulon::RowLayout& entry, void*) { entry.collapsed = true; }, nullptr);
    return 0;
  }
  clear_row_flags(sheet, first, last, [](formulon::RowLayout& entry) { entry.collapsed = false; });
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                 uint32_t xf_index) {
  clear_last_error();
  if (auto rc = check_sheet_row_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_xf_index(wb, xf_index); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  formulon::set_row_style_span(sheet.mutable_layout(), first, last, xf_index, sheet.format_defaults().zero_height);
  return 0;
}

extern "C" fm_status_t fm_sheet_reset_row_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t first,
                                                   uint32_t last) {
  clear_last_error();
  if (auto rc = check_sheet_row_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  formulon::clear_row_style_span(sheet.mutable_layout(), first, last, sheet.format_defaults().zero_height);
  return 0;
}

namespace {

// Validates a selection rect and copies it into the engine's form.
fm_status_t resolve_style_rect(const fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                               formulon::StyleRect* out) {
  if (range == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(*range); rc != 0) {
    return rc;
  }
  out->first_row = range->first_row;
  out->last_row = range->last_row;
  out->first_col = range->first_col;
  out->last_col = range->last_col;
  return 0;
}

fm_status_t range_style_bases_into(const fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                   const fm_cell_address_t* excluded, size_t excluded_count, uint32_t* out,
                                   size_t capacity, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr || (out == nullptr && capacity > 0U) || (excluded_count > 0U && excluded == nullptr)) {
    return null_argument();
  }
  *out_count = 0U;
  formulon::StyleRect rect;
  if (auto rc = resolve_style_rect(wb, sheet_index, range, &rect); rc != 0) {
    return rc;
  }
  std::vector<formulon::StyleCell> skipped(excluded_count);
  for (size_t i = 0; i < excluded_count; ++i) {
    skipped[i] = {excluded[i].row, excluded[i].col};
  }
  auto bases = wb->workbook().range_style_bases(sheet_index, rect, std::move(skipped));
  if (!bases) {
    return set_last_error(bases.error());
  }
  *out_count = bases.value().size();
  std::copy_n(bases.value().begin(), std::min(capacity, bases.value().size()), out);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_range_xf_bases(const fm_workbook_t* wb, size_t sheet_index,
                                                   const fm_cell_rect_t* range, uint32_t* out, size_t capacity,
                                                   size_t* out_count) {
  return range_style_bases_into(wb, sheet_index, range, nullptr, 0U, out, capacity, out_count);
}

extern "C" fm_status_t fm_sheet_get_range_xf_bases_excluding(const fm_workbook_t* wb, size_t sheet_index,
                                                             const fm_cell_rect_t* range,
                                                             const fm_cell_address_t* excluded, size_t excluded_count,
                                                             uint32_t* out, size_t capacity, size_t* out_count) {
  return range_style_bases_into(wb, sheet_index, range, excluded, excluded_count, out, capacity, out_count);
}

extern "C" fm_status_t fm_sheet_get_cell_xf_indices(const fm_workbook_t* wb, size_t sheet_index,
                                                    const fm_cell_rect_t* range, fm_cell_xf_index_t* out,
                                                    size_t capacity, size_t* out_count) {
  clear_last_error();
  if (out_count != nullptr) {
    *out_count = 0U;
  }
  if (range == nullptr || out_count == nullptr || (out == nullptr && capacity > 0U)) {
    return null_argument();
  }
  formulon::StyleRect rect{};
  if (auto rc = resolve_style_rect(wb, sheet_index, range, &rect); rc != 0) {
    return rc;
  }
  auto cells = wb->workbook().cell_xf_indices(sheet_index, rect);
  if (!cells) {
    return set_last_error(cells.error());
  }
  const size_t total = cells.value().size();
  *out_count = total;
  for (size_t i = 0; i < std::min(capacity, total); ++i) {
    const formulon::StyleCellXf& cell = cells.value()[i];
    out[i] = {cell.row, cell.col, cell.xf};
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_remap_range_xfs(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                                const fm_xf_remap_t* pairs, size_t count) {
  clear_last_error();
  if (count > 0U && pairs == nullptr) {
    return null_argument();
  }
  formulon::StyleRect rect;
  if (auto rc = resolve_style_rect(wb, sheet_index, range, &rect); rc != 0) {
    return rc;
  }
  std::vector<formulon::XfRemap> map(count);
  for (size_t i = 0; i < count; ++i) {
    map[i] = {pairs[i].from_xf, pairs[i].to_xf};
  }
  return finish_status(wb->workbook().remap_range_style(sheet_index, rect, std::move(map)));
}

extern "C" fm_status_t fm_sheet_snapshot_range_xfs(const fm_workbook_t* wb, size_t sheet_index,
                                                   const fm_cell_rect_t* range, uint8_t** out_bytes, size_t* out_len) {
  clear_last_error();
  if (out_bytes != nullptr) {
    *out_bytes = nullptr;
  }
  if (out_len != nullptr) {
    *out_len = 0U;
  }
  if (out_bytes == nullptr || out_len == nullptr) {
    return null_argument();
  }
  formulon::StyleRect rect;
  if (auto rc = resolve_style_rect(wb, sheet_index, range, &rect); rc != 0) {
    return rc;
  }
  auto bytes = wb->workbook().snapshot_range_style(sheet_index, rect);
  if (!bytes) {
    return set_last_error(bytes.error());
  }
  copy_owned_bytes(bytes.value(), out_bytes, out_len);
  return 0;
}

extern "C" fm_status_t fm_sheet_restore_range_xfs(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes,
                                                  size_t len) {
  clear_last_error();
  if (bytes == nullptr && len > 0U) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().restore_range_style(sheet_index, bytes, len));
}

extern "C" fm_status_t fm_sheet_reset_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last) {
  clear_last_error();
  if (auto rc = check_sheet_row_span(wb, sheet_index, first, last); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  clear_row_flags(sheet, first, last, [](formulon::RowLayout& entry) {
    entry.height = 0.0;
    entry.has_height = false;
    entry.custom_height = false;
  });
  return 0;
}

extern "C" fm_status_t fm_sheet_get_format_defaults(const fm_workbook_t* wb, size_t sheet_index,
                                                    fm_sheet_format_defaults_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const formulon::SheetFormatDefaults& defaults = wb->workbook().sheet(sheet_index).format_defaults();
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_sheet_format_defaults_t{};
  out->struct_size = struct_size;
  out->default_column_width_chars = defaults.default_col_width;
  out->default_row_height_pt = defaults.default_row_height;
  out->base_column_width_chars = defaults.base_col_width;
  out->has_default_column_width_chars = defaults.has_default_col_width ? 1 : 0;
  out->has_default_row_height_pt = defaults.has_default_row_height ? 1 : 0;
  out->custom_height = defaults.custom_height ? 1 : 0;
  out->zero_height = defaults.zero_height ? 1 : 0;
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  out->effective_default_column_width_chars = formulon::default_column_width_chars(
      sheet, formulon::resolve_column_width_model(wb->workbook().styles(), formulon::GeometryMode::kDisplay));
  out->effective_default_row_height_pt = defaults.zero_height ? 0.0 : formulon::print::default_row_height_pt(sheet);
  return 0;
}

extern "C" fm_status_t fm_sheet_update_format_defaults(fm_workbook_t* wb, size_t sheet_index,
                                                       const fm_sheet_format_defaults_patch_t* patch) {
  clear_last_error();
  if (patch == nullptr) {
    return null_argument("patch");
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (!patch_op_ok(patch->default_column_width_chars_op) || !patch_op_ok(patch->default_row_height_pt_op) ||
      !patch_op_ok_without_clear(patch->base_column_width_chars_op) ||
      !patch_op_ok_without_clear(patch->custom_height_op) || !patch_op_ok_without_clear(patch->zero_height_op)) {
    return invalid_argument("format defaults patch op out of range or not clearable");
  }
  struct NumberField {
    fm_patch_op_t op;
    double value;
    const char* field;
  };
  const NumberField numbers[] = {
      {patch->default_column_width_chars_op, patch->default_column_width_chars, "default_column_width_chars"},
      {patch->default_row_height_pt_op, patch->default_row_height_pt, "default_row_height_pt"},
      {patch->base_column_width_chars_op, patch->base_column_width_chars, "base_column_width_chars"},
  };
  for (const NumberField& number : numbers) {
    if (number.op == FM_PATCH_OP_SET) {
      if (auto rc = check_finite_non_negative(number.value, number.field); rc != 0) {
        return rc;
      }
    }
  }
  formulon::SheetFormatDefaults& target = wb->workbook().sheet(sheet_index).mutable_format_defaults();
  if (patch->default_column_width_chars_op != FM_PATCH_OP_KEEP) {
    const bool set = patch->default_column_width_chars_op == FM_PATCH_OP_SET;
    target.default_col_width = set ? patch->default_column_width_chars : 0.0;
    target.has_default_col_width = set;
  }
  if (patch->default_row_height_pt_op != FM_PATCH_OP_KEEP) {
    const bool set = patch->default_row_height_pt_op == FM_PATCH_OP_SET;
    target.default_row_height = set ? patch->default_row_height_pt : 0.0;
    target.has_default_row_height = set;
  }
  if (patch->base_column_width_chars_op == FM_PATCH_OP_SET) {
    target.base_col_width = patch->base_column_width_chars;
  }
  if (patch->custom_height_op == FM_PATCH_OP_SET) {
    target.custom_height = patch->custom_height != 0;
  }
  if (patch->zero_height_op == FM_PATCH_OP_SET) {
    // Hidden-row state (SUBTOTAL / AGGREGATE) depends on the default.
    const bool zero_height = patch->zero_height != 0;
    const bool zero_height_changed = target.zero_height != zero_height;
    target.zero_height = zero_height;
    if (zero_height_changed) {
      wb->workbook().mark_row_visibility_dependents_dirty();
    }
  }
  return 0;
}
