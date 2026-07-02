// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// C ABI - sheet view / layout (columns, rows, view, zoom, freeze).
//
// Bridges `Sheet::view()` / `Sheet::layout()` over the stable C ABI.
// The `set_*` mutators upsert into the underlying `SheetLayout` vectors
// via small private helpers so concurrent column-span edits stay
// internally consistent (e.g. setting the width on a span that
// overlaps an existing entry replaces only the overlapping portion).

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;

namespace {

// Splits any pre-existing column entries that intersect `[first, last]`
// so the resulting `columns` vector contains at most one entry whose
// span equals `[first, last]`. Pre-existing fields that fall outside
// the requested range are preserved on the residual entry. Returns a
// reference to the entry whose span is `[first, last]`; the caller
// then writes the field it wants to update on that entry.
formulon::ColumnLayout& upsert_column_span(formulon::SheetLayout& layout, std::uint32_t first, std::uint32_t last) {
  // First pass: split any entry that overlaps the target span. We
  // copy non-overlapping residuals into a fresh vector so the caller
  // sees a consistent state regardless of how many splits happened.
  std::vector<formulon::ColumnLayout> next;
  next.reserve(layout.columns.size() + 2);
  for (const formulon::ColumnLayout& entry : layout.columns) {
    // No overlap -> retain verbatim.
    if (entry.last < first || entry.first > last) {
      next.push_back(entry);
      continue;
    }
    // Left residual (entry.first .. first-1).
    if (entry.first < first) {
      formulon::ColumnLayout left = entry;
      left.last = first - 1U;
      next.push_back(left);
    }
    // Right residual (last+1 .. entry.last).
    if (entry.last > last) {
      formulon::ColumnLayout right = entry;
      right.first = last + 1U;
      next.push_back(right);
    }
    // The middle slice `[max(entry.first,first), min(entry.last,last)]`
    // is dropped here; we re-create one canonical entry below so the
    // caller can write its field of interest atomically.
  }
  formulon::ColumnLayout target;
  target.first = first;
  target.last = last;
  next.push_back(target);
  layout.columns = std::move(next);
  // The freshly inserted entry is the last element by construction.
  return layout.columns.back();
}

// Returns a pointer to the row override whose `row` equals `row`,
// creating one if none exists. The returned pointer is valid until
// the next mutation of `layout.row_overrides`.
formulon::RowLayout* upsert_row_override(formulon::SheetLayout& layout, std::uint32_t row) {
  for (formulon::RowLayout& entry : layout.row_overrides) {
    if (entry.row == row) {
      return &entry;
    }
  }
  formulon::RowLayout fresh;
  fresh.row = row;
  layout.row_overrides.push_back(fresh);
  return &layout.row_overrides.back();
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_column_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_column_count: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_column_count"); rc != 0) {
    return rc;
  }
  *out_count = wb->workbook().sheet(sheet_index).layout().columns.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_column(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                           fm_column_layout_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_column: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_column"); rc != 0) {
    return rc;
  }
  const auto& cols = wb->workbook().sheet(sheet_index).layout().columns;
  if (idx >= cols.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_get_column: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(cols.size()));
  }
  *out = fm_column_layout_t{};
  out->first = cols[idx].first;
  out->last = cols[idx].last;
  out->width = cols[idx].width;
  out->hidden = cols[idx].hidden ? 1 : 0;
  out->outline_level = cols[idx].outline_level;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_override_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_row_override_count: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_row_override_count"); rc != 0) {
    return rc;
  }
  *out_count = wb->workbook().sheet(sheet_index).layout().row_overrides.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_override(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                 fm_row_layout_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_row_override: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_row_override"); rc != 0) {
    return rc;
  }
  const auto& rows = wb->workbook().sheet(sheet_index).layout().row_overrides;
  if (idx >= rows.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_get_row_override: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(rows.size()));
  }
  *out = fm_row_layout_t{};
  out->row = rows[idx].row;
  out->height = rows[idx].height;
  out->hidden = rows[idx].hidden ? 1 : 0;
  out->outline_level = rows[idx].outline_level;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_view(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_view: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_view"); rc != 0) {
    return rc;
  }
  const formulon::SheetView& v = wb->workbook().sheet(sheet_index).view();
  out->zoom_scale = v.zoom_scale;
  out->freeze_rows = v.freeze_rows;
  out->freeze_cols = v.freeze_cols;
  out->tab_hidden = v.tab_hidden ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                 double width) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_column_width"); rc != 0) {
    return rc;
  }
  if (last < first) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_column_width: last < first",
                             "first=" + std::to_string(first) + " last=" + std::to_string(last));
  }
  formulon::ColumnLayout& entry = upsert_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last);
  entry.width = width;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                  int32_t hidden) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_column_hidden"); rc != 0) {
    return rc;
  }
  if (last < first) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_column_hidden: last < first",
                             "first=" + std::to_string(first) + " last=" + std::to_string(last));
  }
  formulon::ColumnLayout& entry = upsert_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last);
  entry.hidden = (hidden != 0);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                   uint8_t level) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_column_outline"); rc != 0) {
    return rc;
  }
  if (last < first) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_column_outline: last < first",
                             "first=" + std::to_string(first) + " last=" + std::to_string(last));
  }
  formulon::ColumnLayout& entry = upsert_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last);
  entry.outline_level = level;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row, double height) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_row_height"); rc != 0) {
    return rc;
  }
  formulon::RowLayout* entry = upsert_row_override(wb->workbook().sheet(sheet_index).mutable_layout(), row);
  entry->height = height;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t hidden) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_row_hidden"); rc != 0) {
    return rc;
  }
  formulon::RowLayout* entry = upsert_row_override(wb->workbook().sheet(sheet_index).mutable_layout(), row);
  entry->hidden = (hidden != 0);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint8_t level) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_row_outline"); rc != 0) {
    return rc;
  }
  formulon::RowLayout* entry = upsert_row_override(wb->workbook().sheet(sheet_index).mutable_layout(), row);
  entry->outline_level = level;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_zoom(fm_workbook_t* wb, size_t sheet_index, uint32_t zoom_scale) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_zoom"); rc != 0) {
    return rc;
  }
  // Clamp to the OOXML-valid `[10, 400]` interval; out-of-range values
  // are rounded to the nearest endpoint rather than rejected so JS
  // callers do not have to mirror the bound.
  std::uint32_t clamped = zoom_scale;
  if (clamped < 10U) {
    clamped = 10U;
  } else if (clamped > 400U) {
    clamped = 400U;
  }
  wb->workbook().sheet(sheet_index).mutable_view().zoom_scale = clamped;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_freeze(fm_workbook_t* wb, size_t sheet_index, uint32_t freeze_rows,
                                           uint32_t freeze_cols) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_freeze"); rc != 0) {
    return rc;
  }
  formulon::SheetView& view = wb->workbook().sheet(sheet_index).mutable_view();
  view.freeze_rows = freeze_rows;
  view.freeze_cols = freeze_cols;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_tab_hidden(fm_workbook_t* wb, size_t sheet_index, int32_t hidden) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_tab_hidden"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().tab_hidden = (hidden != 0);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_view_ex(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_ex_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_view_ex: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_view_ex"); rc != 0) {
    return rc;
  }
  const formulon::SheetView& v = wb->workbook().sheet(sheet_index).view();
  out->zoom_scale = v.zoom_scale;
  out->freeze_rows = v.freeze_rows;
  out->freeze_cols = v.freeze_cols;
  out->tab_hidden = v.tab_hidden ? 1 : 0;
  out->show_grid_lines = v.show_grid_lines ? 1 : 0;
  out->show_row_col_headers = v.show_row_col_headers ? 1 : 0;
  out->show_zeros = v.show_zeros ? 1 : 0;
  out->right_to_left = v.right_to_left ? 1 : 0;
  out->tab_selected = v.tab_selected ? 1 : 0;
  out->view_mode = v.view_mode.c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_set_show_grid_lines(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_show_grid_lines"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().show_grid_lines = (show != 0);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_show_row_col_headers(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_show_row_col_headers"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().show_row_col_headers = (show != 0);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_show_zeros(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_show_zeros"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().show_zeros = (show != 0);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_right_to_left(fm_workbook_t* wb, size_t sheet_index, int32_t right_to_left) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_right_to_left"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().right_to_left = (right_to_left != 0);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_tab_selected(fm_workbook_t* wb, size_t sheet_index, int32_t selected) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_tab_selected"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().tab_selected = (selected != 0);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_view_mode(fm_workbook_t* wb, size_t sheet_index, const char* mode) {
  clear_last_error();
  if (mode == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_set_view_mode: NULL mode");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_view_mode"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().view_mode = mode;
  return 0;
}
