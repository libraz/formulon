//
// C ABI - PivotTable layout projection.
//
// Walks the cached `PivotTable` evaluator output and projects it onto a
// flat `fm_pivot_cells_t` buffer for the bindings to render. The
// opaque results handle owns its text store so it can hand out stable
// `c_str()` pointers for field names / number formats.

#include "pivot/pivot_layout.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <new>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/pivot_locale.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::value_to_fm;

struct fm_pivot_cells {
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  std::vector<fm_pivot_cell_t> cells;
  TextStore text_store;
};

namespace {

fm_pivot_cell_kind_t pivot_cell_kind_to_fm(formulon::pivot::PivotCellKind kind) {
  switch (kind) {
    case formulon::pivot::PivotCellKind::Header:
      return FM_PIVOT_CELL_HEADER;
    case formulon::pivot::PivotCellKind::RowLabel:
      return FM_PIVOT_CELL_ROW_LABEL;
    case formulon::pivot::PivotCellKind::ColLabel:
      return FM_PIVOT_CELL_COL_LABEL;
    case formulon::pivot::PivotCellKind::Data:
      return FM_PIVOT_CELL_DATA;
    case formulon::pivot::PivotCellKind::RowSubtotal:
      return FM_PIVOT_CELL_ROW_SUBTOTAL;
    case formulon::pivot::PivotCellKind::ColSubtotal:
      return FM_PIVOT_CELL_COL_SUBTOTAL;
    case formulon::pivot::PivotCellKind::GrandTotal:
      return FM_PIVOT_CELL_GRAND_TOTAL;
    case formulon::pivot::PivotCellKind::Blank:
      return FM_PIVOT_CELL_BLANK;
  }
  return FM_PIVOT_CELL_BLANK;
}

const char* store_cstr(TextStore& store, std::string_view text) {
  store.emplace_back(text.data(), text.size());
  return store.back().c_str();
}

}  // namespace

// `fm_workbook_pivot_count` is now emitted by the binding codegen (see
// `src/c_api/generated/sheet_counts.cpp`).

extern "C" fm_status_t fm_workbook_pivot_layout(const fm_workbook_t* wb, std::size_t sheet_index,
                                                std::size_t pivot_index, fm_pivot_cells_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_layout: NULL argument");
  }
  *out = nullptr;
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_pivot_layout"); rc != 0) {
    return rc;
  }
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  if (pivot_index >= sheet.pivot_tables().size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_pivot_layout: pivot_index out of range",
        "pivot_index=" + std::to_string(pivot_index) + " count=" + std::to_string(sheet.pivot_tables().size()));
  }

  const formulon::pivot::PivotTable* table = sheet.pivot_tables()[pivot_index].get();
  if (table == nullptr) {
    return set_binding_error(
        formulon::FormulonErrorCode::kEvalPivotInvalid, "fm_workbook_pivot_layout: pivot table entry is NULL",
        "sheet_index=" + std::to_string(sheet_index) + " pivot_index=" + std::to_string(pivot_index));
  }

  // Resolved before evaluation, not just before projection: the label of an
  // axis group with no source value is baked into the result rather than
  // synthesised by the projection.
  const formulon::pivot::PivotLayoutOptions layout_options =
      formulon::eval::pivot_layout_options_for(wb->workbook().excel_profile());

  std::shared_ptr<const formulon::pivot::PivotResult> result = table->last_result();
  if (!result) {
    const formulon::pivot::PivotCache* cache = wb->workbook().find_pivot_cache(table->pivot_cache_id());
    if (cache == nullptr) {
      return set_binding_error(formulon::FormulonErrorCode::kEvalPivotMissing,
                               "fm_workbook_pivot_layout: pivot cache not found",
                               "cache_id=" + std::to_string(table->pivot_cache_id()));
    }
    auto eval_or = formulon::pivot::evaluate(*table, *cache, layout_options);
    if (!eval_or) {
      return set_last_error(eval_or.error());
    }
    table->set_last_result(std::move(eval_or.value()));
    result = table->last_result();
  }
  if (!result) {
    return set_binding_error(formulon::FormulonErrorCode::kEvalPivotMissing,
                             "pivot result cache was invalidated during layout projection");
  }

  auto layout_or = formulon::pivot::layout(*table, *result, layout_options);
  if (!layout_or) {
    return set_last_error(layout_or.error());
  }

  auto handle = std::unique_ptr<fm_pivot_cells_t>(new fm_pivot_cells_t{});
  const formulon::pivot::PivotCells& projected = layout_or.value();
  handle->top = projected.top;
  handle->left = projected.left;
  handle->rows = projected.rows;
  handle->cols = projected.cols;
  handle->cells.reserve(projected.cells.size());
  for (const formulon::pivot::PivotCell& src : projected.cells) {
    fm_pivot_cell_t dst{};
    dst.row = src.row;
    dst.col = src.col;
    value_to_fm(src.value, handle->text_store, &dst.value);
    dst.kind = pivot_cell_kind_to_fm(src.kind);
    dst.depth = src.depth;
    dst.field_name = store_cstr(handle->text_store, src.field_name);
    dst.number_format = store_cstr(handle->text_store, src.number_format);
    handle->cells.push_back(dst);
  }

  *out = handle.release();
  return 0;
}

extern "C" void fm_pivot_cells_destroy(fm_pivot_cells_t* cells) {
  delete cells;
}

extern "C" std::size_t fm_pivot_cells_count(const fm_pivot_cells_t* cells) {
  if (cells == nullptr) {
    return 0;
  }
  return cells->cells.size();
}

extern "C" fm_status_t fm_pivot_cells_bounds(const fm_pivot_cells_t* cells, std::uint32_t* out_top,
                                             std::uint32_t* out_left, std::uint32_t* out_rows,
                                             std::uint32_t* out_cols) {
  clear_last_error();
  if (cells == nullptr || out_top == nullptr || out_left == nullptr || out_rows == nullptr || out_cols == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_pivot_cells_bounds: NULL argument");
  }
  *out_top = cells->top;
  *out_left = cells->left;
  *out_rows = cells->rows;
  *out_cols = cells->cols;
  return 0;
}

extern "C" fm_status_t fm_pivot_cells_at(const fm_pivot_cells_t* cells, std::size_t idx, fm_pivot_cell_t* out) {
  clear_last_error();
  if (cells == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_pivot_cells_at: NULL argument");
  }
  if (idx >= cells->cells.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_pivot_cells_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(cells->cells.size()));
  }
  *out = cells->cells[idx];
  return 0;
}
