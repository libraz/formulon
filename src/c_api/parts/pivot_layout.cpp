//
// C ABI - PivotTable layout projection.
//
// Walks the cached `PivotTable` evaluator output and projects it onto a
// flat `fm_pivot_report_t` buffer for the bindings to render. The
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
#include "c_api/parts/pivot_internal.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "eval/pivot_formula.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::cell_rect_from_extent;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::ensure_calculated;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::pivot_num_fmt_code;
using formulon::c_api::parts::resolve_pivot;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::value_to_fm;

struct fm_pivot_cells {
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  std::vector<fm_pivot_cell_info_t> cells;
  TextStore text_store;
};

namespace {

fm_pivot_cell_kind_t pivot_cell_kind_to_fm(formulon::pivot::PivotCellKind kind) {
  switch (kind) {
    case formulon::pivot::PivotCellKind::Header:
      return FM_PIVOT_CELL_KIND_HEADER;
    case formulon::pivot::PivotCellKind::RowLabel:
      return FM_PIVOT_CELL_KIND_ROW_LABEL;
    case formulon::pivot::PivotCellKind::ColLabel:
      return FM_PIVOT_CELL_KIND_COLUMN_LABEL;
    case formulon::pivot::PivotCellKind::Data:
      return FM_PIVOT_CELL_KIND_DATA;
    case formulon::pivot::PivotCellKind::RowSubtotal:
      return FM_PIVOT_CELL_KIND_ROW_SUBTOTAL;
    case formulon::pivot::PivotCellKind::ColSubtotal:
      return FM_PIVOT_CELL_KIND_COLUMN_SUBTOTAL;
    case formulon::pivot::PivotCellKind::GrandTotal:
      return FM_PIVOT_CELL_KIND_GRAND_TOTAL;
    case formulon::pivot::PivotCellKind::Blank:
      return FM_PIVOT_CELL_KIND_BLANK;
  }
  return FM_PIVOT_CELL_KIND_BLANK;
}

}  // namespace

// `fm_sheet_get_pivot_count` is now emitted by the binding codegen (see
// `src/c_api/generated/sheet_counts.cpp`).

extern "C" fm_status_t fm_sheet_get_pivot_layout(const fm_workbook_t* wb, std::size_t sheet_index,
                                                 std::size_t pivot_index, fm_pivot_report_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument();
  }
  *out = nullptr;
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }

  // Resolved before evaluation, not just before projection: the label of an
  // axis group with no source value is baked into the result rather than
  // synthesised by the projection.
  const formulon::pivot::PivotLayoutOptions layout_options =
      formulon::pivot::pivot_layout_options_for(wb->workbook().excel_profile());

  std::shared_ptr<const formulon::pivot::PivotResult> result = table->last_result();
  if (!result) {
    const formulon::pivot::PivotCache* cache = wb->workbook().find_pivot_cache(table->pivot_cache_id());
    if (cache == nullptr) {
      return set_binding_error(formulon::FormulonErrorCode::kEvalPivotMissing, "pivot cache not found",
                               "cache_id=" + std::to_string(table->pivot_cache_id()));
    }
    // Calculated fields and items are evaluated through the formula host.
    formulon::eval::PivotFormulaEvaluator formulas;
    auto eval_or = formulon::pivot::evaluate(
        *table, *cache, layout_options,
        formulon::pivot::PivotFilterEnv{wb->workbook().pinned_now(), wb->workbook().date1904(),
                                        &wb->workbook().styles(), &formulas});
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

  auto handle = std::unique_ptr<fm_pivot_report_t>(new fm_pivot_report_t{});
  const formulon::pivot::PivotCells& projected = layout_or.value();
  handle->top = projected.top;
  handle->left = projected.left;
  handle->rows = projected.rows;
  handle->cols = projected.cols;
  handle->cells.reserve(projected.cells.size());
  for (const formulon::pivot::PivotCell& src : projected.cells) {
    fm_pivot_cell_info_t dst = FM_STRUCT_INIT(fm_pivot_cell_info_t);
    dst.row = src.row;
    dst.col = src.col;
    value_to_fm(src.value, handle->text_store, &dst.value);
    dst.kind = pivot_cell_kind_to_fm(src.kind);
    dst.depth = src.depth;
    dst.field_name = src.field_name.empty() ? nullptr : store_cstr(handle->text_store, src.field_name);
    // Value cells carry their data field's numFmtId text; label cells already carry a format code.
    const bool value_cell = src.role == formulon::pivot::PivotCellRole::DataValue ||
                            src.role == formulon::pivot::PivotCellRole::SubtotalValue ||
                            src.role == formulon::pivot::PivotCellRole::GrandTotalValue;
    const char* value_code = value_cell ? pivot_num_fmt_code(wb->workbook(), src.number_format) : nullptr;
    const std::string_view number_format =
        value_cell ? (value_code != nullptr ? std::string_view(value_code) : std::string_view())
                   : std::string_view(src.number_format);
    dst.number_format = number_format.empty() ? nullptr : store_cstr(handle->text_store, number_format);
    handle->cells.push_back(dst);
  }

  *out = handle.release();
  return 0;
}

extern "C" void fm_pivot_report_destroy(fm_pivot_report_t* cells) {
  delete cells;
}

extern "C" fm_status_t fm_pivot_report_get_cell_count(const fm_pivot_report_t* cells, std::size_t* out_count) {
  clear_last_error();
  if (cells == nullptr || out_count == nullptr) {
    return null_argument();
  }
  *out_count = cells->cells.size();
  return 0;
}

extern "C" fm_status_t fm_pivot_report_get_bounds(const fm_pivot_report_t* cells, std::int32_t* out_present,
                                                  fm_cell_rect_t* out_rect) {
  clear_last_error();
  if (cells == nullptr || out_present == nullptr || out_rect == nullptr) {
    return null_argument();
  }
  *out_rect = fm_cell_rect_t{};
  *out_present = 0;
  if (cells->rows == 0U || cells->cols == 0U) {
    return 0;
  }
  *out_rect = cell_rect_from_extent(cells->top, cells->left, cells->rows, cells->cols);
  *out_present = 1;
  return 0;
}

extern "C" fm_status_t fm_pivot_report_get_cell_at(const fm_pivot_report_t* cells, std::size_t idx,
                                                   fm_pivot_cell_info_t* out) {
  clear_last_error();
  if (cells == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(idx, cells->cells.size(), "index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const std::uint32_t struct_size = out->struct_size;
  *out = cells->cells[idx];
  out->struct_size = struct_size;
  return 0;
}
