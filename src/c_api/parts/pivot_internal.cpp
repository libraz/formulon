
#include "c_api/parts/pivot_internal.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "eval/eval_profile_scope.h"
#include "pivot/field_remap.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_refresh.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"
#include "workbook_internal.h"

namespace formulon {
namespace c_api {
namespace parts {

namespace {

// Every pivot table on every sheet, in sheet order.
std::vector<formulon::pivot::PivotTable*> all_pivots(formulon::Workbook& wb) {
  std::vector<formulon::pivot::PivotTable*> tables;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (auto& table : wb.sheet(s).mutable_pivot_tables()) {
      if (table != nullptr) {
        tables.push_back(table.get());
      }
    }
  }
  return tables;
}

bool on_axis(formulon::pivot::PivotAxis axis) {
  return axis == formulon::pivot::PivotAxis::Row || axis == formulon::pivot::PivotAxis::Col ||
         axis == formulon::pivot::PivotAxis::Page;
}

}  // namespace

formulon::pivot::PivotCache* find_cache_mut(formulon::Workbook& wb, std::uint32_t cache_id) {
  return const_cast<formulon::pivot::PivotCache*>(wb.find_pivot_cache(cache_id));
}

const formulon::pivot::PivotCache* find_cache(const formulon::Workbook& wb, std::uint32_t cache_id) {
  return wb.find_pivot_cache(cache_id);
}

fm_status_t cache_not_found(std::uint32_t cache_id) {
  return not_found(("cache_id " + std::to_string(cache_id) + " names no pivot cache").c_str(),
                   "cache_id=" + std::to_string(cache_id));
}

// Resolves `cache_id` to its cache, writing the binding error and returning
// its status on a miss.
fm_status_t resolve_cache(const fm_workbook_t* wb, std::uint32_t cache_id, const formulon::pivot::PivotCache** out) {
  *out = find_cache(wb->workbook(), cache_id);
  if (*out == nullptr) {
    return cache_not_found(cache_id);
  }
  return 0;
}

// Mutable variant of `resolve_cache` that also rejects a NULL workbook.
fm_status_t resolve_cache_mut(fm_workbook_t* wb, std::uint32_t cache_id, formulon::pivot::PivotCache** out) {
  if (wb == nullptr) {
    return null_argument("wb");
  }
  *out = find_cache_mut(wb->workbook(), cache_id);
  if (*out == nullptr) {
    return cache_not_found(cache_id);
  }
  return 0;
}

fm_status_t resolve_pivot(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                          const formulon::pivot::PivotTable** out) {
  *out = nullptr;
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  if (auto rc = check_index(pivot_index, sheet.pivot_tables().size(), "pivot_index"); rc != 0) {
    return rc;
  }
  const formulon::pivot::PivotTable* table = sheet.pivot_tables()[pivot_index].get();
  if (table == nullptr) {
    return set_binding_error(
        formulon::FormulonErrorCode::kEvalPivotInvalid, "pivot table entry is NULL",
        "sheet_index=" + std::to_string(sheet_index) + " pivot_index=" + std::to_string(pivot_index));
  }
  *out = table;
  return 0;
}

fm_status_t resolve_pivot_mut(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                              formulon::pivot::PivotTable** out) {
  // The pivot tables are owned non-const objects; only the lookup path is shared.
  const formulon::pivot::PivotTable* table = nullptr;
  const fm_status_t rc = resolve_pivot(wb, sheet_index, pivot_index, &table);
  *out = const_cast<formulon::pivot::PivotTable*>(table);
  return rc;
}

fm_status_t resolve_pivot_checked(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                  formulon::pivot::PivotTable** out) {
  if (wb == nullptr) {
    *out = nullptr;
    return null_argument("wb");
  }
  return resolve_pivot_mut(wb, sheet_index, pivot_index, out);
}

fm_status_t resolve_pivot_checked(const fm_workbook_t* wb, std::initializer_list<const void*> out_args,
                                  std::size_t sheet_index, std::size_t pivot_index,
                                  const formulon::pivot::PivotTable** out) {
  bool null_arg = wb == nullptr;
  for (const void* arg : out_args) {
    null_arg = null_arg || arg == nullptr;
  }
  if (null_arg) {
    return null_argument();
  }
  return resolve_pivot(wb, sheet_index, pivot_index, out);
}

bool pivot_spill_info(const formulon::pivot::PivotTable& table, fm_spill_info_t* out) {
  out->anchor = fm_cell_address_t{table.anchor_row(), table.anchor_col()};
  const std::optional<formulon::pivot::BlockedReport>& state = table.blocked();
  if (state.has_value()) {
    const std::uint32_t above = std::min(state->rows_above, table.anchor_row());
    const std::uint32_t first_row = table.anchor_row() - above;
    const std::uint32_t rows = std::min(state->rows, formulon::Sheet::kMaxRows - first_row);
    const std::uint32_t cols = std::min(state->cols, formulon::Sheet::kMaxCols - table.anchor_col());
    out->rect = cell_rect_from_extent(first_row, table.anchor_col(), rows, cols);
    out->blocked = 1;
    out->cause = to_c(state->cause);
    return true;
  }
  if (!table.projected_extent().has_value() && !table.span_read()) {
    return false;
  }
  // A report read from a file and not laid out since covers its location.
  const formulon::MergeRange drawn = formulon::detail::pivot_occupied_extent(table);
  out->rect = fm_cell_rect_t{drawn.first_row, drawn.first_col, drawn.last_row, drawn.last_col};
  out->blocked = 0;
  out->cause = FM_SPILL_CAUSE_ABSENT;
  return true;
}

void invalidate_pivot_result(formulon::Workbook& wb, formulon::pivot::PivotTable& table) {
  formulon::pivot::clear_derived_state(table);
  wb.mark_all_formulas_dirty();
}

void invalidate_pivot_results_for_cache(formulon::Workbook& wb, std::uint32_t cache_id) {
  wb.invalidate_pivot_results(cache_id);
}

void mark_cache_pivots_pending(formulon::Workbook& wb, std::uint32_t cache_id) {
  for (formulon::pivot::PivotTable* table : cache_tables(wb, cache_id)) {
    table->set_projection_pending(true);
  }
}

std::vector<formulon::pivot::PivotTable*> cache_tables(formulon::Workbook& wb, std::uint32_t cache_id) {
  return wb.pivots_of_cache(cache_id);
}

formulon::Expected<void, formulon::Error> project_pivot(formulon::Workbook& wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, formulon::ReportPlacement placement) {
  formulon::Expected<void, formulon::Error> projected = wb.project_pivot(sheet_index, pivot_index, placement);
  formulon::pivot::PivotTable& table = *wb.sheet(sheet_index).mutable_pivot_tables()[pivot_index];
  if (projected && !table.blocked().has_value()) {
    formulon::pivot::drop_derived_location(table);
  }
  return projected;
}

void project_pending_pivots(formulon::Workbook& wb) {
  std::vector<formulon::pivot::PivotTable*> pending = all_pivots(wb);
  pending.erase(std::remove_if(pending.begin(), pending.end(),
                               [](const formulon::pivot::PivotTable* table) { return !table->projection_pending(); }),
                pending.end());
  wb.project_pending_pivots();
  for (formulon::pivot::PivotTable* table : pending) {
    if (!table->projection_pending()) {
      formulon::pivot::drop_derived_location(*table);
    }
  }
}

formulon::Expected<void, formulon::Error> reproject_edited_pivot(formulon::Workbook& wb, std::size_t sheet_index,
                                                                 std::size_t pivot_index) {
  formulon::pivot::PivotTable& table = *wb.sheet(sheet_index).mutable_pivot_tables()[pivot_index];
  if (formulon::pivot::PivotCache* cache = find_cache_mut(wb, table.pivot_cache_id()); cache != nullptr) {
    formulon::pivot::share_axis_fields(*cache, table);
  }
  invalidate_pivot_result(wb, table);
  formulon::Expected<void, formulon::Error> projected =
      project_pivot(wb, sheet_index, pivot_index, formulon::ReportPlacement::kBlock);
  if (!projected) {
    table.set_projection_pending(true);
  }
  return projected;
}

namespace {

// A field this edit put on an axis, or took off every axis, lists its items
// as Excel writes them; a field joining an axis shares its values first.
void list_moved_field_items(formulon::Workbook& wb, formulon::pivot::PivotTable& table,
                            const formulon::pivot::PivotTable& before) {
  formulon::pivot::PivotCache* cache = find_cache_mut(wb, table.pivot_cache_id());
  if (cache == nullptr) {
    return;
  }
  formulon::pivot::share_axis_fields(*cache, table);
  const formulon::eval::ScopedEvalProfile scope(wb.excel_profile());
  const std::size_t count = std::min({table.fields().size(), before.fields().size(), cache->fields().size()});
  for (std::size_t f = 0; f < count; ++f) {
    if (on_axis(before.fields()[f].axis) != on_axis(table.fields()[f].axis)) {
      formulon::pivot::keep_field_items(table.mutable_fields()[f], cache->fields()[f]);
    }
  }
}

}  // namespace

PivotEdit::PivotEdit(formulon::Workbook& wb, std::size_t sheet_index, std::size_t pivot_index,
                     formulon::pivot::PivotTable& table)
    : wb_(wb), sheet_index_(sheet_index), pivot_index_(pivot_index), table_(table), snapshot_(table.clone()) {}

fm_status_t PivotEdit::commit() {
  list_moved_field_items(wb_, table_, *snapshot_);
  formulon::Expected<void, formulon::Error> projected = reproject_edited_pivot(wb_, sheet_index_, pivot_index_);
  if (projected) {
    return 0;
  }
  if (projected.error().code == formulon::FormulonErrorCode::kPreconditionFailed) {
    // The snapshot is a clone, whose runtime id is zero; the restored table is still the same pivot.
    const std::uint32_t runtime_id = table_.runtime_id().value();
    table_ = std::move(*snapshot_);
    table_.runtime_id().assign(runtime_id);
    return set_last_error(projected.error());
  }
  return 0;
}

void set_values_position(formulon::pivot::PivotTable& table, bool row, std::optional<std::size_t> position) {
  if (row) {
    table.set_row_values_position(position);
  } else {
    table.set_col_values_position(position);
  }
}

formulon::Value intern_cache_text(formulon::pivot::PivotCache& cache, std::string_view utf8) {
  cache.mutable_text_storage().emplace_back(utf8.data(), utf8.size());
  return formulon::Value::text(std::string_view(cache.mutable_text_storage().back()));
}

formulon::pivot::PivotAxis pivot_axis_from_fm(fm_pivot_axis_t axis) {
  switch (axis) {
    case FM_PIVOT_AXIS_ABSENT:
      return formulon::pivot::PivotAxis::None;
    case FM_PIVOT_AXIS_ROW:
      return formulon::pivot::PivotAxis::Row;
    case FM_PIVOT_AXIS_COLUMN:
      return formulon::pivot::PivotAxis::Col;
    case FM_PIVOT_AXIS_VALUE:
      return formulon::pivot::PivotAxis::Value;
    case FM_PIVOT_AXIS_PAGE:
      return formulon::pivot::PivotAxis::Page;
  }
  return formulon::pivot::PivotAxis::Row;
}

formulon::pivot::Aggregation pivot_agg_from_fm(fm_pivot_aggregation_t agg) {
  return static_cast<formulon::pivot::Aggregation>(agg);
}

formulon::pivot::SubtotalFn pivot_subtotal_from_fm(fm_pivot_aggregation_t agg) {
  return static_cast<formulon::pivot::SubtotalFn>(agg);
}

formulon::pivot::ShowValuesAs pivot_show_as_from_fm(fm_pivot_show_values_as_t show_as) {
  return static_cast<formulon::pivot::ShowValuesAs>(show_as);
}

formulon::pivot::FilterType pivot_filter_type_from_fm(fm_pivot_filter_kind_t t) {
  return static_cast<formulon::pivot::FilterType>(t);
}

formulon::pivot::CalendarSystem pivot_calendar_from_fm(fm_pivot_calendar_t c) {
  return static_cast<formulon::pivot::CalendarSystem>(c);
}

std::optional<fm_pivot_axis_t> pivot_axis_to_fm(formulon::pivot::PivotAxis axis) {
  switch (axis) {
    case formulon::pivot::PivotAxis::Row:
      return FM_PIVOT_AXIS_ROW;
    case formulon::pivot::PivotAxis::Col:
      return FM_PIVOT_AXIS_COLUMN;
    case formulon::pivot::PivotAxis::Value:
      return FM_PIVOT_AXIS_VALUE;
    case formulon::pivot::PivotAxis::Page:
      return FM_PIVOT_AXIS_PAGE;
    case formulon::pivot::PivotAxis::None:
    default:
      return std::nullopt;
  }
}

fm_status_t pivot_count(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                        std::size_t* out_count, TableCount count) {
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out_count}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  *out_count = count(*table);
  return 0;
}

void grow_record_cells(formulon::pivot::PivotCacheRecord& rec, std::size_t field_index) {
  if (rec.cells.size() <= field_index) {
    rec.cells.resize(field_index + 1U, formulon::Value::blank());
  }
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon
