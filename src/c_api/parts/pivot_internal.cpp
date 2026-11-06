
#include "c_api/parts/pivot_internal.h"

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
#include "pivot/field_grouping.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace c_api {
namespace parts {

namespace {

void clear_pivot_derived_state(formulon::pivot::PivotTable& table) {
  table.clear_last_result();
  table.clear_span_authored();
  table.mark_layout_dirty();
}

// The location offsets read with a span Excel wrote describe that report only;
// once the span is no longer authored and the report has been projected, the
// writer derives them from the model.
void drop_derived_location(formulon::pivot::PivotTable& table) {
  if (!table.has_authored_span() && table.projected_extent().has_value()) {
    table.set_location_attributes(std::nullopt, std::nullopt, std::nullopt, std::nullopt, std::nullopt);
  }
}

}  // namespace

std::vector<std::unique_ptr<formulon::pivot::PivotCache>>& mutable_pivot_caches(formulon::Workbook& wb) {
  return const_cast<std::vector<std::unique_ptr<formulon::pivot::PivotCache>>&>(wb.pivot_caches());
}

formulon::pivot::PivotCache* find_cache_mut(formulon::Workbook& wb, std::uint32_t cache_id) {
  for (std::unique_ptr<formulon::pivot::PivotCache>& c : mutable_pivot_caches(wb)) {
    if (c != nullptr && c->cache_id() == cache_id) {
      return c.get();
    }
  }
  return nullptr;
}

const formulon::pivot::PivotCache* find_cache(const formulon::Workbook& wb, std::uint32_t cache_id) {
  return wb.find_pivot_cache(cache_id);
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

void invalidate_pivot_result(formulon::Workbook& wb, formulon::pivot::PivotTable& table) {
  clear_pivot_derived_state(table);
  wb.mark_all_formulas_dirty();
}

void invalidate_pivot_results_for_cache(formulon::Workbook& wb, std::uint32_t cache_id) {
  // Group fields follow their base, so an edited cache re-derives them.
  if (formulon::pivot::PivotCache* cache = find_cache_mut(wb, cache_id); cache != nullptr) {
    formulon::pivot::materialize_group_fields(*cache, wb.date1904());
  }
  const std::vector<formulon::pivot::PivotTable*> tables = cache_tables(wb, cache_id);
  for (formulon::pivot::PivotTable* table : tables) {
    clear_pivot_derived_state(*table);
    table->set_projection_pending(true);
  }
  if (!tables.empty()) {
    wb.mark_all_formulas_dirty();
  }
}

void mark_cache_pivots_pending(formulon::Workbook& wb, std::uint32_t cache_id) {
  for (formulon::pivot::PivotTable* table : cache_tables(wb, cache_id)) {
    table->set_projection_pending(true);
  }
}

std::vector<formulon::pivot::PivotTable*> cache_tables(formulon::Workbook& wb, std::uint32_t cache_id) {
  std::vector<formulon::pivot::PivotTable*> tables;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (auto& table : wb.sheet(s).mutable_pivot_tables()) {
      if (table != nullptr && table->pivot_cache_id() == cache_id) {
        tables.push_back(table.get());
      }
    }
  }
  return tables;
}

formulon::Expected<void, formulon::Error> project_pivot(formulon::Workbook& wb, std::size_t sheet_index,
                                                        std::size_t pivot_index) {
  formulon::Expected<void, formulon::Error> projected = wb.project_pivot(sheet_index, pivot_index);
  if (projected) {
    drop_derived_location(*wb.sheet(sheet_index).mutable_pivot_tables()[pivot_index]);
  }
  return projected;
}

void project_pending_pivots(formulon::Workbook& wb) {
  std::vector<formulon::pivot::PivotTable*> pending;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (auto& table : wb.sheet(s).mutable_pivot_tables()) {
      if (table != nullptr && table->projection_pending()) {
        pending.push_back(table.get());
      }
    }
  }
  wb.project_pending_pivots();
  for (formulon::pivot::PivotTable* table : pending) {
    if (!table->projection_pending()) {
      drop_derived_location(*table);
    }
  }
}

PivotEdit::PivotEdit(formulon::Workbook& wb, std::size_t sheet_index, std::size_t pivot_index,
                     formulon::pivot::PivotTable& table)
    : wb_(wb), sheet_index_(sheet_index), pivot_index_(pivot_index), table_(table), snapshot_(table.clone()) {}

fm_status_t PivotEdit::commit() {
  invalidate_pivot_result(wb_, table_);
  formulon::Expected<void, formulon::Error> projected = project_pivot(wb_, sheet_index_, pivot_index_);
  if (projected) {
    return 0;
  }
  if (projected.error().code == formulon::FormulonErrorCode::kPreconditionFailed) {
    table_ = std::move(*snapshot_);
    return set_last_error(projected.error());
  }
  table_.set_projection_pending(true);
  return 0;
}

formulon::Value intern_cache_text(formulon::pivot::PivotCache& cache, std::string_view utf8) {
  cache.mutable_text_storage().emplace_back(utf8.data(), utf8.size());
  return formulon::Value::text(std::string_view(cache.mutable_text_storage().back()));
}

formulon::pivot::PivotAxis pivot_axis_from_fm(fm_pivot_axis_t axis) {
  switch (axis) {
    case FM_PIVOT_AXIS_ROW:
      return formulon::pivot::PivotAxis::Row;
    case FM_PIVOT_AXIS_COL:
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

formulon::pivot::FilterType pivot_filter_type_from_fm(fm_pivot_filter_type_t t) {
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
      return FM_PIVOT_AXIS_COL;
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
