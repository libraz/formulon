
#include "c_api/parts/pivot_internal.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace c_api {
namespace parts {

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

formulon::pivot::PivotTable* resolve_pivot_mut(formulon::Workbook& wb, std::size_t sheet_index, std::size_t pivot_index,
                                               const char* fn) {
  if (sheet_index >= wb.sheet_count()) {
    set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "sheet_index out of range",
                      std::string(fn) + ": sheet_index=" + std::to_string(sheet_index));
    return nullptr;
  }
  formulon::Sheet& sheet = wb.sheet(sheet_index);
  auto& pivots = sheet.mutable_pivot_tables();
  if (pivot_index >= pivots.size()) {
    set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "pivot_index out of range",
        std::string(fn) + ": pivot_index=" + std::to_string(pivot_index) + " count=" + std::to_string(pivots.size()));
    return nullptr;
  }
  formulon::pivot::PivotTable* table = pivots[pivot_index].get();
  if (table == nullptr) {
    set_binding_error(formulon::FormulonErrorCode::kEvalPivotInvalid, "pivot table entry is NULL",
                      std::string(fn) + ": sheet_index=" + std::to_string(sheet_index) +
                          " pivot_index=" + std::to_string(pivot_index));
    return nullptr;
  }
  return table;
}

const formulon::pivot::PivotTable* resolve_pivot(const formulon::Workbook& wb, std::size_t sheet_index,
                                                 std::size_t pivot_index, const char* fn) {
  if (sheet_index >= wb.sheet_count()) {
    set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "sheet_index out of range",
                      std::string(fn) + ": sheet_index=" + std::to_string(sheet_index));
    return nullptr;
  }
  const formulon::Sheet& sheet = wb.sheet(sheet_index);
  if (pivot_index >= sheet.pivot_tables().size()) {
    set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "pivot_index out of range",
                      std::string(fn) + ": pivot_index=" + std::to_string(pivot_index) +
                          " count=" + std::to_string(sheet.pivot_tables().size()));
    return nullptr;
  }
  const formulon::pivot::PivotTable* table = sheet.pivot_tables()[pivot_index].get();
  if (table == nullptr) {
    set_binding_error(formulon::FormulonErrorCode::kEvalPivotInvalid, "pivot table entry is NULL",
                      std::string(fn) + ": sheet_index=" + std::to_string(sheet_index) +
                          " pivot_index=" + std::to_string(pivot_index));
    return nullptr;
  }
  return table;
}

void invalidate_pivot_result(formulon::pivot::PivotTable& table) {
  table.clear_last_result();
  table.clear_span_authored();
}

void invalidate_pivot_results_for_cache(formulon::Workbook& wb, std::uint32_t cache_id) {
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (auto& table : wb.sheet(s).mutable_pivot_tables()) {
      if (table != nullptr && table->pivot_cache_id() == cache_id) {
        table->clear_last_result();
      }
    }
  }
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

formulon::pivot::ShowValuesAs pivot_show_as_from_fm(fm_pivot_show_as_t show_as) {
  return static_cast<formulon::pivot::ShowValuesAs>(show_as);
}

formulon::pivot::FilterType pivot_filter_type_from_fm(fm_pivot_filter_type_t t) {
  return static_cast<formulon::pivot::FilterType>(t);
}

formulon::pivot::DateGrouping pivot_date_grouping_from_fm(fm_pivot_date_grouping_t g) {
  return static_cast<formulon::pivot::DateGrouping>(g);
}

formulon::pivot::CalendarSystem pivot_calendar_from_fm(fm_pivot_calendar_t c) {
  return static_cast<formulon::pivot::CalendarSystem>(c);
}

void grow_record_cells(formulon::pivot::PivotCacheRecord& rec, std::size_t field_idx) {
  if (rec.cells.size() <= field_idx) {
    rec.cells.resize(field_idx + 1U, formulon::Value::blank());
  }
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon
