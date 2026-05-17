// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared private helpers for the three pivot-related TUs
// (`pivot_layout.cpp`, `pivot_cache.cpp`, `pivot_table.cpp`).
//
// The pivot mutation surface is large enough that splitting it into the
// cache/table axes is worth a dedicated TU each; this header carries
// the helpers that more than one of them needs (mutable-cache lookups,
// fm_* enum -> engine-enum mapping, pivot-result cache invalidation,
// cache-text interning).

#ifndef FORMULON_C_API_PARTS_PIVOT_INTERNAL_H_
#define FORMULON_C_API_PARTS_PIVOT_INTERNAL_H_

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace c_api {
namespace parts {

// Mutable handle on the pivot-cache list. The accessor we have on
// `Workbook` is const-only; the vector itself is not const because we
// hold a non-const workbook through `fm_workbook_t::workbook()`. The
// const_cast is therefore well-defined.
std::vector<std::unique_ptr<formulon::pivot::PivotCache>>& mutable_pivot_caches(formulon::Workbook& wb);

formulon::pivot::PivotCache* find_cache_mut(formulon::Workbook& wb, std::uint32_t cache_id);
const formulon::pivot::PivotCache* find_cache(const formulon::Workbook& wb, std::uint32_t cache_id);

// Sheet+pivot resolution. On miss the binding error is populated and
// `nullptr` is returned; callers should propagate as `kInvalidArgument`.
formulon::pivot::PivotTable* resolve_pivot_mut(formulon::Workbook& wb, std::size_t sheet_index, std::size_t pivot_index,
                                               const char* fn);
const formulon::pivot::PivotTable* resolve_pivot(const formulon::Workbook& wb, std::size_t sheet_index,
                                                 std::size_t pivot_index, const char* fn);

// Invalidates the pivot's `last_result_` cache so the next layout call
// recomputes. Called after every mutation that could affect projection.
void invalidate_pivot_result(formulon::pivot::PivotTable& table);

// Interns `utf8` in the cache's text storage and returns a Value::text
// referencing the stable backing string.
formulon::Value intern_cache_text(formulon::pivot::PivotCache& cache, std::string_view utf8);

// fm_* enum to engine-enum mapping.
formulon::pivot::PivotAxis pivot_axis_from_fm(fm_pivot_axis_t axis);
formulon::pivot::Aggregation pivot_agg_from_fm(fm_pivot_aggregation_t agg);
formulon::pivot::SubtotalFn pivot_subtotal_from_fm(fm_pivot_aggregation_t agg);
formulon::pivot::ShowValuesAs pivot_show_as_from_fm(fm_pivot_show_as_t show_as);
formulon::pivot::FilterType pivot_filter_type_from_fm(fm_pivot_filter_type_t t);
formulon::pivot::DateGrouping pivot_date_grouping_from_fm(fm_pivot_date_grouping_t g);
formulon::pivot::CalendarSystem pivot_calendar_from_fm(fm_pivot_calendar_t c);

// Ensures the record's cell vector has at least `field_idx + 1` slots.
// Newly-added slots are initialised to Blank so existing semantics
// match the cache reader's behaviour for short rows.
void grow_record_cells(formulon::pivot::PivotCacheRecord& rec, std::size_t field_idx);

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_PIVOT_INTERNAL_H_
