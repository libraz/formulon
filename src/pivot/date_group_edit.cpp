// Workbook-facing date- and numeric-group editing: the pivot-field and axis
// rewiring that follows a cache regroup, behind set_date_group /
// clear_date_group and set_numeric_group / clear_numeric_group.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "pivot/field_grouping.h"
#include "pivot/field_grouping_internal.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_refresh.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon::pivot {
namespace {

using detail::group_error;
using detail::group_run;
using detail::GroupRun;
using detail::resolve_date_group_bounds;
using detail::resolve_numeric_group;

// A row or column axis in document order, the Values marker included.
std::vector<std::int64_t> axis_entries(const PivotTable& table, bool row) {
  const std::vector<std::uint32_t>& order = row ? table.row_field_order() : table.col_field_order();
  const std::optional<std::size_t> values = row ? table.row_values_position() : table.col_values_position();
  std::vector<std::int64_t> entries(order.begin(), order.end());
  if (values.has_value()) {
    entries.insert(entries.begin() + static_cast<std::ptrdiff_t>(std::min(*values, entries.size())),
                   kValuesFieldPosition);
  }
  return entries;
}

void set_axis_entries(PivotTable& table, bool row, const std::vector<std::int64_t>& entries) {
  std::vector<std::uint32_t> order;
  std::optional<std::size_t> values;
  for (std::size_t i = 0; i < entries.size(); ++i) {
    if (entries[i] == kValuesFieldPosition) {
      values = i;
    } else {
      order.push_back(static_cast<std::uint32_t>(entries[i]));
    }
  }
  if (row) {
    table.mutable_row_field_order() = std::move(order);
    table.set_row_values_position(values);
  } else {
    table.mutable_col_field_order() = std::move(order);
    table.set_col_values_position(values);
  }
}

PivotAxis axis_of(bool row) {
  return row ? PivotAxis::Row : PivotAxis::Col;
}

// Gives `table` a pivot field for each cache field from `first` on and puts
// them on the axis that holds the base: in its place on the grouped pivot,
// before it on the others.
void add_group_pivot_fields(PivotTable& table, const PivotCache& cache, std::size_t base, std::size_t first,
                            bool replace_base) {
  std::vector<PivotField>& fields = table.mutable_fields();
  while (fields.size() < cache.fields().size()) {
    const std::size_t fi = fields.size();
    PivotField field;
    field.source_name = cache.fields()[fi].name;
    field.axis = PivotAxis::None;
    if (fi >= first) {
      // Excel lists group items in group order, with no sort stated.
      field.sort.manual = true;
      const std::size_t item_count = cache.fields()[fi].field_group->group_items.size();
      for (std::size_t i = 0; i < item_count; ++i) {
        PivotItem item;
        item.has_cache_index = true;
        item.cache_index = static_cast<std::uint32_t>(i);
        field.items.push_back(std::move(item));
      }
    }
    fields.push_back(std::move(field));
  }
  for (const bool row : {true, false}) {
    std::vector<std::int64_t> entries = axis_entries(table, row);
    const auto at = std::find(entries.begin(), entries.end(), static_cast<std::int64_t>(base));
    if (at == entries.end()) {
      continue;
    }
    // Coarsest outermost.
    std::vector<std::int64_t> groups;
    for (std::size_t fi = cache.fields().size(); fi > first; --fi) {
      groups.push_back(static_cast<std::int64_t>(fi - 1U));
      fields[fi - 1U].axis = axis_of(row);
    }
    const auto pos = entries.erase(at, replace_base ? at + 1 : at);
    entries.insert(pos, groups.begin(), groups.end());
    if (replace_base) {
      fields[base].axis = PivotAxis::None;
    }
    set_axis_entries(table, row, entries);
  }
}

// Drops `table`'s pivot fields from `first` on; the base takes the axis
// position the first of them held unless it is on an axis already.
void remove_group_pivot_fields(PivotTable& table, std::size_t base, std::size_t first) {
  std::vector<PivotField>& fields = table.mutable_fields();
  if (fields.size() <= first) {
    return;
  }
  for (const bool row : {true, false}) {
    std::vector<std::int64_t> entries = axis_entries(table, row);
    std::vector<std::int64_t> kept;
    for (const std::int64_t entry : entries) {
      if (entry < 0 || static_cast<std::size_t>(entry) < first) {
        kept.push_back(entry);
        continue;
      }
      if (base < fields.size() && fields[base].axis == PivotAxis::None) {
        kept.push_back(static_cast<std::int64_t>(base));
        fields[base].axis = axis_of(row);
      }
    }
    if (kept != entries) {
      set_axis_entries(table, row, kept);
    }
  }
  std::vector<PivotPageField>& pages = table.mutable_page_fields();
  pages.erase(std::remove_if(pages.begin(), pages.end(),
                             [first](const PivotPageField& page) { return page.field_index >= first; }),
              pages.end());
  std::vector<PivotDataField>& data = table.mutable_data_fields();
  data.erase(
      std::remove_if(data.begin(), data.end(), [first](const PivotDataField& df) { return df.field_index >= first; }),
      data.end());
  table.erase_active_filters_if(
      [first](const PivotFilter& filter) { return filter.source.has_value() && filter.source->field_index >= first; });
  fields.resize(first);
}

// Calls `fn(PivotTable&)` for every pivot table on cache `cache_id`.
template <class Fn>
void for_each_cache_table(Workbook& wb, std::uint32_t cache_id, Fn&& fn) {
  for (std::size_t si = 0; si < wb.sheet_count(); ++si) {
    for (std::unique_ptr<PivotTable>& table : wb.sheet(si).mutable_pivot_tables()) {
      if (table != nullptr && table->pivot_cache_id() == cache_id) {
        fn(*table);
      }
    }
  }
}

// Shares the fields every pivot on `cache` places on an axis, as a pivot edit
// does: a cleared group puts its base back on the axes.
void share_cache_axis_fields(Workbook& wb, PivotCache& cache) {
  for_each_cache_table(wb, cache.cache_id(), [&](const PivotTable& table) { share_axis_fields(cache, table); });
}

// Marks the shared items of a date group's base as dates, so the base keeps
// its date-ness when it was shared from numbers.
void mark_base_dates(PivotCache& cache, std::size_t base) {
  PivotCacheField& field = cache.mutable_fields()[base];
  field.shared_item_is_date.resize(field.shared_items.size(), false);
  for (std::size_t k = 0; k < field.shared_items.size(); ++k) {
    field.shared_item_is_date[k] = field.shared_item_is_date[k] || field.shared_items[k].is_number();
  }
}

void invalidate_tables(Workbook& wb, std::uint32_t cache_id) {
  for_each_cache_table(wb, cache_id, [](PivotTable& table) {
    table.clear_last_result();
    table.clear_span_authored();
    table.mark_layout_dirty();
  });
  wb.mark_all_formulas_dirty();
}

// The pivot `(sheet, pivot)`, its cache and the base of its field `field`.
struct GroupTarget {
  PivotTable* table = nullptr;
  PivotCache* cache = nullptr;
  std::size_t base = 0;
};

Expected<GroupTarget, Error> resolve_group_target(Workbook& wb, std::size_t sheet, std::size_t pivot,
                                                  std::size_t field) {
  const std::string where =
      "sheet=" + std::to_string(sheet) + " pivot=" + std::to_string(pivot) + " field=" + std::to_string(field);
  if (sheet >= wb.sheet_count()) {
    return group_error(FormulonErrorCode::kSheetIndexOutOfRange, "no such sheet", where);
  }
  if (pivot >= wb.sheet(sheet).pivot_tables().size()) {
    return group_error(FormulonErrorCode::kIndexOutOfRange, "no such pivot table", where);
  }
  if (wb.sheet(sheet).pivot_tables()[pivot] == nullptr) {
    return group_error(FormulonErrorCode::kInvalidArgument, "no such pivot table", where);
  }
  GroupTarget target;
  target.table = wb.sheet(sheet).mutable_pivot_tables()[pivot].get();
  for (std::unique_ptr<PivotCache>& cache : wb.mutable_pivot_caches()) {
    if (cache != nullptr && cache->cache_id() == target.table->pivot_cache_id()) {
      target.cache = cache.get();
    }
  }
  if (target.cache == nullptr) {
    return group_error(FormulonErrorCode::kInvalidArgument, "the pivot table has no cache", where);
  }
  if (field >= target.table->fields().size() || field >= target.cache->fields().size()) {
    return group_error(FormulonErrorCode::kIndexOutOfRange, "no such pivot field", where);
  }
  target.base = group_base_index(*target.cache, field);
  return target;
}

// Takes the group fields over `base` off every pivot on `cache` and out of
// the cache; `run` is their checked position.
void remove_group(Workbook& wb, PivotCache& cache, std::size_t base, const GroupRun& run) {
  for_each_cache_table(wb, cache.cache_id(),
                       [&](PivotTable& table) { remove_group_pivot_fields(table, base, run.first); });
  (void)ungroup_cache_field(cache, base);
}

}  // namespace

Expected<void, Error> set_date_group(Workbook& wb, std::size_t sheet, std::size_t pivot, std::size_t field,
                                     std::uint32_t periods, CalendarSystem calendar, std::uint32_t interval_days,
                                     std::optional<double> start, std::optional<double> end) {
  ASSIGN_OR_RETURN(GroupTarget target, resolve_group_target(wb, sheet, pivot, field));
  PivotCache& cache = *target.cache;
  const std::size_t base = target.base;
  const DateGroupSpec spec{periods, calendar, interval_days, start, end};
  RETURN_IF_ERROR(resolve_date_group_bounds(cache, base, spec, wb.date1904()));
  ASSIGN_OR_RETURN(GroupRun run, group_run(cache, base));
  remove_group(wb, cache, base, run);
  ASSIGN_OR_RETURN(std::size_t first, group_cache_by_date(cache, base, spec, wb.date1904(), wb.excel_profile()));
  for_each_cache_table(wb, cache.cache_id(), [&](PivotTable& table) {
    add_group_pivot_fields(table, cache, base, first, &table == target.table);
    resolve_pivot_names(table, cache);
  });
  share_cache_axis_fields(wb, cache);
  mark_base_dates(cache, base);
  invalidate_tables(wb, cache.cache_id());
  return Expected<void, Error>::Ok();
}

Expected<void, Error> clear_date_group(Workbook& wb, std::size_t sheet, std::size_t pivot, std::size_t field) {
  ASSIGN_OR_RETURN(GroupTarget target, resolve_group_target(wb, sheet, pivot, field));
  PivotCache& cache = *target.cache;
  ASSIGN_OR_RETURN(GroupRun run, group_run(cache, target.base));
  if (run.count == 0U) {
    return Expected<void, Error>::Ok();
  }
  remove_group(wb, cache, target.base, run);
  share_cache_axis_fields(wb, cache);
  invalidate_tables(wb, cache.cache_id());
  return Expected<void, Error>::Ok();
}

Expected<void, Error> set_numeric_group(Workbook& wb, std::size_t sheet, std::size_t pivot, std::size_t field,
                                        std::optional<double> start, std::optional<double> end, double interval) {
  ASSIGN_OR_RETURN(GroupTarget target, resolve_group_target(wb, sheet, pivot, field));
  PivotCache& cache = *target.cache;
  const std::size_t base = target.base;
  ASSIGN_OR_RETURN(NumericGroupState state, resolve_numeric_group(cache, base, start, end, interval));
  ASSIGN_OR_RETURN(GroupRun run, group_run(cache, base));
  remove_group(wb, cache, base, run);
  ASSIGN_OR_RETURN(std::size_t first,
                   detail::group_cache_by_number(cache, base, state, wb.date1904(), wb.excel_profile()));
  for_each_cache_table(wb, cache.cache_id(), [&](PivotTable& table) {
    add_group_pivot_fields(table, cache, base, first, &table == target.table);
    resolve_pivot_names(table, cache);
  });
  share_cache_axis_fields(wb, cache);
  invalidate_tables(wb, cache.cache_id());
  return Expected<void, Error>::Ok();
}

Expected<void, Error> clear_numeric_group(Workbook& wb, std::size_t sheet, std::size_t pivot, std::size_t field) {
  ASSIGN_OR_RETURN(GroupTarget target, resolve_group_target(wb, sheet, pivot, field));
  PivotCache& cache = *target.cache;
  if (!numeric_group_state(cache, target.base).has_value()) {
    return Expected<void, Error>::Ok();
  }
  ASSIGN_OR_RETURN(GroupRun run, group_run(cache, target.base));
  remove_group(wb, cache, target.base, run);
  share_cache_axis_fields(wb, cache);
  invalidate_tables(wb, cache.cache_id());
  return Expected<void, Error>::Ok();
}

}  // namespace formulon::pivot
