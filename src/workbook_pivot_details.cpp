// @size-budget: 20 KB
//
// `Workbook::show_pivot_details`: Excel's Show Details on a pivot value cell.
// The records are selected and ordered before anything is mutated; the new
// sheet is then appended, filled, given its table, and moved in front of the
// pivot's sheet.

#include <cstddef>
#include <cstdint>
#include <map>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "eval/recalc_engine.h"
#include "new_workbook_defaults.h"
#include "pivot/calculated.h"
#include "pivot/field_grouping.h"
#include "pivot/field_remap.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_refresh.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/record_access.h"
#include "pivot/value_order.h"
#include "sheet.h"
#include "sheet_name.h"
#include "styles.h"
#include "table.h"
#include "utils/a1_ref.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using ValueRank = std::map<Value, std::size_t, pivot::ValueLess>;

constexpr std::string_view kDetailTableStyle = "TableStyleMedium9";

bool same_item(const Value& lhs, const Value& rhs) {
  return !pivot::value_less(lhs, rhs) && !pivot::value_less(rhs, lhs);
}

bool record_on_path(const pivot::PivotCache& cache, const pivot::PivotCacheRecord& record,
                    const std::vector<pivot::PivotItemRef>& items) {
  for (const pivot::PivotItemRef& item : items) {
    if (!same_item(pivot::cell_value(cache, record, item.field_index), item.identity)) {
      return false;
    }
  }
  return true;
}

/// Identities of the axis nodes `depth` levels down, in display order.
void collect_level(const std::vector<pivot::AxisHierarchyNode>& nodes, std::size_t depth, std::vector<Value>& out) {
  for (const pivot::AxisHierarchyNode& node : nodes) {
    if (depth > 0U) {
      collect_level(node.children, depth - 1U, out);
    } else if (node.identity.has_value()) {
      out.push_back(*node.identity);
    }
  }
}

/// Display position of field `field`'s value in each record of `records`.
/// `axis` is the evaluated axis the field sits on, at `depth`.
std::vector<std::size_t> display_ranks(const pivot::PivotTable& table, const pivot::PivotCache& cache,
                                       std::uint32_t field, const std::vector<pivot::AxisHierarchyNode>& axis,
                                       std::size_t depth, const std::vector<std::size_t>& records) {
  const pivot::PivotField* spec = field < table.fields().size() ? &table.fields()[field] : nullptr;
  ValueRank rank;
  if (spec != nullptr && spec->sort.manual) {
    for (std::size_t i = 0; i < spec->items.size(); ++i) {
      if (const std::optional<Value> key = pivot::pivot_item_record_value(cache, field, spec->items[i])) {
        rank.emplace(*key, i);
      }
    }
    // Values `items` never lists follow in source first-appearance order.
    std::size_t next = spec->items.size();
    for (const std::size_t r : records) {
      if (rank.emplace(pivot::cell_value(cache, cache.records()[r], field), next).second) {
        ++next;
      }
    }
  } else {
    ValueRank labels;
    for (const std::size_t r : records) {
      labels.emplace(pivot::cell_value(cache, cache.records()[r], field), 0U);
    }
    std::size_t position = 0;
    for (auto& entry : labels) {
      entry.second = position++;
    }
    if (spec != nullptr && !spec->sort.by_field.empty()) {
      // A value-sorted field ranks as the evaluated axis shows it; items it no longer shows follow by label.
      std::vector<Value> shown;
      collect_level(axis, depth, shown);
      for (const Value& value : shown) {
        rank.emplace(value, rank.size());
      }
      const std::size_t base = rank.size();
      for (const auto& entry : labels) {
        rank.emplace(entry.first, base + entry.second);
      }
    } else {
      const bool ascending = spec == nullptr || spec->sort.ascending;
      for (const auto& entry : labels) {
        rank.emplace(entry.first, ascending ? entry.second : labels.size() - 1U - entry.second);
      }
    }
  }
  std::vector<std::size_t> out;
  out.reserve(records.size());
  for (const std::size_t r : records) {
    out.push_back(rank.find(pivot::cell_value(cache, cache.records()[r], field))->second);
  }
  return out;
}

}  // namespace

Expected<std::size_t, Error> Workbook::show_pivot_details(std::size_t sheet, std::size_t pivot, std::uint32_t row,
                                                          std::uint32_t col) {
  if (sheet >= sheet_count() || pivot >= sheets_[sheet].pivot_tables().size() ||
      sheets_[sheet].pivot_tables()[pivot] == nullptr) {
    return make_error(FormulonErrorCode::kInvalidArgument, "sheet or pivot index out of range",
                      "sheet=" + std::to_string(sheet) + " pivot=" + std::to_string(pivot));
  }
  const pivot::PivotTable& table = *sheets_[sheet].pivot_tables()[pivot];
  if (!table.enable_drill()) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "The pivot disables drill-down");
  }
  const pivot::PivotCache* cache = find_pivot_cache(table.pivot_cache_id());
  if (cache == nullptr) {
    return make_error(FormulonErrorCode::kEvalPivotMissing, "pivot cache not found",
                      "cache_id=" + std::to_string(table.pivot_cache_id()));
  }

  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(excel_profile_);
  pivot::PivotFilterEnv env{pinned_now_, date1904_, &styles_};
  if (!env.pinned_now.has_value()) {
    env.pinned_now = date_time::host_civil_time();
  }
  std::shared_ptr<const pivot::PivotResult> result = table.last_result();
  if (!result) {
    auto evaluated = pivot::evaluate(table, *cache, options, env);
    if (!evaluated) {
      return std::move(evaluated.error());
    }
    table.set_last_result(std::move(evaluated.value()));
    result = table.last_result();
  }
  auto path = pivot::cell_item_path(table, *result, options, row, col);
  if (!path) {
    return std::move(path.error());
  }

  // Selection and order collate items under the workbook's profile, as evaluation does.
  std::vector<std::size_t> records;
  std::vector<std::vector<std::size_t>> keys;
  {
    const eval::ScopedEvalProfile profile_scope(excel_profile_);
    const pivot::PreparedRecordFilter filter(table, *cache, env);
    for (std::size_t r = 0; r < cache->records().size(); ++r) {
      const pivot::PivotCacheRecord& record = cache->records()[r];
      if (filter.passes(record) && record_on_path(*cache, record, path.value().row_items) &&
          record_on_path(*cache, record, path.value().col_items)) {
        records.push_back(r);
      }
    }
    const auto add_axis_keys = [&](const std::vector<std::uint32_t>& order,
                                   const std::vector<pivot::AxisHierarchyNode>& axis) {
      for (std::size_t depth = order.size(); depth-- > 0U;) {
        keys.push_back(display_ranks(table, *cache, order[depth], axis, depth, records));
      }
    };
    add_axis_keys(table.row_field_order(), result->rows);
    add_axis_keys(table.col_field_order(), result->cols);
  }
  const auto key_less = [&keys](std::uint32_t lhs, std::uint32_t rhs) {
    for (const std::vector<std::size_t>& key : keys) {
      if (key[lhs] != key[rhs]) {
        return key[lhs] < key[rhs];
      }
    }
    return false;
  };
  std::vector<std::uint32_t> order;
  sorted_index_order(order, static_cast<std::uint32_t>(records.size()), make_index_less(key_less));

  std::vector<std::size_t> columns;
  for (std::size_t f = 0; f < cache->fields().size(); ++f) {
    if (cache->fields()[f].is_database_field) {
      columns.push_back(f);
    }
  }

  const NewWorkbookDefaults& defaults = new_workbook_defaults(excel_profile_.locale);
  const std::string name = next_default_sheet_name();

  TableMetadata detail;
  detail.name = next_default_table_name(defaults.table_word);
  detail.display_name = detail.name;
  detail.ref = "A1:";
  a1::append_a1(detail.ref, static_cast<std::uint32_t>(records.size()),
                static_cast<std::uint32_t>(columns.empty() ? 0U : columns.size() - 1U));
  detail.totals_row_shown = false;
  for (const std::size_t f : columns) {
    TableColumn column;
    column.name = cache->fields()[f].name;
    detail.columns.push_back(std::move(column));
  }
  detail.style_info.present = true;
  detail.style_info.name = std::string(kDetailTableStyle);
  detail.style_info.show_row_stripes = true;

  // Mutation starts here.
  auto added = add_sheet_entered(name);
  if (!added) {
    return std::move(added.error());
  }
  const std::size_t appended = sheet_count() - 1U;
  detail.sheet_index = appended;
  if (auto made = add_table(std::move(detail)); !made) {
    (void)remove_sheet(static_cast<std::uint32_t>(appended));
    return std::move(made.error());
  }

  for (std::size_t c = 0; c < columns.size(); ++c) {
    const pivot::PivotCacheField& field = cache->fields()[columns[c]];
    const auto out_col = static_cast<std::uint32_t>(c);
    (void)set_cell_text(appended, 0U, out_col, field.name);
    std::optional<std::uint32_t> xf;
    if (field.num_fmt_id.value_or(0U) != 0U && *field.num_fmt_id <= 0xFFFFU) {
      CellXf format = styles_.cell_xfs.empty() ? CellXf{} : styles_.cell_xfs.front();
      format.num_fmt_id = static_cast<std::uint16_t>(*field.num_fmt_id);
      format.apply_number_format = true;
      xf = intern_cell_xf(styles_, std::move(format));
    }
    for (std::size_t i = 0; i < order.size(); ++i) {
      const auto out_row = static_cast<std::uint32_t>(i + 1U);
      const Value value = pivot::cell_value(*cache, cache->records()[records[order[i]]], columns[c]);
      if (value.is_text()) {
        (void)set_cell_text(appended, out_row, out_col, value.as_text());
      } else if (!value.is_blank()) {
        (void)set_cell_value(appended, out_row, out_col, value);
      }
      if (xf.has_value()) {
        (void)set_cell_xf_index(appended, out_row, out_col, *xf);
      }
    }
  }

  {
    // A new sheet is no span endpoint, so its placement rewrites nothing.
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    move_sheet_locked(static_cast<std::uint32_t>(appended), static_cast<std::uint32_t>(sheet));
  }
  set_active_sheet(static_cast<std::uint32_t>(sheet));
  return sheet;
}

std::vector<pivot::PivotTable*> Workbook::pivots_of_cache(std::uint32_t cache_id) {
  std::vector<pivot::PivotTable*> tables;
  for (Sheet& sheet : sheets_) {
    for (std::unique_ptr<pivot::PivotTable>& table : sheet.mutable_pivot_tables()) {
      if (table != nullptr && table->pivot_cache_id() == cache_id) {
        tables.push_back(table.get());
      }
    }
  }
  return tables;
}

void Workbook::invalidate_pivot_results(std::uint32_t cache_id) {
  // Group fields follow their base, so an edited cache re-derives them.
  for (std::unique_ptr<pivot::PivotCache>& cache : pivot_caches_) {
    if (cache != nullptr && cache->cache_id() == cache_id) {
      pivot::materialize_group_fields(*cache, date1904_);
      break;
    }
  }
  const std::vector<pivot::PivotTable*> tables = pivots_of_cache(cache_id);
  for (pivot::PivotTable* table : tables) {
    pivot::clear_derived_state(*table);
    table->set_projection_pending(true);
  }
  if (!tables.empty()) {
    mark_all_formulas_dirty();
  }
}

namespace {

Error unknown_cache(std::uint32_t cache_id) {
  return make_error(FormulonErrorCode::kNotFound, "cache_id does not name a pivot cache",
                    "cache_id=" + std::to_string(cache_id));
}

// Refuses an empty pivot name or one another pivot on `sheet` holds.
Expected<void, Error> check_pivot_name(const Sheet& sheet, std::string_view name, std::size_t self) {
  if (name.empty()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "name is empty");
  }
  if (pivot::pivot_name_taken(sheet, name, self)) {
    return make_error(FormulonErrorCode::kAlreadyExists, "name is already in use by another pivot",
                      "name=" + std::string(name));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<std::size_t, Error> Workbook::add_pivot_cache_field(std::uint32_t cache_id, pivot::PivotCacheField field) {
  pivot::PivotCache* cache = nullptr;
  for (std::unique_ptr<pivot::PivotCache>& c : pivot_caches_) {
    if (c != nullptr && c->cache_id() == cache_id) {
      cache = c.get();
      break;
    }
  }
  if (cache == nullptr) {
    return unknown_cache(cache_id);
  }
  cache->mutable_fields().push_back(std::move(field));
  pivot::sync_table_fields(*cache, pivots_of_cache(cache_id));
  return cache->fields().size() - 1U;
}

Expected<std::size_t, Error> Workbook::add_pivot(std::size_t sheet, std::unique_ptr<pivot::PivotTable> table) {
  if (sheet >= sheet_count() || table == nullptr) {
    return make_error(FormulonErrorCode::kInvalidArgument, "sheet_index is out of range or the pivot is missing",
                      "sheet=" + std::to_string(sheet));
  }
  Sheet& target = this->sheet(sheet);
  RETURN_IF_ERROR(check_pivot_name(target, table->name(), target.pivot_tables().size()));
  pivot::PivotCache* cache = nullptr;
  for (std::unique_ptr<pivot::PivotCache>& c : pivot_caches_) {
    if (c != nullptr && c->cache_id() == table->pivot_cache_id()) {
      cache = c.get();
      break;
    }
  }
  if (cache == nullptr) {
    return unknown_cache(table->pivot_cache_id());
  }
  pivot::sync_table_fields(*cache, {table.get()});
  pivot::share_axis_fields(*cache, *table);
  {
    const eval::ScopedEvalProfile scope(excel_profile());
    pivot::list_axis_field_items(*table, *cache);
  }
  // Creation order: one past every pivot of the workbook, so it stays monotonic after removals.
  std::uint64_t order = 0;
  for (const Sheet& each : sheets_) {
    for (const std::unique_ptr<pivot::PivotTable>& other : each.pivot_tables()) {
      order = other != nullptr && other->creation_order() > order ? other->creation_order() : order;
    }
  }
  table->set_creation_order(order + 1U);
  target.add_pivot_table(std::move(table));
  return target.pivot_tables().size() - 1U;
}

Expected<void, Error> Workbook::rename_pivot(std::size_t sheet, std::size_t pivot, std::string_view name) {
  if (sheet >= sheet_count()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "sheet_index is out of range",
                      "sheet=" + std::to_string(sheet));
  }
  Sheet& target = this->sheet(sheet);
  if (pivot >= target.pivot_tables().size()) {
    return make_error(FormulonErrorCode::kIndexOutOfRange, "pivot_index is out of range",
                      "pivot=" + std::to_string(pivot));
  }
  RETURN_IF_ERROR(check_pivot_name(target, name, pivot));
  target.mutable_pivot_tables()[pivot]->set_name(std::string(name));
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
