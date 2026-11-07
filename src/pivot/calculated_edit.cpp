//
// Calculated-field and calculated-item structural edits: adding, removing and
// reformulating them in the cache and keeping the dependent tables in step.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "pivot/calculated.h"
#include "pivot/calculated_internal.h"
#include "pivot/field_lookup.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_formula_host.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon::pivot {
using namespace detail;

namespace {

Error invalid(std::string message, std::string_view context = {}) {
  return make_error(FormulonErrorCode::kInvalidArgument, std::move(message), std::string(context));
}

// Same wording and context as the C API's collection-index check.
Error out_of_range(std::string_view field, std::size_t index, std::size_t count) {
  const std::string i = std::to_string(index);
  const std::string n = std::to_string(count);
  return make_error(FormulonErrorCode::kIndexOutOfRange,
                    std::string(field) + " " + i + " is out of range (" + n + " items)",
                    std::string(field) + "=" + i + " count=" + n);
}

Error precondition(std::string message) {
  return make_error(FormulonErrorCode::kPreconditionFailed, std::move(message));
}

bool same_name(std::string_view a, std::string_view b) {
  return same_value(Value::text(a), Value::text(b));
}

Expected<void, Error> check_name(std::string_view name) {
  if (name.empty()) {
    return invalid("calculated name is empty");
  }
  if (name.find('[') != std::string_view::npos || name.find(']') != std::string_view::npos) {
    return invalid("calculated name contains a bracket", name);
  }
  if (same_name(name, "All") || same_name(name, "Blank")) {
    return invalid("calculated name is reserved", name);
  }
  return Expected<void, Error>::Ok();
}

/// The references of `formula`, compiled in `cache` for a calculated field
/// (no `item_field`) or an item of `item_field`. A non-null `canonical`
/// receives the form Excel stores.
Expected<std::vector<FormulaRef>, Error> compile_refs(PivotFormulaHost& host, const PivotCache& cache,
                                                      std::optional<std::size_t> item_field, std::string_view formula,
                                                      std::string* canonical = nullptr) {
  std::vector<FormulaRef> refs;
  RETURN_IF_ERROR(host.compile(without_equals(formula), FormulaScope{&cache, item_field}, &refs, canonical));
  return refs;
}

/// The entry defining shared item `item` of field `field`.
std::optional<std::size_t> entry_of_item(const PivotCache& cache, std::size_t field, std::size_t item) {
  const std::vector<PivotCalculatedItem>& entries = cache.calculated_items();
  for (std::size_t e = 0; e < entries.size(); ++e) {
    if (entries[e].cache_field_index == field && entry_item(entries[e]) == item) {
      return e;
    }
  }
  return std::nullopt;
}

/// True when following calculated fields from `refs` reaches `target`.
bool field_refs_reach(PivotFormulaHost& host, const PivotCache& cache, const std::vector<FormulaRef>& refs,
                      std::size_t target) {
  std::vector<bool> seen(cache.fields().size(), false);
  std::vector<std::size_t> pending;
  for (const FormulaRef& ref : refs) {
    if (ref.kind == FormulaRef::Kind::Field) {
      pending.push_back(ref.field);
    }
  }
  while (!pending.empty()) {
    const std::size_t field = pending.back();
    pending.pop_back();
    if (field == target) {
      return true;
    }
    if (field >= seen.size() || seen[field] || !cache.fields()[field].formula.has_value()) {
      continue;
    }
    seen[field] = true;
    Expected<std::vector<FormulaRef>, Error> next =
        compile_refs(host, cache, std::nullopt, *cache.fields()[field].formula);
    if (next) {
      for (const FormulaRef& ref : next.value()) {
        if (ref.kind == FormulaRef::Kind::Field) {
          pending.push_back(ref.field);
        }
      }
    }
  }
  return false;
}

/// True when following calculated items of `field` by name from `refs`
/// reaches shared item `target`.
bool item_refs_reach(PivotFormulaHost& host, const PivotCache& cache, std::size_t field,
                     const std::vector<FormulaRef>& refs, std::size_t target) {
  const PivotCacheField& cache_field = cache.fields()[field];
  std::vector<bool> seen(cache_field.shared_items.size(), false);
  std::vector<std::size_t> pending;
  for (const FormulaRef& ref : refs) {
    if (ref.kind == FormulaRef::Kind::Item) {
      pending.push_back(ref.item);
    }
  }
  while (!pending.empty()) {
    const std::size_t item = pending.back();
    pending.pop_back();
    if (item == target) {
      return true;
    }
    if (item >= seen.size() || seen[item]) {
      continue;
    }
    seen[item] = true;
    const std::optional<std::size_t> entry = entry_of_item(cache, field, item);
    if (!entry.has_value()) {
      continue;
    }
    Expected<std::vector<FormulaRef>, Error> next =
        compile_refs(host, cache, field, cache.calculated_items()[*entry].formula);
    if (next) {
      for (const FormulaRef& ref : next.value()) {
        if (ref.kind == FormulaRef::Kind::Item) {
          pending.push_back(ref.item);
        }
      }
    }
  }
  return false;
}

template <class T>
void erase_at(std::vector<T>& values, std::size_t index) {
  if (index < values.size()) {
    values.erase(values.begin() + static_cast<std::ptrdiff_t>(index));
  }
}

bool on_axis(const PivotField& field) {
  return field.axis == PivotAxis::Row || field.axis == PivotAxis::Col || field.axis == PivotAxis::Page;
}

/// Position of `field` on its axis in `table`, for the `<pivotArea>` Excel
/// writes with a calculated item.
std::optional<std::uint32_t> axis_position(const PivotTable& table, std::size_t field) {
  for (const std::vector<std::uint32_t>& order :
       {table.row_field_order(), table.col_field_order(), table.page_field_order()}) {
    if (const auto position = field_position_in_order(order, field)) {
      return static_cast<std::uint32_t>(*position);
    }
  }
  return std::nullopt;
}

/// Renumbers field indices after cache field `removed` is erased; `true` when
/// `index` named it.
bool renumber_field(std::uint32_t& index, std::size_t removed) {
  if (index == removed) {
    return true;
  }
  if (index > removed) {
    --index;
  }
  return false;
}

void erase_field_from_table(PivotTable& table, std::size_t field) {
  std::vector<PivotDataField>& data_fields = table.mutable_data_fields();
  std::vector<std::size_t> removed_data_fields;
  std::vector<PivotDataField> kept;
  for (std::size_t d = 0; d < data_fields.size(); ++d) {
    PivotDataField data_field = std::move(data_fields[d]);
    if (renumber_field(data_field.field_index, field)) {
      removed_data_fields.push_back(d);
      continue;
    }
    if (data_field.show_as_base_field.has_value() && renumber_field(*data_field.show_as_base_field, field)) {
      data_field.show_as_base_field.reset();
    }
    kept.push_back(std::move(data_field));
  }
  data_fields = std::move(kept);

  for (PivotFilter& filter : table.mutable_active_filters()) {
    std::uint32_t moved = 0;
    for (const std::size_t d : removed_data_fields) {
      if (filter.data_field_index > d) {
        ++moved;
      }
    }
    filter.data_field_index -= moved;
    if (filter.source.has_value() && filter.source->field_index > field) {
      filter.source.reset();
    }
  }

  for (std::vector<std::uint32_t>* order : {&table.mutable_row_field_order(), &table.mutable_col_field_order()}) {
    std::vector<std::uint32_t> renumbered;
    for (std::uint32_t index : *order) {
      if (!renumber_field(index, field)) {
        renumbered.push_back(index);
      }
    }
    *order = std::move(renumbered);
  }
  std::vector<PivotPageField> pages;
  for (PivotPageField page : table.mutable_page_fields()) {
    if (!renumber_field(page.field_index, field)) {
      pages.push_back(page);
    }
  }
  table.mutable_page_fields() = std::move(pages);

  erase_at(table.mutable_fields(), field);
}

}  // namespace

void sync_table_fields(const PivotCache& cache, const std::vector<PivotTable*>& tables) {
  for (PivotTable* table : tables) {
    std::vector<PivotField>& fields = table->mutable_fields();
    while (fields.size() < cache.fields().size()) {
      const PivotCacheField& cache_field = cache.fields()[fields.size()];
      PivotField field;
      field.source_name = cache_field.name;
      field.axis = PivotAxis::None;
      if (cache_field.formula.has_value()) {
        field.default_subtotal = false;
        field.passthrough_attrs = {{"dragToRow", "0"}, {"dragToCol", "0"}, {"dragToPage", "0"}};
      }
      fields.push_back(std::move(field));
    }
  }
}

Expected<std::size_t, Error> add_calculated_field(PivotCache& cache, const std::vector<PivotTable*>& tables,
                                                  std::string_view name, std::string_view formula,
                                                  PivotFormulaHost& host) {
  RETURN_IF_ERROR(check_name(name));
  for (const PivotCacheField& field : cache.fields()) {
    if (same_name(field.name, name)) {
      return invalid("calculated field name is already a field", name);
    }
  }
  for (const PivotTable* table : tables) {
    for (const PivotField& field : table->fields()) {
      if (!field.custom_name.empty() && same_name(field.custom_name, name)) {
        return invalid("calculated field name is already a field", name);
      }
    }
  }
  RETURN_IF_ERROR(compile_refs(host, cache, std::nullopt, formula));
  PivotCacheField field;
  field.name = std::string(name);
  field.num_fmt_id = 0U;
  field.is_database_field = false;
  field.formula = std::string(without_equals(formula));
  cache.mutable_fields().push_back(std::move(field));
  sync_table_fields(cache, tables);
  return cache.fields().size() - 1U;
}

Expected<void, Error> set_calculated_field_formula(PivotCache& cache, std::size_t field_index, std::string_view formula,
                                                   PivotFormulaHost& host) {
  if (field_index >= cache.fields().size() || !cache.fields()[field_index].formula.has_value()) {
    return invalid("not a calculated field");
  }
  ASSIGN_OR_RETURN(std::vector<FormulaRef> refs, compile_refs(host, cache, std::nullopt, formula));
  if (field_refs_reach(host, cache, refs, field_index)) {
    return invalid("calculated field refers to itself", formula);
  }
  cache.mutable_fields()[field_index].formula = std::string(without_equals(formula));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> remove_calculated_field(PivotCache& cache, const std::vector<PivotTable*>& tables,
                                              std::size_t field_index, PivotFormulaHost& host) {
  if (field_index >= cache.fields().size() || !cache.fields()[field_index].formula.has_value()) {
    return invalid("not a calculated field");
  }
  for (std::size_t f = 0; f < cache.fields().size(); ++f) {
    if (f == field_index || !cache.fields()[f].formula.has_value()) {
      continue;
    }
    Expected<std::vector<FormulaRef>, Error> refs = compile_refs(host, cache, std::nullopt, *cache.fields()[f].formula);
    if (!refs) {
      continue;
    }
    for (const FormulaRef& ref : refs.value()) {
      if (ref.kind == FormulaRef::Kind::Field && ref.field == field_index) {
        return precondition("calculated field is used by " + cache.fields()[f].name);
      }
    }
  }

  for (PivotTable* table : tables) {
    erase_field_from_table(*table, field_index);
  }
  std::vector<PivotCacheField>& fields = cache.mutable_fields();
  erase_at(fields, field_index);
  for (PivotCacheField& field : fields) {
    if (field.field_group.has_value()) {
      for (std::optional<std::uint32_t>* link : {&field.field_group->par, &field.field_group->base}) {
        if (link->has_value() && renumber_field(**link, field_index)) {
          link->reset();
        }
      }
    }
  }
  for (PivotCacheRecord& record : cache.mutable_records()) {
    erase_at(record.cells, field_index);
    erase_at(record.cell_is_index, field_index);
    erase_at(record.cell_is_date, field_index);
    erase_at(record.cell_text_slot, field_index);
  }
  for (PivotCalculatedItem& entry : cache.mutable_calculated_items()) {
    renumber_field(entry.cache_field_index, field_index);
    for (PivotAreaReference& reference : entry.area.references) {
      if (reference.field.has_value()) {
        renumber_field(*reference.field, field_index);
      }
    }
  }
  return Expected<void, Error>::Ok();
}

bool aggregation_excludes_calculated_items(Aggregation aggregation) {
  switch (aggregation) {
    case Aggregation::Average:
    case Aggregation::StdDev:
    case Aggregation::StdDevP:
    case Aggregation::Var:
    case Aggregation::VarP:
      return true;
    default:
      return false;
  }
}

Expected<void, Error> check_aggregation_with_calculated_items(const PivotCache& cache, Aggregation aggregation) {
  if (!cache.calculated_items().empty() && aggregation_excludes_calculated_items(aggregation)) {
    return precondition("the function cannot summarise a pivot with calculated items");
  }
  return Expected<void, Error>::Ok();
}

Expected<std::size_t, Error> add_calculated_item(PivotCache& cache, const std::vector<PivotTable*>& tables,
                                                 std::size_t field_index, std::string_view name,
                                                 std::string_view formula, PivotFormulaHost& host) {
  if (field_index >= cache.fields().size()) {
    return out_of_range("field_index", field_index, cache.fields().size());
  }
  const PivotCacheField& cache_field = cache.fields()[field_index];
  if (cache_field.formula.has_value() || cache_field.field_group.has_value()) {
    return precondition("calculated items need an ungrouped source field");
  }
  std::optional<std::uint32_t> position;
  for (const PivotTable* table : tables) {
    if (field_index < table->fields().size() && on_axis(table->fields()[field_index]) && !position.has_value()) {
      position = axis_position(*table, field_index).value_or(0U);
    }
    for (const PivotDataField& data_field : table->data_fields()) {
      if (aggregation_excludes_calculated_items(data_field.aggregation)) {
        return precondition("the function cannot summarise a pivot with calculated items");
      }
    }
  }
  if (!position.has_value()) {
    return precondition("calculated item field is on no axis");
  }
  RETURN_IF_ERROR(check_name(name));
  for (const Value& item : cache_field.shared_items) {
    if (same_name(display_string(item), name)) {
      return invalid("calculated item name is already an item", name);
    }
  }
  std::string canonical;
  RETURN_IF_ERROR(compile_refs(host, cache, field_index, formula, &canonical));

  PivotCacheField& field = cache.mutable_fields()[field_index];
  const auto item = static_cast<std::uint32_t>(field.shared_items.size());
  cache.mutable_text_storage().emplace_back(name);
  field.shared_items.push_back(Value::text(cache.mutable_text_storage().back()));
  field.shared_item_is_calculated.resize(item, false);
  field.shared_item_is_calculated.push_back(true);

  PivotCalculatedItem entry;
  entry.cache_field_index = static_cast<std::uint32_t>(field_index);
  entry.formula = std::move(canonical);
  entry.area.cache_index = true;
  entry.area.outline = false;
  entry.area.field_position = position;
  PivotAreaReference reference;
  reference.field = static_cast<std::uint32_t>(field_index);
  reference.items.push_back(item);
  entry.area.references.push_back(std::move(reference));
  cache.mutable_calculated_items().push_back(std::move(entry));

  for (PivotTable* table : tables) {
    if (field_index >= table->fields().size()) {
      continue;
    }
    PivotField& pivot_field = table->mutable_fields()[field_index];
    if (pivot_field.items.empty()) {
      if (!on_axis(pivot_field)) {
        continue;
      }
      // Materialise the implicit items so the new one lands after them.
      for (std::uint32_t k = 0; k < item; ++k) {
        PivotItem existing;
        existing.name = display_string(field.shared_items[k]);
        existing.has_cache_index = true;
        existing.cache_index = k;
        pivot_field.items.push_back(std::move(existing));
      }
    }
    PivotItem calculated;
    calculated.name = std::string(name);
    calculated.has_cache_index = true;
    calculated.cache_index = item;
    calculated.calculated = true;
    pivot_field.items.push_back(std::move(calculated));
  }
  return cache.calculated_items().size() - 1U;
}

Expected<void, Error> set_calculated_item_formula(PivotCache& cache, std::size_t index, std::string_view formula,
                                                  PivotFormulaHost& host) {
  if (index >= cache.calculated_items().size()) {
    return out_of_range("index", index, cache.calculated_items().size());
  }
  const PivotCalculatedItem& entry = cache.calculated_items()[index];
  const std::size_t field = entry.cache_field_index;
  std::string canonical;
  ASSIGN_OR_RETURN(std::vector<FormulaRef> refs, compile_refs(host, cache, field, formula, &canonical));
  if (const std::optional<std::uint32_t> item = entry_item(entry);
      item.has_value() && item_refs_reach(host, cache, field, refs, *item)) {
    return invalid("calculated item refers to itself", formula);
  }
  cache.mutable_calculated_items()[index].formula = std::move(canonical);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> move_calculated_item(PivotCache& cache, std::size_t index, std::size_t new_index) {
  std::vector<PivotCalculatedItem>& entries = cache.mutable_calculated_items();
  if (index >= entries.size()) {
    return out_of_range("index", index, entries.size());
  }
  if (new_index >= entries.size()) {
    return out_of_range("new_index", new_index, entries.size());
  }
  PivotCalculatedItem entry = std::move(entries[index]);
  entries.erase(entries.begin() + static_cast<std::ptrdiff_t>(index));
  entries.insert(entries.begin() + static_cast<std::ptrdiff_t>(new_index), std::move(entry));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> remove_calculated_item(PivotCache& cache, const std::vector<PivotTable*>& tables,
                                             std::size_t index, PivotFormulaHost& host) {
  if (index >= cache.calculated_items().size()) {
    return out_of_range("index", index, cache.calculated_items().size());
  }
  const std::size_t field = cache.calculated_items()[index].cache_field_index;
  const std::optional<std::uint32_t> item = entry_item(cache.calculated_items()[index]);
  if (item.has_value()) {
    for (std::size_t e = 0; e < cache.calculated_items().size(); ++e) {
      const PivotCalculatedItem& other = cache.calculated_items()[e];
      if (e == index || other.cache_field_index != field) {
        continue;
      }
      Expected<std::vector<FormulaRef>, Error> refs = compile_refs(host, cache, field, other.formula);
      if (!refs) {
        continue;
      }
      for (const FormulaRef& ref : refs.value()) {
        if (ref.kind == FormulaRef::Kind::Item && ref.item == *item) {
          return precondition("calculated item is used by another calculated item");
        }
      }
    }
  }

  std::vector<PivotCalculatedItem>& entries = cache.mutable_calculated_items();
  erase_at(entries, index);
  if (!item.has_value() || field >= cache.fields().size()) {
    return Expected<void, Error>::Ok();
  }
  const std::uint32_t removed = *item;
  PivotCacheField& cache_field = cache.mutable_fields()[field];
  erase_at(cache_field.shared_items, removed);
  erase_at(cache_field.shared_item_is_calculated, removed);
  erase_at(cache_field.shared_item_is_date, removed);
  for (PivotCacheRecord& record : cache.mutable_records()) {
    if (field >= record.cells.size() || !record.cells[field].is_number()) {
      continue;
    }
    const bool is_index =
        field < record.cell_is_index.size() ? record.cell_is_index[field] : !cache_field.shared_items.empty();
    const double index_value = record.cells[field].as_number();
    if (is_index && index_value > static_cast<double>(removed)) {
      record.cells[field] = Value::number(index_value - 1.0);
    }
  }
  for (PivotCalculatedItem& entry : entries) {
    for (PivotAreaReference& reference : entry.area.references) {
      if (reference.field != field) {
        continue;
      }
      std::vector<std::uint32_t> items;
      for (std::uint32_t x : reference.items) {
        if (!renumber_field(x, removed)) {
          items.push_back(x);
        }
      }
      reference.items = std::move(items);
    }
  }
  for (PivotTable* table : tables) {
    if (field >= table->fields().size()) {
      continue;
    }
    std::vector<PivotItem>& items = table->mutable_fields()[field].items;
    std::optional<std::size_t> position;
    for (std::size_t p = 0; p < items.size(); ++p) {
      if (items[p].has_cache_index && items[p].cache_index == removed && !position.has_value()) {
        position = p;
      } else if (items[p].has_cache_index && items[p].cache_index > removed) {
        --items[p].cache_index;
      }
    }
    if (!position.has_value()) {
      continue;
    }
    erase_at(items, *position);
    for (PivotPageField& page : table->mutable_page_fields()) {
      if (page.field_index != field || !page.item_index.has_value()) {
        continue;
      }
      if (*page.item_index == *position) {
        page.item_index.reset();
      } else if (*page.item_index > *position) {
        --*page.item_index;
      }
    }
  }
  return Expected<void, Error>::Ok();
}

}  // namespace formulon::pivot
