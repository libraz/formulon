//
// Calculated-field and calculated-item model operations and evaluation.

#include "pivot/calculated.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "pivot/hierarchy_builder.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_formula_host.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// Formulas nested deeper than this, a cycle included, evaluate to #REF!.
constexpr std::size_t kMaxCalculatedDepth = 64;

Error invalid(std::string message, std::string_view context = {}) {
  return make_error(FormulonErrorCode::kInvalidArgument, std::move(message), std::string(context));
}

Error precondition(std::string message) {
  return make_error(FormulonErrorCode::kPreconditionFailed, std::move(message));
}

bool same_value(const Value& a, const Value& b) {
  return !value_less(a, b) && !value_less(b, a);
}

bool same_name(std::string_view a, std::string_view b) {
  return same_value(Value::text(a), Value::text(b));
}

std::string_view without_equals(std::string_view formula) {
  if (!formula.empty() && formula.front() == '=') {
    formula.remove_prefix(1);
  }
  return formula;
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
  Expected<std::uint32_t, Error> id =
      host.compile(without_equals(formula), FormulaScope{&cache, item_field}, &refs, canonical);
  if (!id) {
    return id.error();
  }
  return refs;
}

/// The shared item a `<calculatedItem>` entry defines: the first item its area
/// selects on its own field.
std::optional<std::uint32_t> entry_item(const PivotCalculatedItem& entry) {
  if (!entry.has_area) {
    return std::nullopt;
  }
  for (const PivotAreaReference& reference : entry.area.references) {
    if (reference.field == entry.cache_field_index && !reference.items.empty()) {
      return reference.items.front();
    }
  }
  return std::nullopt;
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

bool on_axis(const PivotField& field) {
  return field.axis == PivotAxis::Row || field.axis == PivotAxis::Col || field.axis == PivotAxis::Page;
}

/// Position of `field` on its axis in `table`, for the `<pivotArea>` Excel
/// writes with a calculated item.
std::optional<std::uint32_t> axis_position(const PivotTable& table, std::size_t field) {
  for (const std::vector<std::uint32_t>& order :
       {table.row_field_order(), table.col_field_order(), table.page_field_order()}) {
    const auto it = std::find(order.begin(), order.end(), static_cast<std::uint32_t>(field));
    if (it != order.end()) {
      return static_cast<std::uint32_t>(it - order.begin());
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

  if (field < table.fields().size()) {
    table.mutable_fields().erase(table.mutable_fields().begin() + static_cast<std::ptrdiff_t>(field));
  }
}

template <class T>
void erase_at(std::vector<T>& values, std::size_t index) {
  if (index < values.size()) {
    values.erase(values.begin() + static_cast<std::ptrdiff_t>(index));
  }
}

/// Every per-data-field vector of `result`.
template <class Fn>
void for_each_slot_vector(PivotResult& result, Fn&& fn) {
  for (std::vector<std::vector<Value>>& row : result.values) {
    for (std::vector<Value>& cell : row) {
      fn(cell);
    }
  }
  for (std::vector<Value>& slots : result.subtotals) {
    fn(slots);
  }
  for (RowSubtotal& subtotal : result.row_subtotals) {
    fn(subtotal.values);
    for (std::vector<Value>& slots : subtotal.col_values) {
      fn(slots);
    }
    for (std::vector<Value>& slots : subtotal.col_subtotal_values) {
      fn(slots);
    }
  }
  for (ColSubtotal& subtotal : result.col_subtotals) {
    for (std::vector<Value>& slots : subtotal.values) {
      fn(slots);
    }
    fn(subtotal.total);
  }
  fn(result.grand_totals);
  for (std::vector<Value>& slots : result.row_leaf_totals) {
    fn(slots);
  }
  for (std::vector<Value>& slots : result.col_leaf_totals) {
    fn(slots);
  }
}

/// Copies a text result into `result`, which outlives the host's buffer.
Value keep(const Value& value, PivotResult& result) {
  if (!value.is_text()) {
    return value;
  }
  result.text_storage.emplace_back(value.as_text());
  return Value::text(result.text_storage.back());
}

struct PathLess {
  bool operator()(const std::vector<Value>& a, const std::vector<Value>& b) const {
    return std::lexicographical_compare(a.begin(), a.end(), b.begin(), b.end(), value_less);
  }
};

using PathIndex = std::map<std::vector<Value>, std::size_t, PathLess>;

PathIndex index_paths(const AxisLeafPaths& axis) {
  PathIndex index;
  for (std::size_t leaf = 0; leaf < axis.paths.size(); ++leaf) {
    std::vector<Value> keys;
    keys.reserve(axis.paths[leaf].size());
    for (const LeafPathStep& step : axis.paths[leaf]) {
      keys.push_back(step.key);
    }
    index.emplace(std::move(keys), leaf);
  }
  return index;
}

std::optional<std::size_t> find_path(const PathIndex& index, const std::vector<LeafPathStep>& path) {
  std::vector<Value> keys;
  keys.reserve(path.size());
  for (const LeafPathStep& step : path) {
    keys.push_back(step.key);
  }
  const auto it = index.find(keys);
  return it == index.end() ? std::nullopt : std::optional<std::size_t>(it->second);
}

const LeafPathStep* find_step(std::size_t field, const std::vector<LeafPathStep>& row,
                              const std::vector<LeafPathStep>& col) {
  for (const std::vector<LeafPathStep>* path : {&row, &col}) {
    for (const LeafPathStep& step : *path) {
      if (step.field_index == field) {
        return &step;
      }
    }
  }
  return nullptr;
}

bool is_calculated_key(const PivotCache& cache, std::size_t field, const Value& key) {
  if (field >= cache.fields().size()) {
    return false;
  }
  const PivotCacheField& cache_field = cache.fields()[field];
  for (std::size_t k = 0; k < cache_field.shared_item_is_calculated.size() && k < cache_field.shared_items.size();
       ++k) {
    if (cache_field.shared_item_is_calculated[k] && same_value(cache_field.shared_items[k], key)) {
      return true;
    }
  }
  return false;
}

bool path_on_calculated_item(const PivotCache& cache, const std::vector<LeafPathStep>& path) {
  return std::any_of(path.begin(), path.end(),
                     [&](const LeafPathStep& step) { return is_calculated_key(cache, step.field_index, step.key); });
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
  if (Expected<void, Error> checked = check_name(name); !checked) {
    return checked.error();
  }
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
  if (Expected<std::vector<FormulaRef>, Error> refs = compile_refs(host, cache, std::nullopt, formula); !refs) {
    return refs.error();
  }
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
  Expected<std::vector<FormulaRef>, Error> refs = compile_refs(host, cache, std::nullopt, formula);
  if (!refs) {
    return refs.error();
  }
  if (field_refs_reach(host, cache, refs.value(), field_index)) {
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
  fields.erase(fields.begin() + static_cast<std::ptrdiff_t>(field_index));
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
    return invalid("calculated item field out of range");
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
  if (Expected<void, Error> checked = check_name(name); !checked) {
    return checked.error();
  }
  for (const Value& item : cache_field.shared_items) {
    if (same_name(display_string(item), name)) {
      return invalid("calculated item name is already an item", name);
    }
  }
  std::string canonical;
  if (Expected<std::vector<FormulaRef>, Error> refs = compile_refs(host, cache, field_index, formula, &canonical);
      !refs) {
    return refs.error();
  }

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
    return invalid("calculated item out of range");
  }
  const PivotCalculatedItem& entry = cache.calculated_items()[index];
  const std::size_t field = entry.cache_field_index;
  std::string canonical;
  Expected<std::vector<FormulaRef>, Error> refs = compile_refs(host, cache, field, formula, &canonical);
  if (!refs) {
    return refs.error();
  }
  if (const std::optional<std::uint32_t> item = entry_item(entry);
      item.has_value() && item_refs_reach(host, cache, field, refs.value(), *item)) {
    return invalid("calculated item refers to itself", formula);
  }
  cache.mutable_calculated_items()[index].formula = std::move(canonical);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> move_calculated_item(PivotCache& cache, std::size_t index, std::size_t new_index) {
  std::vector<PivotCalculatedItem>& entries = cache.mutable_calculated_items();
  if (index >= entries.size() || new_index >= entries.size()) {
    return invalid("calculated item out of range");
  }
  PivotCalculatedItem entry = std::move(entries[index]);
  entries.erase(entries.begin() + static_cast<std::ptrdiff_t>(index));
  entries.insert(entries.begin() + static_cast<std::ptrdiff_t>(new_index), std::move(entry));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> remove_calculated_item(PivotCache& cache, const std::vector<PivotTable*>& tables,
                                             std::size_t index, PivotFormulaHost& host) {
  if (index >= cache.calculated_items().size()) {
    return invalid("calculated item out of range");
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
  entries.erase(entries.begin() + static_cast<std::ptrdiff_t>(index));
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
    items.erase(items.begin() + static_cast<std::ptrdiff_t>(*position));
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

struct CalculatedPlan::ItemContext {
  const AxisLeafPaths& rows;
  const AxisLeafPaths& cols;
  PathIndex row_index;
  PathIndex col_index;
  std::vector<std::vector<std::vector<Value>>> base;
  PivotResult& result;
};

struct CalculatedPlan::Inputs {
  bool any = false;
  bool non_blank = false;
};

CalculatedPlan::CalculatedPlan(const PivotTable& table, const PivotCache& cache, PivotFormulaHost* host)
    : cache_(&cache), host_(host), visible_slots_(table.data_fields().size()) {
  fields_.resize(cache.fields().size());
  for (const PivotDataField& data_field : table.data_fields()) {
    slot_fields_.push_back(data_field.field_index);
  }
  if (host_ != nullptr) {
    std::vector<std::uint8_t> state(cache.fields().size(), 0U);
    for (const std::uint32_t field : slot_fields_) {
      if (field < cache.fields().size() && cache.fields()[field].formula.has_value()) {
        compile_field(field, state);
      }
    }
    for (const PivotCalculatedItem& entry : cache.calculated_items()) {
      Compiled compiled;
      if (entry.cache_field_index < cache.fields().size()) {
        Expected<std::uint32_t, Error> id = host_->compile(
            without_equals(entry.formula), FormulaScope{&cache, entry.cache_field_index}, &compiled.refs, nullptr);
        if (id) {
          compiled.ok = true;
          compiled.id = id.value();
        }
      }
      items_.push_back(std::move(compiled));
    }
  }
  calculated_slots_.assign(visible_slots_ + sum_fields_.size(), false);
  for (std::size_t slot = 0; slot < visible_slots_; ++slot) {
    const std::uint32_t field = slot_fields_[slot];
    calculated_slots_[slot] = field < cache.fields().size() && cache.fields()[field].formula.has_value();
  }
}

void CalculatedPlan::compile_field(std::size_t field_index, std::vector<std::uint8_t>& state) {
  if (state[field_index] != 0U) {
    return;
  }
  state[field_index] = 1U;
  Compiled compiled;
  Expected<std::uint32_t, Error> id = host_->compile(without_equals(*cache_->fields()[field_index].formula),
                                                     FormulaScope{cache_, std::nullopt}, &compiled.refs, nullptr);
  if (id) {
    compiled.ok = true;
    compiled.id = id.value();
  }
  for (const FormulaRef& ref : compiled.refs) {
    if (ref.kind != FormulaRef::Kind::Field || ref.field >= cache_->fields().size()) {
      continue;
    }
    if (cache_->fields()[ref.field].formula.has_value()) {
      if (state[ref.field] == 1U) {
        compiled.ok = false;
      } else {
        compile_field(ref.field, state);
      }
    } else if (std::find(sum_fields_.begin(), sum_fields_.end(), ref.field) == sum_fields_.end()) {
      sum_fields_.push_back(static_cast<std::uint32_t>(ref.field));
    }
  }
  state[field_index] = 2U;
  fields_[field_index] = std::move(compiled);
}

Value CalculatedPlan::field_value(std::size_t field_index, const std::vector<Value>& slots, std::size_t depth,
                                  Inputs& inputs, PivotResult& result) const {
  if (depth > kMaxCalculatedDepth) {
    return Value::error(ErrorCode::Ref);
  }
  if (!cache_->fields()[field_index].formula.has_value()) {
    const auto it = std::find(sum_fields_.begin(), sum_fields_.end(), field_index);
    const std::size_t slot = visible_slots_ + static_cast<std::size_t>(it - sum_fields_.begin());
    inputs.any = true;
    if (it == sum_fields_.end() || slot >= slots.size()) {
      return Value::blank();
    }
    inputs.non_blank = inputs.non_blank || !slots[slot].is_blank();
    return slots[slot];
  }
  const std::optional<Compiled>& compiled = fields_[field_index];
  if (!compiled.has_value() || !compiled->ok) {
    return Value::error(ErrorCode::Ref);
  }
  std::vector<Value> args;
  args.reserve(compiled->refs.size());
  for (const FormulaRef& ref : compiled->refs) {
    args.push_back(ref.kind == FormulaRef::Kind::Field && ref.field < cache_->fields().size()
                       ? field_value(ref.field, slots, depth + 1U, inputs, result)
                       : Value::error(ErrorCode::Ref));
  }
  return keep(host_->evaluate(compiled->id, args.data(), args.size()), result);
}

void CalculatedPlan::fill_field_slots(std::vector<Value>& slots, PivotResult& result) const {
  if (host_ == nullptr) {
    return;
  }
  for (std::size_t slot = 0; slot < visible_slots_ && slot < slots.size(); ++slot) {
    if (!calculated_slots_[slot]) {
      continue;
    }
    Inputs inputs;
    const Value value = field_value(slot_fields_[slot], slots, 0U, inputs, result);
    // A cell none of whose summed fields has data stays empty.
    slots[slot] = inputs.any && !inputs.non_blank ? Value::blank() : value;
  }
}

void CalculatedPlan::fill_field_slots(PivotResult& result) const {
  if (host_ == nullptr ||
      std::find(calculated_slots_.begin(), calculated_slots_.end(), true) == calculated_slots_.end()) {
    return;
  }
  for_each_slot_vector(result, [&](std::vector<Value>& slots) { fill_field_slots(slots, result); });
  if (!result.grand_totals.empty()) {
    result.grand_total = result.grand_totals[0];
  }
}

Value CalculatedPlan::item_sort_key(std::size_t field_index, const Value& key,
                                    const std::map<Value, Value, ValueLess>& sibling_keys) const {
  if (host_ == nullptr || field_index >= cache_->fields().size()) {
    return Value::blank();
  }
  const std::vector<Value>& shared = cache_->fields()[field_index].shared_items;
  std::optional<std::size_t> winner;
  for (std::size_t e = 0; e < cache_->calculated_items().size(); ++e) {
    const PivotCalculatedItem& entry = cache_->calculated_items()[e];
    const std::optional<std::uint32_t> item = entry_item(entry);
    if (entry.cache_field_index == field_index && item.has_value() && *item < shared.size() &&
        same_value(shared[*item], key)) {
      winner = e;
    }
  }
  if (!winner.has_value() || !items_[*winner].ok) {
    return Value::blank();
  }
  const Compiled& compiled = items_[*winner];
  std::vector<Value> args;
  args.reserve(compiled.refs.size());
  for (const FormulaRef& ref : compiled.refs) {
    if (ref.kind != FormulaRef::Kind::Item || ref.field != field_index || ref.item >= shared.size()) {
      args.push_back(Value::error(ErrorCode::Ref));
      continue;
    }
    const auto it = sibling_keys.find(shared[ref.item]);
    args.push_back(it == sibling_keys.end() ? Value::blank() : it->second);
  }
  const Value value = host_->evaluate(compiled.id, args.data(), args.size());
  return value.is_text() ? Value::blank() : value;
}

void CalculatedPlan::drop_sum_slots(PivotResult& result) const {
  if (sum_fields_.empty()) {
    return;
  }
  for_each_slot_vector(result, [&](std::vector<Value>& slots) {
    if (slots.size() > visible_slots_) {
      slots.erase(slots.begin() + static_cast<std::ptrdiff_t>(visible_slots_), slots.end());
    }
  });
}

bool CalculatedPlan::entry_applies(std::size_t entry, const std::vector<LeafPathStep>& row,
                                   const std::vector<LeafPathStep>& col) const {
  const PivotCalculatedItem& item = cache_->calculated_items()[entry];
  if (!item.has_area) {
    return false;
  }
  bool own_field = false;
  for (const PivotAreaReference& reference : item.area.references) {
    if (!reference.field.has_value() || *reference.field >= cache_->fields().size()) {
      continue;
    }
    const LeafPathStep* step = find_step(*reference.field, row, col);
    if (step == nullptr) {
      return false;
    }
    own_field = own_field || *reference.field == item.cache_field_index;
    if (reference.items.empty()) {
      continue;
    }
    const std::vector<Value>& shared = cache_->fields()[*reference.field].shared_items;
    const bool selected = std::any_of(reference.items.begin(), reference.items.end(), [&](std::uint32_t x) {
      return x < shared.size() && same_value(shared[x], step->key);
    });
    if (!selected) {
      return false;
    }
  }
  return own_field;
}

Value CalculatedPlan::item_cell(const ItemContext& context, const std::vector<LeafPathStep>& row,
                                const std::vector<LeafPathStep>& col, std::size_t slot, std::size_t depth) const {
  if (depth > kMaxCalculatedDepth) {
    return Value::error(ErrorCode::Ref);
  }
  std::optional<std::size_t> winner;
  for (std::size_t e = 0; e < cache_->calculated_items().size(); ++e) {
    if (entry_applies(e, row, col)) {
      winner = e;
    }
  }
  if (!winner.has_value()) {
    const std::optional<std::size_t> r = find_path(context.row_index, row);
    const std::optional<std::size_t> c = find_path(context.col_index, col);
    if (!r.has_value() || !c.has_value() || *r >= context.base.size() || *c >= context.base[*r].size() ||
        slot >= context.base[*r][*c].size()) {
      return Value::blank();
    }
    return context.base[*r][*c][slot];
  }
  if (host_ == nullptr) {
    return Value::blank();
  }
  const Compiled& compiled = items_[*winner];
  if (!compiled.ok) {
    return Value::error(ErrorCode::Ref);
  }
  const std::size_t field = cache_->calculated_items()[*winner].cache_field_index;
  const auto step_of = [&](const std::vector<LeafPathStep>& path) {
    std::size_t i = 0;
    while (i < path.size() && path[i].field_index != field) {
      ++i;
    }
    return i;
  };
  // `entry_applies` guarantees the item's own field is on one of the paths.
  const bool on_row = step_of(row) < row.size();
  const std::size_t step_index = on_row ? step_of(row) : step_of(col);
  const std::vector<LeafPathStep>& path = on_row ? row : col;
  const LeafPathStep& step = path[step_index];
  const std::vector<Value>& siblings = (on_row ? context.rows : context.cols).sibling_keys[step.siblings];
  std::optional<std::size_t> position;
  for (std::size_t p = 0; p < siblings.size(); ++p) {
    if (same_value(siblings[p], step.key)) {
      position = p;
      break;
    }
  }
  const std::vector<Value>& shared = cache_->fields()[field].shared_items;

  std::vector<Value> args;
  args.reserve(compiled.refs.size());
  for (const FormulaRef& ref : compiled.refs) {
    std::optional<Value> key;
    switch (ref.kind) {
      case FormulaRef::Kind::Item:
        if (ref.field == field && ref.item < shared.size()) {
          key = shared[ref.item];
        }
        break;
      case FormulaRef::Kind::ItemAt:
        if (ref.position >= 1 && static_cast<std::size_t>(ref.position) <= siblings.size()) {
          key = siblings[static_cast<std::size_t>(ref.position) - 1U];
        }
        break;
      case FormulaRef::Kind::ItemOffset:
        if (position.has_value()) {
          const long long target = static_cast<long long>(*position) + ref.position;
          if (target >= 0 && static_cast<std::size_t>(target) < siblings.size()) {
            key = siblings[static_cast<std::size_t>(target)];
          }
        }
        break;
      case FormulaRef::Kind::Field:
        break;
    }
    if (!key.has_value()) {
      args.push_back(Value::error(ErrorCode::Ref));
      continue;
    }
    std::vector<LeafPathStep> moved_row = row;
    std::vector<LeafPathStep> moved_col = col;
    (on_row ? moved_row : moved_col)[step_index].key = *key;
    args.push_back(item_cell(context, moved_row, moved_col, slot, depth + 1U));
  }
  return keep(host_->evaluate(compiled.id, args.data(), args.size()), context.result);
}

bool CalculatedPlan::fill_item_cells(PivotResult& result, const AxisLeafPaths& rows, const AxisLeafPaths& cols) const {
  if (cache_->calculated_items().empty()) {
    return false;
  }
  std::vector<bool> row_calculated(rows.paths.size());
  std::vector<bool> col_calculated(cols.paths.size());
  bool any = false;
  for (std::size_t r = 0; r < rows.paths.size(); ++r) {
    row_calculated[r] = path_on_calculated_item(*cache_, rows.paths[r]);
    any = any || row_calculated[r];
  }
  for (std::size_t c = 0; c < cols.paths.size(); ++c) {
    col_calculated[c] = path_on_calculated_item(*cache_, cols.paths[c]);
    any = any || col_calculated[c];
  }
  if (!any) {
    return false;
  }
  const ItemContext context{rows, cols, index_paths(rows), index_paths(cols), result.values, result};
  for (std::size_t r = 0; r < result.values.size() && r < rows.paths.size(); ++r) {
    for (std::size_t c = 0; c < result.values[r].size() && c < cols.paths.size(); ++c) {
      if (!row_calculated[r] && !col_calculated[c]) {
        continue;
      }
      std::vector<Value>& slots = result.values[r][c];
      for (std::size_t slot = 0; slot < slots.size(); ++slot) {
        if (slot < calculated_slots_.size() && calculated_slots_[slot]) {
          continue;
        }
        slots[slot] = item_cell(context, rows.paths[r], cols.paths[c], slot, 0U);
      }
    }
  }
  return true;
}

}  // namespace formulon::pivot
