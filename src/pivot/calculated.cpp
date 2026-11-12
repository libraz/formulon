//
// Calculated-field and calculated-item evaluation; the structural edits live
// in `calculated_edit.cpp`.
//
// @size-budget: 28 KB (with calculated_edit.cpp)

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

#include "pivot/calculated_internal.h"
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
using namespace detail;

namespace {

// Formulas nested deeper than this, a cycle included, evaluate to #REF!.
// @limit: internal
constexpr std::size_t kMaxCalculatedDepth = 64;

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
  return intern_text(result.text_storage, value.as_text());
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
