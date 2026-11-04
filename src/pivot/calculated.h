//
// Calculated fields and calculated items: the model operations that keep a
// cache and every table bound to it consistent, and the evaluation of their
// cells.
//
// A calculated field is a cache field with a formula and no record slot; a
// cell shows the formula over the Sums of the fields it names at that cell,
// totals included, whatever the data field's function. A calculated item is
// a shared item flagged calculated with a `<calculatedItem>` entry; a cell
// shows the formula over the cells of the items it names in the same context,
// the totals along its field include it, and where several apply the later
// entry wins.
//
// Names collate, and formulas compile, under the caller's
// `eval::ScopedEvalProfile`.

#ifndef FORMULON_PIVOT_CALCULATED_H_
#define FORMULON_PIVOT_CALCULATED_H_

#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <string_view>
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

/// Extends every table's `fields()` with unplaced fields until it has one per
/// cache field. A calculated field's entry cannot be dragged to an axis.
void sync_table_fields(const PivotCache& cache, const std::vector<PivotTable*>& tables);

/// Appends calculated field `name` = `formula` (a leading `=` is dropped) and
/// syncs `tables`. Returns its cache field index.
///
/// Errors: `kInvalidArgument` for an empty, `All` / `Blank`, bracketed or
/// duplicate name, and for a formula the host rejects or that refers back to
/// the field through other calculated fields.
Expected<std::size_t, Error> add_calculated_field(PivotCache& cache, const std::vector<PivotTable*>& tables,
                                                  std::string_view name, std::string_view formula,
                                                  PivotFormulaHost& host);

/// Replaces a calculated field's formula. `kInvalidArgument` when `field_index`
/// is not a calculated field or the formula is rejected as on add.
Expected<void, Error> set_calculated_field_formula(PivotCache& cache, std::size_t field_index, std::string_view formula,
                                                   PivotFormulaHost& host);

/// Removes a calculated field, its data fields in every table, and renumbers
/// the fields after it. `kPreconditionFailed` while another formula names it.
Expected<void, Error> remove_calculated_field(PivotCache& cache, const std::vector<PivotTable*>& tables,
                                              std::size_t field_index, PivotFormulaHost& host);

/// Appends calculated item `name` = `formula` to cache field `field_index`:
/// a calculated shared item, its `<calculatedItem>` entry (last in solve
/// order) and an item in every table's field. Returns the entry index.
///
/// Errors: `kInvalidArgument` for a bad or duplicate name or a rejected
/// formula; `kPreconditionFailed` when the field is grouped, calculated, on no
/// table's axis, or a table summarises with a function that excludes
/// calculated items.
Expected<std::size_t, Error> add_calculated_item(PivotCache& cache, const std::vector<PivotTable*>& tables,
                                                 std::size_t field_index, std::string_view name,
                                                 std::string_view formula, PivotFormulaHost& host);

/// Replaces entry `index`'s formula, checked as on add.
Expected<void, Error> set_calculated_item_formula(PivotCache& cache, std::size_t index, std::string_view formula,
                                                  PivotFormulaHost& host);

/// Moves entry `index` to `new_index` in solve order.
Expected<void, Error> move_calculated_item(PivotCache& cache, std::size_t index, std::size_t new_index);

/// Removes entry `index`, its shared item and its item in every table, and
/// renumbers the shared items after it. `kPreconditionFailed` while another
/// calculated item names it.
Expected<void, Error> remove_calculated_item(PivotCache& cache, const std::vector<PivotTable*>& tables,
                                             std::size_t index, PivotFormulaHost& host);

/// True for the functions a data field cannot use beside calculated items:
/// Average, StdDev, StdDevP, Var and VarP.
bool aggregation_excludes_calculated_items(Aggregation aggregation);

/// `kPreconditionFailed` when `cache` has calculated items and `aggregation`
/// excludes them.
Expected<void, Error> check_aggregation_with_calculated_items(const PivotCache& cache, Aggregation aggregation);

/// The calculated formulas one evaluation of `table` needs, compiled once.
///
/// The evaluation appends one hidden Sum data field per entry of
/// `sum_fields()` after the table's own; the calculated-field slots read them.
/// Without a host nothing compiles and every calculated cell stays blank.
class CalculatedPlan {
 public:
  CalculatedPlan(const PivotTable& table, const PivotCache& cache, PivotFormulaHost* host);

  /// Database fields a calculated data field sums, in hidden-slot order.
  const std::vector<std::uint32_t>& sum_fields() const { return sum_fields_; }

  /// Per slot of the extended data-field list, true for a calculated field.
  const std::vector<bool>& calculated_slots() const { return calculated_slots_; }

  /// Evaluates every leaf cell on a calculated item, in every slot that is not
  /// a calculated field. Returns false when no leaf is on one.
  bool fill_item_cells(PivotResult& result, const AxisLeafPaths& rows, const AxisLeafPaths& cols) const;

  /// Evaluates the calculated-field slots at every cell and total.
  void fill_field_slots(PivotResult& result) const;

  /// Evaluates the calculated-field slots of one cell or total, `slots`
  /// carrying the hidden Sum slots it reads. Text results live in `result`.
  void fill_field_slots(std::vector<Value>& slots, PivotResult& result) const;

  /// The value-sort key of calculated item `key` of field `field_index`: its
  /// formula over the keys `sibling_keys` holds for the items it names. Blank
  /// when it cannot be resolved, or for a text result.
  Value item_sort_key(std::size_t field_index, const Value& key,
                      const std::map<Value, Value, ValueLess>& sibling_keys) const;

  /// Drops the hidden Sum slots from every cell and total.
  void drop_sum_slots(PivotResult& result) const;

 private:
  struct Compiled {
    bool ok = false;
    std::uint32_t id = 0;
    std::vector<FormulaRef> refs;
  };
  struct ItemContext;
  struct Inputs;

  void compile_field(std::size_t field_index, std::vector<std::uint8_t>& state);
  Value field_value(std::size_t field_index, const std::vector<Value>& slots, std::size_t depth, Inputs& inputs,
                    PivotResult& result) const;
  Value item_cell(const ItemContext& context, const std::vector<LeafPathStep>& row,
                  const std::vector<LeafPathStep>& col, std::size_t slot, std::size_t depth) const;
  bool entry_applies(std::size_t entry, const std::vector<LeafPathStep>& row,
                     const std::vector<LeafPathStep>& col) const;

  const PivotCache* cache_;
  PivotFormulaHost* host_;
  std::size_t visible_slots_;
  std::vector<std::uint32_t> sum_fields_;
  std::vector<bool> calculated_slots_;
  /// Cache field of each of the table's own data-field slots.
  std::vector<std::uint32_t> slot_fields_;
  /// By cache field index.
  std::vector<std::optional<Compiled>> fields_;
  /// By `<calculatedItem>` index.
  std::vector<Compiled> items_;
};

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_CALCULATED_H_
