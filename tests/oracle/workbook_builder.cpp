
#include "tests/oracle/workbook_builder.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "pivot/pivot_locale.h"
#include "tests/oracle/oracle_runner.h"
#include "tests/oracle/workbook_builder_internal.h"
#include "utils/a1_ref.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon {
namespace tests {
namespace oracle {

namespace workbook_builder_detail {

Error invalid(std::string message) {
  return make_error(FormulonErrorCode::kInvalidArgument, std::move(message));
}

namespace {

/// Converts one declarative cell record into a `Value`. Accepts both the
/// normalised `{kind, value}` shape the workbook case schema emits and the
/// bare JSON shorthands (number / string / bool) for hand-written specs.
/// `text` payloads are interned into the workbook so the returned `Value`
/// holds a workbook-lifetime view. `formula` records are handled by the
/// caller (`build_workbook`), which writes and recalcs them instead of
/// producing a literal `Value` here -- a formula's result depends on the
/// rest of the workbook, not just its own record.
Expected<Value, Error> value_from_record(const JsonValue& rec, Workbook* workbook, const std::string& where) {
  if (rec.is_number()) {
    return Value::number(rec.as_number());
  }
  if (rec.is_bool()) {
    return Value::boolean(rec.as_bool());
  }
  if (rec.is_string()) {
    return Value::text(workbook->intern_text(rec.as_string()));
  }
  if (rec.is_null()) {
    return Value::blank();
  }
  if (!rec.is_object()) {
    return invalid(where + ": cell record must be an object or scalar");
  }
  const JsonValue* kind_v = rec.find("kind");
  if (kind_v == nullptr || !kind_v->is_string()) {
    return invalid(where + ": cell record missing string 'kind'");
  }
  const std::string& kind = kind_v->as_string();
  if (kind == "blank") {
    return Value::blank();
  }
  const JsonValue* val_v = rec.find("value");
  if (kind == "number") {
    if (val_v == nullptr || !val_v->is_number()) {
      return invalid(where + ": number cell missing numeric 'value'");
    }
    return Value::number(val_v->as_number());
  }
  if (kind == "bool") {
    if (val_v == nullptr || !val_v->is_bool()) {
      return invalid(where + ": bool cell missing boolean 'value'");
    }
    return Value::boolean(val_v->as_bool());
  }
  if (kind == "text") {
    if (val_v == nullptr || !val_v->is_string()) {
      return invalid(where + ": text cell missing string 'value'");
    }
    return Value::text(workbook->intern_text(val_v->as_string()));
  }
  return invalid(where + ": unsupported cell kind '" + kind + "' for a pivot source");
}
}  // namespace

/// Builds the case's workbook from the `sheets` block. Each sheet's cell
/// map is written via `Sheet::set_cell_value`.
///
/// Sheets are added in the block's declaration order, not by name: the
/// capture side builds the same fixture in the order the case states, so
/// ordering by name here would give the two halves different sheet indices
/// and the golden diff would compare the wrong sheet.
Expected<std::unique_ptr<Workbook>, Error> build_workbook(const JsonValue& spec, ExcelProfile profile) {
  auto workbook = std::make_unique<Workbook>(Workbook::create_empty());
  workbook->set_excel_profile(profile);

  const JsonValue* sheets_v = spec.find("sheets");
  if (sheets_v == nullptr || !sheets_v->is_object()) {
    return invalid("spec is missing a 'sheets' object");
  }
  bool wrote_formula = false;
  for (const std::string& sheet_name : sheets_v->object_keys()) {
    const JsonValue* cells = sheets_v->find(sheet_name);
    const std::size_t sheet_index = workbook->add_sheet(sheet_name);
    Sheet& sheet = workbook->sheet(sheet_index);
    if (cells == nullptr || !cells->is_object()) {
      return invalid("sheets/" + sheet_name + ": expected an A1 -> value object");
    }
    for (const auto& [addr, rec] : cells->as_object()) {
      std::uint32_t row = 0;
      std::uint32_t col = 0;
      if (!a1_to_row_col(addr, &row, &col)) {
        std::string detail = "sheets/";
        detail += sheet_name;
        detail += ": malformed A1 address '";
        detail += addr;
        detail += "'";
        return invalid(std::move(detail));
      }
      std::string where = "sheets/";
      where += sheet_name;
      where += "/";
      where += addr;
      // A formula cell is written and left for the recalc pass below,
      // rather than resolved to a `Value` here: its result can depend on
      // other cells (and, for a pivot source column, needs to carry
      // through as an error/text/number the same way a real Excel-authored
      // formula cell would when the pivot reads it back).
      if (rec.is_object()) {
        const JsonValue* kind_v = rec.find("kind");
        if (kind_v != nullptr && kind_v->is_string() && kind_v->as_string() == "formula") {
          const JsonValue* formula_v = rec.find("formula");
          if (formula_v == nullptr || !formula_v->is_string()) {
            return invalid(where + ": formula cell missing string 'formula'");
          }
          RETURN_IF_ERROR(workbook->set_cell_formula(sheet_index, row, col, formula_v->as_string()));
          wrote_formula = true;
          continue;
        }
      }
      ASSIGN_OR_RETURN(Value v, value_from_record(rec, workbook.get(), where));
      sheet.set_cell_value(row, col, v);
    }
  }
  if (wrote_formula) {
    RETURN_IF_ERROR(workbook->recalc(eval::default_registry()));
  }
  return workbook;
}
}  // namespace workbook_builder_detail

using ::formulon::pivot::PivotCache;
using ::formulon::pivot::PivotTable;
using workbook_builder_detail::invalid;

Expected<std::vector<FormulaProbeResult>, Error> evaluate_pivot_formula_probes(BuiltPivot* built,
                                                                               const JsonValue& spec) {
  if (built == nullptr || built->workbook == nullptr) {
    return invalid("formula probes require a live BuiltPivot workbook");
  }
  const JsonValue* pivot_v = spec.find("pivot");
  if (pivot_v == nullptr || !pivot_v->is_object()) {
    return invalid("formula probes require a pivot block");
  }
  const JsonValue* probes_v = pivot_v->find("formula_probes");
  if (probes_v == nullptr || probes_v->is_null()) {
    return std::vector<FormulaProbeResult>();
  }
  if (!probes_v->is_array() || probes_v->as_array().empty()) {
    return invalid("pivot 'formula_probes' must be a non-empty array");
  }

  struct PendingProbe {
    std::string id;
    std::string sheet;
    std::string address;
    std::string formula;
    std::uint32_t row = 0;
    std::uint32_t col = 0;
  };
  std::vector<PendingProbe> pending;
  pending.reserve(probes_v->as_array().size());
  std::vector<std::string> seen_ids;
  for (std::size_t i = 0; i < probes_v->as_array().size(); ++i) {
    const JsonValue& probe = probes_v->as_array()[i];
    const std::string where = "pivot/formula_probes/" + std::to_string(i);
    if (!probe.is_object()) {
      return invalid(where + ": expected an object");
    }
    const JsonValue* id_v = probe.find("id");
    const JsonValue* cell_v = probe.find("cell");
    const JsonValue* formula_v = probe.find("formula");
    if (id_v == nullptr || !id_v->is_string() || id_v->as_string().empty() || cell_v == nullptr ||
        !cell_v->is_string() || cell_v->as_string().empty() || formula_v == nullptr || !formula_v->is_string() ||
        formula_v->as_string().empty()) {
      return invalid(where + ": requires non-empty string id, cell, and formula");
    }
    if (std::find(seen_ids.begin(), seen_ids.end(), id_v->as_string()) != seen_ids.end()) {
      return invalid(where + ": duplicate id '" + id_v->as_string() + "'");
    }
    seen_ids.push_back(id_v->as_string());
    auto [sheet, address] = split_sheet_qualified_addr(cell_v->as_string());
    if (sheet.empty()) {
      return invalid(where + "/cell must be sheet-qualified (Sheet!A1)");
    }
    PendingProbe item;
    item.id = id_v->as_string();
    item.sheet = std::move(sheet);
    item.address = std::move(address);
    item.formula = formula_v->as_string();
    if (!a1_to_row_col(item.address, &item.row, &item.col)) {
      return invalid(where + "/cell has a malformed A1 address");
    }
    pending.push_back(std::move(item));
  }

  const JsonValue* anchor_v = pivot_v->find("anchor");
  if (anchor_v == nullptr || !anchor_v->is_string()) {
    return invalid("pivot block missing string 'anchor'");
  }
  auto [anchor_sheet_name, anchor_address] = split_sheet_qualified_addr(anchor_v->as_string());
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  if (anchor_sheet_name.empty() || !a1_to_row_col(anchor_address, &anchor_row, &anchor_col)) {
    return invalid("pivot 'anchor' has a malformed sheet-qualified A1 address");
  }
  const std::size_t anchor_sheet = built->workbook->sheet_index_by_name(anchor_sheet_name);
  if (anchor_sheet == static_cast<std::size_t>(-1)) {
    return invalid("pivot 'anchor' names an unknown sheet '" + anchor_sheet_name + "'");
  }
  for (const PendingProbe& probe : pending) {
    if (built->workbook->sheet_index_by_name(probe.sheet) == static_cast<std::size_t>(-1)) {
      return invalid("formula probe '" + probe.id + "' names an unknown sheet '" + probe.sheet + "'");
    }
  }

  // The declarative builder leaves span rows/cols at zero because the grid
  // verifier derives the extent from the rendered result. GETPIVOTDATA's
  // anchor lookup is structural, so give the attached table the remaining
  // sheet rectangle; this mirrors Excel's PivotTable range for all probes
  // without hard-coding a fixture-specific output size.
  built->table.set_anchor(anchor_row, anchor_col, Sheet::kMaxRows - anchor_row, Sheet::kMaxCols - anchor_col);
  auto cache = std::make_unique<PivotCache>(std::move(built->cache));
  auto table = std::make_unique<PivotTable>(std::move(built->table));
  built->workbook->add_pivot_cache(std::move(cache));
  built->workbook->sheet(anchor_sheet).add_pivot_table(std::move(table));

  for (const PendingProbe& probe : pending) {
    const std::size_t sheet_index = built->workbook->sheet_index_by_name(probe.sheet);
    auto stored = built->workbook->set_cell_formula(sheet_index, probe.row, probe.col, probe.formula);
    if (!stored) {
      return invalid("formula probe '" + probe.id + "' could not be stored: " + stored.error().message);
    }
  }
  auto recalculated = built->workbook->recalc(eval::default_registry());
  if (!recalculated) {
    return invalid("formula probe recalc failed: " + recalculated.error().message);
  }

  std::vector<FormulaProbeResult> out;
  out.reserve(pending.size());
  for (const PendingProbe& probe : pending) {
    const std::size_t sheet_index = built->workbook->sheet_index_by_name(probe.sheet);
    const Value value = built->workbook->sheet(sheet_index).resolve_cell_value(probe.row, probe.col);
    out.push_back(FormulaProbeResult{probe.id, value});
  }
  return out;
}

}  // namespace oracle
}  // namespace tests
}  // namespace formulon
