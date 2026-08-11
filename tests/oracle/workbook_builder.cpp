
#include "tests/oracle/workbook_builder.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "eval/compat.h"
#include "eval/function_registry.h"
#include "eval/pivot_locale.h"
#include "eval/recalc_engine.h"
#include "io/a1_ref.h"
#include "io/defined_names.h"
#include "tests/oracle/oracle_runner.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon {
namespace tests {
namespace oracle {

namespace {

using ::formulon::pivot::Aggregation;
using ::formulon::pivot::PivotAxis;
using ::formulon::pivot::PivotCache;
using ::formulon::pivot::PivotCacheField;
using ::formulon::pivot::PivotCacheRecord;
using ::formulon::pivot::PivotDataField;
using ::formulon::pivot::PivotField;
using ::formulon::pivot::PivotItem;
using ::formulon::pivot::PivotLayout;
using ::formulon::pivot::PivotTable;

/// Cache id stamped on the single cache the builder constructs. The
/// workbook-oracle harness never assembles more than one pivot per case,
/// so a fixed id is sufficient and keeps the table/cache binding obvious.
constexpr std::uint32_t kBuilderCacheId = 1;

Error invalid(std::string message) {
  return make_error(FormulonErrorCode::kInvalidArgument, std::move(message));
}

/// Converts one declarative cell record into a `Value`. Accepts both the
/// normalised `{kind, value}` shape the workbook case schema emits and the
/// bare JSON shorthands (number / string / bool) for hand-written specs.
/// `text` payloads are interned into the workbook so the returned `Value`
/// holds a workbook-lifetime view. `formula` records are not evaluated:
/// the pivot harness only consumes literal source data, so a formula cell
/// is rejected rather than silently producing a blank.
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

/// Builds the case's workbook from the `sheets` block. Each sheet's cell
/// map is written via `Sheet::set_cell_value`.
Expected<std::unique_ptr<Workbook>, Error> build_workbook(const JsonValue& spec) {
  auto workbook = std::make_unique<Workbook>(Workbook::create_empty());

  const JsonValue* sheets_v = spec.find("sheets");
  if (sheets_v == nullptr || !sheets_v->is_object()) {
    return invalid("spec is missing a 'sheets' object");
  }
  for (const auto& [sheet_name, cells] : sheets_v->as_object()) {
    Sheet& sheet = workbook->add_sheet(sheet_name);
    if (!cells.is_object()) {
      return invalid("sheets/" + sheet_name + ": expected an A1 -> value object");
    }
    for (const auto& [addr, rec] : cells.as_object()) {
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
      ASSIGN_OR_RETURN(Value v, value_from_record(rec, workbook.get(), where));
      sheet.set_cell_value(row, col, v);
    }
  }
  return workbook;
}

/// Parses a sheet-qualified A1 range ("Data!A1:C13") into the sheet name
/// plus 0-based inclusive row/col bounds.
struct SourceRange {
  std::string sheet;
  std::uint32_t r0 = 0;
  std::uint32_t c0 = 0;
  std::uint32_t r1 = 0;
  std::uint32_t c1 = 0;
};

Expected<SourceRange, Error> parse_source_range(const std::string& spec_text) {
  auto [sheet, bare] = split_sheet_qualified_addr(spec_text);
  if (sheet.empty()) {
    return invalid("pivot 'source' must be sheet-qualified: '" + spec_text + "'");
  }
  const std::size_t colon = bare.find(':');
  if (colon == std::string::npos) {
    return invalid("pivot 'source' must be an A1 range: '" + spec_text + "'");
  }
  SourceRange out;
  out.sheet = sheet;
  if (!a1_to_row_col(bare.substr(0, colon), &out.r0, &out.c0) ||
      !a1_to_row_col(bare.substr(colon + 1), &out.r1, &out.c1)) {
    return invalid("pivot 'source' has a malformed endpoint: '" + spec_text + "'");
  }
  if (out.r1 < out.r0 || out.c1 < out.c0) {
    return invalid("pivot 'source' range is inverted: '" + spec_text + "'");
  }
  return out;
}

/// Maps an `agg` string from the declarative spec onto the `Aggregation`
/// enum. The accepted names mirror the enumerators one-to-one.
Expected<Aggregation, Error> aggregation_from_string(const std::string& name) {
  if (name == "Sum") {
    return Aggregation::Sum;
  }
  if (name == "Count") {
    return Aggregation::Count;
  }
  if (name == "Average") {
    return Aggregation::Average;
  }
  if (name == "Max") {
    return Aggregation::Max;
  }
  if (name == "Min") {
    return Aggregation::Min;
  }
  if (name == "Product") {
    return Aggregation::Product;
  }
  if (name == "CountNumbers") {
    return Aggregation::CountNumbers;
  }
  if (name == "StdDev") {
    return Aggregation::StdDev;
  }
  if (name == "StdDevP") {
    return Aggregation::StdDevP;
  }
  if (name == "Var") {
    return Aggregation::Var;
  }
  if (name == "VarP") {
    return Aggregation::VarP;
  }
  return invalid("unknown aggregation '" + name + "'");
}

/// Maps a `layout` string onto the `PivotLayout` enum. Defaults to Compact
/// when the spec omits the field; an unrecognised value is an error.
Expected<PivotLayout, Error> layout_from_string(const std::string& name) {
  if (name == "Compact") {
    return PivotLayout::Compact;
  }
  if (name == "Tabular") {
    return PivotLayout::Tabular;
  }
  if (name == "Outline") {
    return PivotLayout::Outline;
  }
  return invalid("unknown pivot layout '" + name + "'");
}

/// Returns the 0-based column offset of `header` within the source-range
/// header row, or `npos`-equivalent (`field_count`) when not found.
std::uint32_t header_index(const std::vector<std::string>& headers, const std::string& name) {
  for (std::uint32_t i = 0; i < headers.size(); ++i) {
    if (headers[i] == name) {
      return i;
    }
  }
  return static_cast<std::uint32_t>(headers.size());
}

/// Collects the string list at `obj[key]`, defaulting to empty when the
/// key is absent. Each element must be a string.
Expected<std::vector<std::string>, Error> string_list(const JsonValue& obj, const char* key) {
  std::vector<std::string> out;
  const JsonValue* v = obj.find(key);
  if (v == nullptr || v->is_null()) {
    return out;
  }
  if (!v->is_array()) {
    return invalid(std::string("pivot '") + key + "' must be an array");
  }
  for (const JsonValue& item : v->as_array()) {
    if (!item.is_string()) {
      return invalid(std::string("pivot '") + key + "' entries must be strings");
    }
    out.push_back(item.as_string());
  }
  return out;
}

}  // namespace

Expected<BuiltPivot, Error> build_pivot_from_spec(const JsonValue& spec) {
  if (!spec.is_object()) {
    return invalid("workbook spec must be an object");
  }
  const JsonValue* pivot_v = spec.find("pivot");
  if (pivot_v == nullptr || !pivot_v->is_object()) {
    return invalid("workbook spec has no 'pivot' block");
  }
  const JsonValue& pivot = *pivot_v;

  ASSIGN_OR_RETURN(std::unique_ptr<Workbook> workbook, build_workbook(spec));

  // --- source range --------------------------------------------------------
  const JsonValue* source_v = pivot.find("source");
  if (source_v == nullptr || !source_v->is_string()) {
    return invalid("pivot block missing string 'source'");
  }
  ASSIGN_OR_RETURN(SourceRange src, parse_source_range(source_v->as_string()));
  const Sheet* src_sheet = workbook->sheet_by_name(src.sheet);
  if (src_sheet == nullptr) {
    return invalid("pivot 'source' names an unknown sheet '" + src.sheet + "'");
  }

  // --- anchor --------------------------------------------------------------
  const JsonValue* anchor_v = pivot.find("anchor");
  if (anchor_v == nullptr || !anchor_v->is_string()) {
    return invalid("pivot block missing string 'anchor'");
  }
  auto [anchor_sheet, anchor_bare] = split_sheet_qualified_addr(anchor_v->as_string());
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  if (!a1_to_row_col(anchor_bare, &anchor_row, &anchor_col)) {
    return invalid("pivot 'anchor' has a malformed A1 address");
  }

  // --- cache: header row + records ----------------------------------------
  std::vector<std::string> headers;
  headers.reserve(src.c1 - src.c0 + 1);
  for (std::uint32_t c = src.c0; c <= src.c1; ++c) {
    const Value header = src_sheet->resolve_cell_value(src.r0, c);
    if (header.kind() != ValueKind::Text) {
      return invalid("pivot 'source' header row must be text in every column");
    }
    headers.emplace_back(header.as_text());
  }

  PivotCache cache;
  cache.set_cache_id(kBuilderCacheId);
  for (const std::string& name : headers) {
    cache.mutable_fields().push_back(PivotCacheField{name, {}});
  }
  for (std::uint32_t r = src.r0 + 1; r <= src.r1; ++r) {
    PivotCacheRecord rec;
    rec.cells.reserve(headers.size());
    for (std::uint32_t c = src.c0; c <= src.c1; ++c) {
      Value cell = src_sheet->resolve_cell_value(r, c);
      // The cache must outlive the workbook's text views once the
      // workbook is moved into `BuiltPivot`, so any text payload is
      // re-interned into the cache's own pointer-stable storage.
      if (cell.kind() == ValueKind::Text) {
        cache.mutable_text_storage().emplace_back(cell.as_text());
        cell = Value::text(cache.text_storage().back());
      }
      rec.cells.push_back(cell);
    }
    cache.mutable_records().push_back(std::move(rec));
  }

  // Distinct shared items per column, in first-seen order. Numeric /
  // boolean columns also get shared items here; the evaluator keys on the
  // record cells directly, so this is purely informational metadata.
  for (std::uint32_t f = 0; f < headers.size(); ++f) {
    std::vector<std::string> seen_text;
    for (const PivotCacheRecord& rec : cache.records()) {
      const Value& cell = rec.cells[f];
      if (cell.kind() != ValueKind::Text) {
        continue;
      }
      const std::string text(cell.as_text());
      if (std::find(seen_text.begin(), seen_text.end(), text) == seen_text.end()) {
        seen_text.push_back(text);
        cache.mutable_text_storage().emplace_back(text);
        cache.mutable_fields()[f].shared_items.push_back(Value::text(cache.text_storage().back()));
      }
    }
  }

  // --- table: fields + axes ------------------------------------------------
  ASSIGN_OR_RETURN(std::vector<std::string> row_fields, string_list(pivot, "row_fields"));
  ASSIGN_OR_RETURN(std::vector<std::string> col_fields, string_list(pivot, "col_fields"));
  ASSIGN_OR_RETURN(std::vector<std::string> page_fields, string_list(pivot, "page_fields"));

  PivotTable table;
  table.set_pivot_cache_id(kBuilderCacheId);

  // Manual item filters: collect a hidden-item set per field name so the
  // matching `PivotField::items` can carry `visible = false`.
  std::map<std::string, std::vector<std::string>> hidden_items;
  if (const JsonValue* filters_v = pivot.find("filters"); filters_v != nullptr && !filters_v->is_null()) {
    if (!filters_v->is_array()) {
      return invalid("pivot 'filters' must be an array");
    }
    for (const JsonValue& filter : filters_v->as_array()) {
      if (!filter.is_object()) {
        return invalid("pivot 'filters' entries must be objects");
      }
      const JsonValue* field_v = filter.find("field");
      const JsonValue* hide_v = filter.find("hide");
      if (field_v == nullptr || !field_v->is_string() || hide_v == nullptr || !hide_v->is_array()) {
        return invalid("pivot filter needs string 'field' and array 'hide'");
      }
      std::vector<std::string>& hidden = hidden_items[field_v->as_string()];
      for (const JsonValue& item : hide_v->as_array()) {
        if (!item.is_string()) {
          return invalid("pivot filter 'hide' entries must be strings");
        }
        hidden.push_back(item.as_string());
      }
    }
  }

  // Build one `PivotField` per source header, in source-column order, so a
  // data field's `field_index` lines up with the cache field index. The
  // axis defaults to Page (treated as "not on row/col") and is promoted to
  // Row / Col when the field name appears in `row_fields` / `col_fields`.
  for (std::uint32_t f = 0; f < headers.size(); ++f) {
    PivotField field;
    field.source_name = headers[f];
    field.axis = PivotAxis::Page;
    auto hidden_it = hidden_items.find(headers[f]);
    if (hidden_it != hidden_items.end()) {
      for (const Value& shared : cache.fields()[f].shared_items) {
        if (shared.kind() != ValueKind::Text) {
          continue;
        }
        const std::string item_name(shared.as_text());
        const bool hidden =
            std::find(hidden_it->second.begin(), hidden_it->second.end(), item_name) != hidden_it->second.end();
        field.items.push_back(PivotItem{item_name, !hidden});
      }
    }
    table.mutable_fields().push_back(std::move(field));
  }

  for (const std::string& name : row_fields) {
    const std::uint32_t idx = header_index(headers, name);
    if (idx >= headers.size()) {
      return invalid("pivot 'row_fields' names a non-source field '" + name + "'");
    }
    table.mutable_fields()[idx].axis = PivotAxis::Row;
    table.mutable_row_field_order().push_back(idx);
  }
  for (const std::string& name : col_fields) {
    const std::uint32_t idx = header_index(headers, name);
    if (idx >= headers.size()) {
      return invalid("pivot 'col_fields' names a non-source field '" + name + "'");
    }
    table.mutable_fields()[idx].axis = PivotAxis::Col;
    table.mutable_col_field_order().push_back(idx);
  }
  for (const std::string& name : page_fields) {
    const std::uint32_t idx = header_index(headers, name);
    if (idx >= headers.size()) {
      return invalid("pivot 'page_fields' names a non-source field '" + name + "'");
    }
    if (table.mutable_fields()[idx].axis == PivotAxis::Row || table.mutable_fields()[idx].axis == PivotAxis::Col) {
      return invalid("pivot 'page_fields' overlaps a row/column field '" + name + "'");
    }
    table.mutable_fields()[idx].axis = PivotAxis::Page;
  }

  // Excel's compact layout (the default) implicitly shows a subtotal
  // row for every non-leaf row field. The declarative spec does not
  // model that directly, so flip `subtotal_top` on every row / col
  // field that has at least one descendant on the same axis. The
  // leaf-most field on each axis is left alone — its values already
  // occupy the data rows / columns.
  if (table.mutable_row_field_order().size() > 1) {
    for (std::size_t i = 0; i + 1 < table.mutable_row_field_order().size(); ++i) {
      const std::uint32_t idx = table.mutable_row_field_order()[i];
      table.mutable_fields()[idx].subtotal_top = true;
    }
  }
  if (table.mutable_col_field_order().size() > 1) {
    for (std::size_t i = 0; i + 1 < table.mutable_col_field_order().size(); ++i) {
      const std::uint32_t idx = table.mutable_col_field_order()[i];
      table.mutable_fields()[idx].subtotal_top = true;
    }
  }

  // --- table: data fields --------------------------------------------------
  // Excel auto-disambiguates the displayed source-field name when the
  // same source column is aggregated more than once on the value axis
  // (e.g. Sum(Amount) + Count(Amount) renders as "合計 / Amount" and
  // "個数 / Amount2"). Track the running per-source-name occurrence
  // count so the synthesised display names match the Excel pivot UI.
  const JsonValue* data_v = pivot.find("data_fields");
  if (data_v == nullptr || !data_v->is_array() || data_v->as_array().empty()) {
    return invalid("pivot block needs a non-empty 'data_fields' array");
  }
  const eval::ExcelProfile profile = workbook->excel_profile();
  std::map<std::string, std::uint32_t> source_name_occurrences;
  for (const JsonValue& df : data_v->as_array()) {
    if (!df.is_object()) {
      return invalid("pivot 'data_fields' entries must be objects");
    }
    const JsonValue* field_v = df.find("field");
    const JsonValue* agg_v = df.find("agg");
    if (field_v == nullptr || !field_v->is_string()) {
      return invalid("pivot data field missing string 'field'");
    }
    if (agg_v == nullptr || !agg_v->is_string()) {
      return invalid("pivot data field missing string 'agg'");
    }
    const std::uint32_t idx = header_index(headers, field_v->as_string());
    if (idx >= headers.size()) {
      return invalid("pivot data field names a non-source field '" + field_v->as_string() + "'");
    }
    ASSIGN_OR_RETURN(Aggregation agg, aggregation_from_string(agg_v->as_string()));
    table.mutable_fields()[idx].axis = PivotAxis::Value;

    const std::string& source = field_v->as_string();
    const std::uint32_t occurrence = ++source_name_occurrences[source];
    std::string display_source = source;
    if (occurrence > 1) {
      display_source += std::to_string(occurrence);
    }

    PivotDataField data_field;
    data_field.name = eval::data_field_display_name(agg, display_source, profile);
    data_field.field_index = idx;
    data_field.aggregation = agg;
    table.mutable_data_fields().push_back(std::move(data_field));
  }

  // --- table: layout / anchor / grand totals -------------------------------
  if (const JsonValue* layout_v = pivot.find("layout"); layout_v != nullptr && !layout_v->is_null()) {
    if (!layout_v->is_string()) {
      return invalid("pivot 'layout' must be a string");
    }
    ASSIGN_OR_RETURN(PivotLayout layout, layout_from_string(layout_v->as_string()));
    table.set_layout(layout);
  }

  bool grand_rows = true;
  bool grand_cols = true;
  if (const JsonValue* gt_v = pivot.find("grand_totals"); gt_v != nullptr && !gt_v->is_null()) {
    if (!gt_v->is_object()) {
      return invalid("pivot 'grand_totals' must be an object");
    }
    if (const JsonValue* rv = gt_v->find("rows"); rv != nullptr && rv->is_bool()) {
      grand_rows = rv->as_bool();
    }
    if (const JsonValue* cv = gt_v->find("cols"); cv != nullptr && cv->is_bool()) {
      grand_cols = cv->as_bool();
    }
  }
  table.set_grand_totals(grand_rows, grand_cols);
  // Span is left at zero: the layout pass derives the rendered extent from
  // the evaluated result, not from the anchor's span fields.
  table.set_anchor(anchor_row, anchor_col, /*rows=*/0, /*cols=*/0);

  BuiltPivot out;
  out.workbook = std::move(workbook);
  out.cache = std::move(cache);
  out.table = std::move(table);
  return out;
}

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

// ---------------------------------------------------------------------------
// Print-spec builder
// ---------------------------------------------------------------------------

namespace {

/// OOXML built-in defined-name identifiers for the print area and titles.
constexpr const char* kPrintAreaName = "_xlnm.Print_Area";
constexpr const char* kPrintTitlesName = "_xlnm.Print_Titles";

/// Applies the case-level `column_widths` / `row_heights` maps onto the
/// first sheet's layout overrides.
///
/// A `column_widths` key is a column letter or `first:last` span
/// ("A" / "A:D"); each entry becomes one `ColumnLayout{first,last,width}`.
/// A `row_heights` key is a 1-based Excel row number ("3"); each entry
/// becomes one `RowLayout{row,height}`. Both maps are optional.
Expected<void, Error> apply_layout_dimensions(const JsonValue& spec, Sheet* sheet) {
  if (const JsonValue* widths_v = spec.find("column_widths"); widths_v != nullptr && !widths_v->is_null()) {
    if (!widths_v->is_object()) {
      return invalid("'column_widths' must be an object");
    }
    for (const auto& [key, value] : widths_v->as_object()) {
      if (!value.is_number()) {
        return invalid("column_widths/" + key + ": width must be a number");
      }
      const std::size_t colon = key.find(':');
      std::string_view lhs = key;
      std::string_view rhs = key;
      if (colon != std::string::npos) {
        lhs = std::string_view(key).substr(0, colon);
        rhs = std::string_view(key).substr(colon + 1);
      }
      std::size_t p_lhs = 0;
      std::size_t p_rhs = 0;
      std::uint32_t first = 0;
      std::uint32_t last = 0;
      if (!io::parse_column_letters(lhs, &p_lhs, &first) || p_lhs != lhs.size() ||
          !io::parse_column_letters(rhs, &p_rhs, &last) || p_rhs != rhs.size()) {
        return invalid("column_widths/" + key + ": malformed column key");
      }
      // `parse_column_letters` yields a 1-based column ordinal.
      ColumnLayout col;
      col.first = std::min(first, last) - 1U;
      col.last = std::max(first, last) - 1U;
      col.width = value.as_number();
      sheet->mutable_layout().columns.push_back(col);
    }
  }

  if (const JsonValue* heights_v = spec.find("row_heights"); heights_v != nullptr && !heights_v->is_null()) {
    if (!heights_v->is_object()) {
      return invalid("'row_heights' must be an object");
    }
    for (const auto& [key, value] : heights_v->as_object()) {
      if (!value.is_number()) {
        return invalid("row_heights/" + key + ": height must be a number");
      }
      std::size_t pos = 0;
      std::uint32_t row1 = 0;
      if (!io::parse_uint(key, &pos, &row1) || pos != key.size() || row1 == 0U) {
        return invalid("row_heights/" + key + ": malformed 1-based row key");
      }
      RowLayout row;
      row.row = row1 - 1U;
      row.height = value.as_number();
      sheet->mutable_layout().row_overrides.push_back(row);
    }
  }
  return {};
}

/// Maps an `orientation` string onto the `Orientation` enum.
Expected<Orientation, Error> orientation_from_string(const std::string& name) {
  if (name == "portrait") {
    return Orientation::kPortrait;
  }
  if (name == "landscape") {
    return Orientation::kLandscape;
  }
  if (name == "default") {
    return Orientation::kDefault;
  }
  return invalid("unknown page orientation '" + name + "'");
}

/// Translates the declarative `page_setup` block into a `PageSetup`.
///
/// Absent fields keep the struct/OOXML defaults. A non-zero
/// `fit_to_width` / `fit_to_height` flips `fit_to_page` on so the
/// pagination engine derives a shrink factor instead of using `scale`.
Expected<PageSetup, Error> page_setup_from_spec(const JsonValue& block) {
  PageSetup setup;
  if (!block.is_object()) {
    return invalid("'page_setup' must be an object");
  }
  if (const JsonValue* v = block.find("orientation"); v != nullptr && !v->is_null()) {
    if (!v->is_string()) {
      return invalid("page_setup 'orientation' must be a string");
    }
    ASSIGN_OR_RETURN(setup.orientation, orientation_from_string(v->as_string()));
  }
  if (const JsonValue* v = block.find("paper"); v != nullptr && !v->is_null()) {
    if (!v->is_number()) {
      return invalid("page_setup 'paper' must be a number");
    }
    setup.paper_size = static_cast<std::uint32_t>(v->as_number());
  }
  if (const JsonValue* v = block.find("scale"); v != nullptr && !v->is_null()) {
    if (!v->is_number()) {
      return invalid("page_setup 'scale' must be a number");
    }
    setup.scale = static_cast<std::uint32_t>(v->as_number());
  }
  // `PageSetup`'s struct default leaves `fit_to_width` / `fit_to_height`
  // at 1, which would imply fit-to-page even when the case never asked
  // for it. The declarative spec instead treats an absent fit field as
  // 0 (axis unconstrained); only an explicit non-zero entry turns the
  // fit-to-page toggle on. This mirrors Excel's mutually-exclusive
  // scale-vs-fit radio: a `scale` case keeps both fit counts at 0.
  bool fit_specified = false;
  std::uint32_t fit_width = 0;
  std::uint32_t fit_height = 0;
  if (const JsonValue* v = block.find("fit_to_width"); v != nullptr && !v->is_null()) {
    if (!v->is_number()) {
      return invalid("page_setup 'fit_to_width' must be a number");
    }
    fit_width = static_cast<std::uint32_t>(v->as_number());
    fit_specified = true;
  }
  if (const JsonValue* v = block.find("fit_to_height"); v != nullptr && !v->is_null()) {
    if (!v->is_number()) {
      return invalid("page_setup 'fit_to_height' must be a number");
    }
    fit_height = static_cast<std::uint32_t>(v->as_number());
    fit_specified = true;
  }
  setup.fit_to_width = fit_width;
  setup.fit_to_height = fit_height;
  setup.fit_to_page = fit_specified && (fit_width != 0U || fit_height != 0U);
  return setup;
}

/// Collects the integer list at `block[key]`, defaulting to empty. Each
/// element must be a number; `out` receives every value cast to
/// `std::uint32_t`.
Expected<std::vector<std::uint32_t>, Error> uint_list(const JsonValue& block, const char* key) {
  std::vector<std::uint32_t> out;
  const JsonValue* v = block.find(key);
  if (v == nullptr || v->is_null()) {
    return out;
  }
  if (!v->is_array()) {
    return invalid(std::string("manual_breaks '") + key + "' must be an array");
  }
  for (const JsonValue& item : v->as_array()) {
    if (!item.is_number()) {
      return invalid(std::string("manual_breaks '") + key + "' entries must be numbers");
    }
    out.push_back(static_cast<std::uint32_t>(item.as_number()));
  }
  return out;
}

}  // namespace

Expected<BuiltPrint, Error> build_print_from_spec(const JsonValue& spec) {
  if (!spec.is_object()) {
    return invalid("workbook spec must be an object");
  }
  const JsonValue* print_v = spec.find("print");
  if (print_v == nullptr || !print_v->is_object()) {
    return invalid("workbook spec has no 'print' block");
  }
  const JsonValue& print = *print_v;

  ASSIGN_OR_RETURN(std::unique_ptr<Workbook> workbook, build_workbook(spec));

  // --- target sheet --------------------------------------------------------
  const JsonValue* sheet_v = print.find("sheet");
  if (sheet_v == nullptr || !sheet_v->is_string()) {
    return invalid("print block missing string 'sheet'");
  }
  const std::string& sheet_name = sheet_v->as_string();
  std::uint32_t sheet_index = 0;
  bool found = false;
  for (std::size_t i = 0; i < workbook->sheet_count(); ++i) {
    if (workbook->sheet(i).name() == sheet_name) {
      sheet_index = static_cast<std::uint32_t>(i);
      found = true;
      break;
    }
  }
  if (!found) {
    return invalid("print 'sheet' names an unknown sheet '" + sheet_name + "'");
  }
  Sheet& sheet = workbook->sheet(sheet_index);

  // --- case-level layout dimensions ---------------------------------------
  RETURN_IF_ERROR(apply_layout_dimensions(spec, &sheet));

  // --- page setup ----------------------------------------------------------
  if (const JsonValue* setup_v = print.find("page_setup"); setup_v != nullptr && !setup_v->is_null()) {
    ASSIGN_OR_RETURN(PageSetup setup, page_setup_from_spec(*setup_v));
    sheet.mutable_print_settings().page_setup = setup;
  }

  // --- manual breaks -------------------------------------------------------
  if (const JsonValue* breaks_v = print.find("manual_breaks"); breaks_v != nullptr && !breaks_v->is_null()) {
    if (!breaks_v->is_object()) {
      return invalid("'manual_breaks' must be an object");
    }
    // `rows` carries 1-based Excel row numbers; convert to the 0-based
    // index the break sits before. `cols` carries column letters.
    ASSIGN_OR_RETURN(std::vector<std::uint32_t> break_rows, uint_list(*breaks_v, "rows"));
    for (std::uint32_t row1 : break_rows) {
      if (row1 == 0U) {
        return invalid("manual_breaks 'rows' entries must be 1-based row numbers");
      }
      ManualBreak brk;
      brk.id = row1 - 1U;
      brk.manual = true;
      sheet.mutable_print_settings().manual_row_breaks.push_back(brk);
    }
    if (const JsonValue* cols_v = breaks_v->find("cols"); cols_v != nullptr && !cols_v->is_null()) {
      if (!cols_v->is_array()) {
        return invalid("manual_breaks 'cols' must be an array");
      }
      for (const JsonValue& item : cols_v->as_array()) {
        if (!item.is_string()) {
          return invalid("manual_breaks 'cols' entries must be column letters");
        }
        const std::string& letters = item.as_string();
        std::size_t pos = 0;
        std::uint32_t col1 = 0;
        if (!io::parse_column_letters(letters, &pos, &col1) || pos != letters.size() || col1 == 0U) {
          return invalid("manual_breaks 'cols' has a malformed column letter '" + letters + "'");
        }
        ManualBreak brk;
        brk.id = col1 - 1U;
        brk.manual = true;
        sheet.mutable_print_settings().manual_col_breaks.push_back(brk);
      }
    }
  }

  // --- print-area / print-titles defined names -----------------------------
  // Excel stores the print area / titles as sheet-scoped built-in defined
  // names whose formula is a fully-qualified A1 range. The print-area
  // resolver strips the sheet qualifier and `$` anchors, so a plain
  // `Sheet1!A1:H80` form is sufficient here.
  std::vector<io::DefinedName> defined_names = workbook->defined_names();
  if (const JsonValue* area_v = print.find("print_area"); area_v != nullptr && !area_v->is_null()) {
    if (!area_v->is_string()) {
      return invalid("print 'print_area' must be a string");
    }
    io::DefinedName dn;
    dn.name = kPrintAreaName;
    dn.formula = sheet_name + "!" + area_v->as_string();
    dn.local_sheet_id = static_cast<std::int32_t>(sheet_index);
    defined_names.push_back(std::move(dn));
  }
  if (const JsonValue* titles_v = print.find("print_titles"); titles_v != nullptr && !titles_v->is_null()) {
    if (!titles_v->is_object()) {
      return invalid("print 'print_titles' must be an object");
    }
    std::string formula;
    if (const JsonValue* rows_v = titles_v->find("rows"); rows_v != nullptr && !rows_v->is_null()) {
      if (!rows_v->is_string()) {
        return invalid("print_titles 'rows' must be a string");
      }
      formula = sheet_name + "!" + rows_v->as_string();
    }
    if (const JsonValue* cols_v = titles_v->find("cols"); cols_v != nullptr && !cols_v->is_null()) {
      if (!cols_v->is_string()) {
        return invalid("print_titles 'cols' must be a string");
      }
      if (!formula.empty()) {
        formula += ",";
      }
      formula += sheet_name + "!" + cols_v->as_string();
    }
    if (!formula.empty()) {
      io::DefinedName dn;
      dn.name = kPrintTitlesName;
      dn.formula = std::move(formula);
      dn.local_sheet_id = static_cast<std::int32_t>(sheet_index);
      defined_names.push_back(std::move(dn));
    }
  }
  workbook->set_defined_names(std::move(defined_names));

  BuiltPrint out;
  out.workbook = std::move(workbook);
  out.sheet_index = sheet_index;
  return out;
}

}  // namespace oracle
}  // namespace tests
}  // namespace formulon
