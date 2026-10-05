
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
#include "tests/oracle/workbook_builder.h"
#include "tests/oracle/workbook_builder_internal.h"
#include "utils/a1_ref.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon {
namespace tests {
namespace oracle {

using workbook_builder_detail::build_workbook;
using workbook_builder_detail::invalid;

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
    // Every cell below is stored inline (verbatim), never as an index into
    // `shared_items` -- explicit so `cell_value()` (record_access.h) does
    // not fall back to inferring the encoding from `shared_items` being
    // non-empty. That fallback assumes a field with shared items indexes
    // *every* cell through them; this harness only populates shared_items
    // from a field's text values (informational, for filter-item lookup),
    // so a field mixing numbers and text -- e.g. Count(Amount) with one
    // error/text cell among the numbers -- would otherwise have its
    // numeric cells misread as out-of-range shared-item indices and
    // collapse to Blank.
    rec.cell_is_index.assign(headers.size(), false);
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
  // axis starts at `None` — an available field the report does not place —
  // and is promoted when the field name appears in `row_fields` /
  // `col_fields` / `page_fields`. `Page` is a placement of its own and
  // draws a header above the report, so a source column nobody positioned
  // must not borrow it.
  for (std::uint32_t f = 0; f < headers.size(); ++f) {
    PivotField field;
    field.source_name = headers[f];
    field.axis = PivotAxis::None;
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
  const ExcelProfile profile = workbook->excel_profile();
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
    data_field.name = pivot::data_field_display_name(agg, display_source, profile);
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

}  // namespace oracle
}  // namespace tests
}  // namespace formulon
