//
// C ABI - workbook tables: read-side iteration and create / update / remove.

#include <algorithm>
#include <cctype>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>

#include "auto_filter.h"
#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "drawing/slicer_edit.h"
#include "io/auto_filter_xml.h"
#include "io/xml_escape.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_stored_formula;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::set_binding_error;

namespace {

std::string ascii_lower(std::string value) {
  for (char& ch : value) {
    ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
  }
  return value;
}

std::string xml_attr_escape(std::string_view value) {
  std::string out;
  out.reserve(value.size());
  formulon::io::AppendXmlAttrEscaped(out, value);
  return out;
}

/// Column span of an A1 range such as `"A1:C10"`, or 0 when the text is not
/// the plain single-area form that `xl/tables/tableN.xml` accepts for `ref`.
/// A table whose column count disagrees with its range is a file Excel
/// refuses to open without repair, so the span is checked up front.
std::size_t table_ref_column_count(std::string_view ref) {
  const auto rect = formulon::parse_a1_rectangle(ref);
  return rect ? static_cast<std::size_t>(rect->last_col - rect->first_col) + 1U : 0U;
}

/// Style of a newly styled table: the named style with row stripes on, the
/// toggles the desktop application starts a table with.
formulon::TableStyleInfo default_style_info(std::string_view style_name) {
  formulon::TableStyleInfo info;
  if (!style_name.empty()) {
    info.present = true;
    info.name = std::string(style_name);
    info.show_row_stripes = true;
  }
  return info;
}

/// `SUBTOTAL` function number of each built-in `ST_TotalsRowFunction`
/// token (the 1xx forms skip hidden rows, as the desktop application uses).
struct TotalsFunction {
  std::string_view token;
  int subtotal_code;
};
constexpr TotalsFunction kTotalsFunctions[] = {
    {"none", 0},    {"sum", 109},       {"min", 105},    {"max", 104}, {"average", 101},
    {"count", 103}, {"countNums", 102}, {"stdDev", 107}, {"var", 110}, {"custom", 0},
};

const TotalsFunction* find_totals_function(std::string_view token) {
  for (const TotalsFunction& entry : kTotalsFunctions) {
    if (entry.token == token) {
      return &entry;
    }
  }
  return nullptr;
}

/// Column name as it appears inside a structured reference's brackets:
/// `'` escapes `[`, `]`, `#` and `'`, and a name the payload grammar would
/// split or trim (a comma, a colon, edge whitespace) gets its own brackets.
std::string structured_ref_column(std::string_view name) {
  std::string escaped;
  escaped.reserve(name.size() + 2U);
  bool needs_brackets = !name.empty() && (name.front() == ' ' || name.back() == ' ');
  for (const char ch : name) {
    if (ch == '[' || ch == ']' || ch == '#' || ch == '\'') {
      escaped.push_back('\'');
    } else if (ch == ',' || ch == ':') {
      needs_brackets = true;
    }
    escaped.push_back(ch);
  }
  return needs_brackets ? "[" + escaped + "]" : escaped;
}

/// Writes column `column`'s totals-row cell from its totals settings: the
/// `SUBTOTAL` formula, the custom formula, the label, or nothing.
fm_status_t write_totals_cell(formulon::Workbook& book, const formulon::TableMetadata& table, std::size_t column) {
  const auto rect = formulon::parse_a1_rectangle(table.ref);
  if (!rect || column >= table.columns.size()) {
    return 0;
  }
  const formulon::TableColumn& col = table.columns[column];
  const std::uint32_t row = rect->last_row;
  const std::uint32_t cell_col = rect->first_col + static_cast<std::uint32_t>(column);
  const TotalsFunction* function = find_totals_function(col.totals_function.empty() ? "none" : col.totals_function);
  formulon::Expected<void, formulon::Error> written = formulon::Expected<void, formulon::Error>::Ok();
  if (function != nullptr && function->subtotal_code != 0) {
    written = book.set_cell_formula(table.sheet_index, row, cell_col,
                                    "=SUBTOTAL(" + std::to_string(function->subtotal_code) + "," + table.display_name +
                                        "[" + structured_ref_column(col.name) + "])");
  } else if (col.totals_function == "custom" && !col.totals_row_formula.empty()) {
    written = book.set_cell_formula(table.sheet_index, row, cell_col, "=" + col.totals_row_formula);
  } else if (!col.totals_label.empty()) {
    written = book.set_cell_text(table.sheet_index, row, cell_col, col.totals_label);
  } else {
    written = book.set_cell_value(table.sheet_index, row, cell_col, formulon::Value::blank());
  }
  return finish_status(written);
}

/// Rejects an empty name or one another table (other than `self`) already
/// uses as its name or display name. Table names share one case-insensitive
/// namespace.
fm_status_t check_table_names(const formulon::Workbook& book, std::string_view name, std::string_view display_name,
                              std::size_t self, const char* api) {
  if (name.empty() || display_name.empty()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, api, "reason=empty table name");
  }
  const std::string name_key = ascii_lower(std::string(name));
  const std::string display_name_key = ascii_lower(std::string(display_name));
  const auto& tables = book.tables();
  for (std::size_t i = 0; i < tables.size(); ++i) {
    if (i == self) {
      continue;
    }
    const std::string existing_name_key = ascii_lower(tables[i].name);
    const std::string existing_display_name_key = ascii_lower(tables[i].display_name);
    if (existing_name_key == name_key || existing_display_name_key == name_key ||
        existing_name_key == display_name_key || existing_display_name_key == display_name_key) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, api,
                               "reason=table name already exists name=" + std::string(name));
    }
  }
  return 0;
}

// Changes only the opening element's ref attribute of an AutoFilter the
// model could not parse, which is retained verbatim.
std::string table_auto_filter_xml(std::string_view raw_xml, std::string_view ref) {
  const std::string escaped_ref = xml_attr_escape(ref);
  if (raw_xml.empty()) {
    return "<autoFilter ref=\"" + escaped_ref + "\"/>";
  }

  const std::size_t name_start = raw_xml.find("<autoFilter");
  if (name_start == std::string_view::npos) {
    return std::string(raw_xml);
  }
  const std::size_t name_end = name_start + std::string_view("<autoFilter").size();
  if (name_end < raw_xml.size() && raw_xml[name_end] != ' ' && raw_xml[name_end] != '\t' && raw_xml[name_end] != '\r' &&
      raw_xml[name_end] != '\n' && raw_xml[name_end] != '/' && raw_xml[name_end] != '>') {
    return std::string(raw_xml);
  }

  bool in_quote = false;
  char quote = '\0';
  std::size_t opening_end = std::string_view::npos;
  for (std::size_t i = name_end; i < raw_xml.size(); ++i) {
    const char ch = raw_xml[i];
    if (in_quote) {
      if (ch == quote) {
        in_quote = false;
      }
    } else if (ch == '\'' || ch == '"') {
      in_quote = true;
      quote = ch;
    } else if (ch == '>') {
      opening_end = i;
      break;
    }
  }
  if (opening_end == std::string_view::npos || in_quote) {
    return std::string(raw_xml);
  }

  // Locate a ref attribute in the opening element. This intentionally edits
  // only the attribute value, retaining its original quoting and whitespace.
  char attribute_quote = '\0';
  for (std::size_t i = name_end; i + 3U <= opening_end; ++i) {
    const char ch = raw_xml[i];
    if (attribute_quote != '\0') {
      if (ch == attribute_quote) {
        attribute_quote = '\0';
      }
      continue;
    }
    if (ch == '\'' || ch == '"') {
      attribute_quote = ch;
      continue;
    }
    if (raw_xml.substr(i, 3U) != "ref") {
      continue;
    }
    const char before = i == name_end ? ' ' : raw_xml[i - 1U];
    const char after = i + 3U < opening_end ? raw_xml[i + 3U] : ' ';
    const bool before_is_space = before == ' ' || before == '\t' || before == '\r' || before == '\n';
    const bool after_is_space = after == ' ' || after == '\t' || after == '\r' || after == '\n' || after == '=';
    if (!before_is_space || !after_is_space) {
      continue;
    }
    std::size_t equal = i + 3U;
    while (equal < opening_end &&
           (raw_xml[equal] == ' ' || raw_xml[equal] == '\t' || raw_xml[equal] == '\r' || raw_xml[equal] == '\n')) {
      ++equal;
    }
    if (equal >= opening_end || raw_xml[equal] != '=') {
      continue;
    }
    ++equal;
    while (equal < opening_end &&
           (raw_xml[equal] == ' ' || raw_xml[equal] == '\t' || raw_xml[equal] == '\r' || raw_xml[equal] == '\n')) {
      ++equal;
    }
    if (equal >= opening_end) {
      continue;
    }
    const char value_quote = raw_xml[equal];
    const bool quoted = value_quote == '\'' || value_quote == '"';
    const std::size_t value_start = quoted ? equal + 1U : equal;
    std::size_t value_end = value_start;
    if (quoted) {
      value_end = raw_xml.find(value_quote, value_start);
      if (value_end == std::string_view::npos || value_end > opening_end) {
        continue;
      }
    } else {
      while (value_end < opening_end && raw_xml[value_end] != ' ' && raw_xml[value_end] != '\t' &&
             raw_xml[value_end] != '\r' && raw_xml[value_end] != '\n' && raw_xml[value_end] != '/') {
        ++value_end;
      }
    }
    std::string out(raw_xml);
    out.replace(value_start, value_end - value_start, escaped_ref);
    return out;
  }

  // No ref attribute: insert it immediately before the closing `>` (or the
  // self-closing slash) without touching any existing attributes.
  std::size_t insert_at = opening_end;
  while (insert_at > name_end && (raw_xml[insert_at - 1U] == ' ' || raw_xml[insert_at - 1U] == '\t' ||
                                  raw_xml[insert_at - 1U] == '\r' || raw_xml[insert_at - 1U] == '\n')) {
    --insert_at;
  }
  if (insert_at > name_end && raw_xml[insert_at - 1U] == '/') {
    --insert_at;
  }
  std::string out(raw_xml);
  out.insert(insert_at, " ref=\"" + escaped_ref + "\"");
  return out;
}

}  // namespace

extern "C" fm_status_t fm_workbook_table_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                            const char** out_display_name, const char** out_ref,
                                            size_t* out_sheet_index) {
  clear_last_error();
  if (wb == nullptr || out_name == nullptr || out_display_name == nullptr || out_ref == nullptr ||
      out_sheet_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_table_at: NULL argument");
  }
  const auto& tables = wb->workbook().tables();
  if (auto rc = check_index(idx, tables.size(), "fm_workbook_table_at", "idx"); rc != 0) {
    return rc;
  }
  *out_name = tables[idx].name.c_str();
  *out_display_name = tables[idx].display_name.c_str();
  *out_ref = tables[idx].ref.c_str();
  *out_sheet_index = tables[idx].sheet_index;
  return 0;
}

extern "C" fm_status_t fm_workbook_table_create(fm_workbook_t* wb, size_t sheet_index, const char* ref,
                                                const char* name, const char* display_name,
                                                const char* const* column_names, size_t column_count,
                                                const char* style_name, int32_t header_row, int32_t totals_row,
                                                size_t* out_index) {
  clear_last_error();
  if (wb == nullptr || ref == nullptr || name == nullptr || display_name == nullptr || column_names == nullptr ||
      out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_table_create: NULL argument");
  }
  formulon::Workbook& book = wb->workbook();
  if (sheet_index >= book.sheet_count() || ref[0] == '\0' || name[0] == '\0' || display_name[0] == '\0' ||
      column_count == 0U) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_create: invalid sheet, table identity, range, or columns");
  }
  if (table_ref_column_count(ref) != column_count) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_create: column_count does not match the range width",
                             "ref=" + std::string(ref) + " column_count=" + std::to_string(column_count));
  }
  if (auto rc = check_table_names(book, name, display_name, book.tables().size(),
                                  "fm_workbook_table_create: table name already exists");
      rc != 0) {
    return rc;
  }
  formulon::TableMetadata table;
  table.sheet_index = sheet_index;
  table.name = name;
  table.display_name = display_name;
  table.ref = ref;
  table.header_row = header_row != 0;
  table.totals_row = totals_row != 0;
  uint32_t next_id = 1;
  for (const formulon::TableMetadata& existing : book.tables()) {
    next_id = std::max(next_id, existing.id + 1U);
  }
  table.id = next_id;
  table.columns.reserve(column_count);
  std::unordered_set<std::string> column_keys;
  column_keys.reserve(column_count);
  for (size_t i = 0; i < column_count; ++i) {
    if (column_names[i] == nullptr || column_names[i][0] == '\0') {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_table_create: column name is empty");
    }
    std::string column_key = ascii_lower(column_names[i]);
    if (!column_keys.insert(std::move(column_key)).second) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_table_create: column name is duplicated");
    }
    formulon::TableColumn column;
    column.id = static_cast<uint32_t>(i + 1U);
    column.name = column_names[i];
    table.columns.push_back(std::move(column));
  }
  table.auto_filter_xml.set(
      formulon::io::auto_filter_from_xml("<autoFilter ref=\"" + xml_attr_escape(table.ref) + "\"/>"));
  table.style_info = default_style_info(style_name != nullptr ? style_name : "");
  auto& tables = book.mutable_tables();
  tables.push_back(std::move(table));
  *out_index = tables.size() - 1U;
  // A structured reference can already exist in a formula authored ahead of
  // its table (Excel resolves it once the table appears, surfacing #NAME?
  // until then, exactly like a forward-referenced defined name), so the new
  // table's name and display name both need the same scoped reindex a
  // defined-name edit gets.
  const formulon::TableMetadata& created = tables.back();
  book.reindex_formulas_for_table_change({created.name, created.display_name});
  return 0;
}

extern "C" fm_status_t fm_workbook_table_update(fm_workbook_t* wb, size_t index, const char* ref,
                                                const char* style_name, int32_t header_row, int32_t totals_row) {
  clear_last_error();
  if (wb == nullptr || ref == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_table_update: NULL argument");
  }
  auto& tables = wb->workbook().mutable_tables();
  if (index >= tables.size() || ref[0] == '\0') {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_update: table index or ref is invalid");
  }
  formulon::TableMetadata& table = tables[index];
  if (!table.columns.empty() && table_ref_column_count(ref) != table.columns.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_update: ref width does not match the table's column count",
                             "ref=" + std::string(ref) + " column_count=" + std::to_string(table.columns.size()));
  }

  // Compute every replacement before committing any metadata. In
  // particular, a raw style payload survives a NULL style_name, while an
  // empty string explicitly removes it.
  const std::string next_ref = ref;
  formulon::AutoFilterSlot next_auto_filter = table.auto_filter_xml;
  const std::optional<formulon::MergeRange> next_range = formulon::parse_a1_rectangle(next_ref);
  if (formulon::AutoFilter* filter = next_auto_filter.get();
      filter != nullptr && !filter->is_opaque() && next_range.has_value()) {
    filter->range = *next_range;
  } else {
    next_auto_filter.set(formulon::io::auto_filter_from_xml(
        table_auto_filter_xml(formulon::io::auto_filter_xml(next_auto_filter.get()), next_ref)));
  }
  formulon::TableStyleInfo next_style = table.style_info;
  if (style_name != nullptr) {
    if (style_name[0] == '\0') {
      next_style = formulon::TableStyleInfo{};
    } else if (next_style.present) {
      next_style.name = style_name;
    } else {
      next_style = default_style_info(style_name);
    }
  }
  const bool next_header_row = header_row < 0 ? table.header_row : header_row > 0;
  const bool next_totals_row = totals_row < 0 ? table.totals_row : totals_row > 0;
  const bool totals_turned_on = next_totals_row && !table.totals_row;

  table.ref = next_ref;
  table.header_row = next_header_row;
  table.totals_row = next_totals_row;
  if (totals_turned_on) {
    // Shown now, so the "never shown" marker no longer applies.
    table.totals_row_shown.reset();
  }
  table.auto_filter_xml = std::move(next_auto_filter);
  table.style_info = std::move(next_style);
  // The `ref` just changed, so every structured reference into this table
  // resolves to a different rectangle (`eval/dep_extractor.cpp`); reindex
  // the formulas that name it rather than leave their dep-graph edges and
  // cached values pointing at the pre-edit extent.
  wb->workbook().reindex_formulas_for_table_change({table.name, table.display_name});
  wb->workbook().mark_row_visibility_dependents_dirty();
  if (totals_turned_on) {
    for (std::size_t column = 0; column < table.columns.size(); ++column) {
      const formulon::TableColumn& col = table.columns[column];
      if (col.totals_function.empty() && col.totals_label.empty()) {
        continue;
      }
      if (auto rc = write_totals_cell(wb->workbook(), table, column); rc != 0) {
        return rc;
      }
    }
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_table_remove(fm_workbook_t* wb, size_t index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_table_remove: wb is NULL");
  }
  auto& tables = wb->workbook().mutable_tables();
  if (index >= tables.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_remove: table index out of range");
  }
  // Capture identity before erasing: a formula naming the removed table now
  // resolves to nothing (Excel surfaces #NAME? / #REF! at the structured
  // reference), so it needs the same scoped reindex a table create/update
  // triggers, using the name it can no longer find.
  const std::string removed_name = tables[index].name;
  const std::string removed_display_name = tables[index].display_name;
  formulon::remove_table_slicers(wb->workbook(), tables[index].id);
  tables.erase(tables.begin() + static_cast<std::ptrdiff_t>(index));
  wb->workbook().reindex_formulas_for_table_change({removed_name, removed_display_name});
  // The removed table's AutoFilter may have been the sheet's only one with
  // criteria.
  wb->workbook().mark_row_visibility_dependents_dirty();
  return 0;
}

extern "C" fm_status_t fm_workbook_table_get_info(const fm_workbook_t* wb, size_t index, fm_table_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_table_get_info: NULL argument");
  }
  const auto& tables = wb->workbook().tables();
  if (auto rc = check_index(index, tables.size(), "fm_workbook_table_get_info", "index"); rc != 0) {
    return rc;
  }
  const formulon::TableMetadata& table = tables[index];
  const formulon::TableStyleInfo& style = table.style_info;
  *out = fm_table_info_t{};
  out->style_name = style.name.c_str();
  out->style_present = style.present ? 1 : 0;
  out->show_first_column = style.present && style.show_first_column ? 1 : 0;
  out->show_last_column = style.present && style.show_last_column ? 1 : 0;
  out->show_row_stripes = style.present && style.show_row_stripes ? 1 : 0;
  out->show_column_stripes = style.present && style.show_column_stripes ? 1 : 0;
  out->header_row_count = table.header_row ? 1U : 0U;
  out->totals_row_count = table.totals_row ? 1U : 0U;
  out->column_count = static_cast<uint32_t>(table.columns.size());
  return 0;
}

extern "C" fm_status_t fm_workbook_table_get_column(const fm_workbook_t* wb, size_t index, size_t column,
                                                    fm_table_column_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_table_get_column: NULL argument");
  }
  const auto& tables = wb->workbook().tables();
  if (auto rc = check_index(index, tables.size(), "fm_workbook_table_get_column", "index"); rc != 0) {
    return rc;
  }
  const auto& columns = tables[index].columns;
  if (auto rc = check_index(column, columns.size(), "fm_workbook_table_get_column", "column"); rc != 0) {
    return rc;
  }
  const formulon::TableColumn& col = columns[column];
  out->name = col.name.c_str();
  out->totals_row_function = col.totals_function.empty() ? "none" : col.totals_function.c_str();
  out->totals_row_label = col.totals_label.c_str();
  out->totals_row_formula = col.totals_row_formula.c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_table_set_style_info(fm_workbook_t* wb, size_t index, const char* style_name,
                                                        int32_t show_first_column, int32_t show_last_column,
                                                        int32_t show_row_stripes, int32_t show_column_stripes) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_table_set_style_info: wb is NULL");
  }
  auto& tables = wb->workbook().mutable_tables();
  if (auto rc = check_index(index, tables.size(), "fm_workbook_table_set_style_info", "index"); rc != 0) {
    return rc;
  }
  formulon::TableStyleInfo& style = tables[index].style_info;
  style.present = true;
  style.name = style_name != nullptr ? style_name : "";
  style.show_first_column = show_first_column != 0;
  style.show_last_column = show_last_column != 0;
  style.show_row_stripes = show_row_stripes != 0;
  style.show_column_stripes = show_column_stripes != 0;
  return 0;
}

extern "C" fm_status_t fm_workbook_table_rename(fm_workbook_t* wb, size_t index, const char* name,
                                                const char* display_name) {
  clear_last_error();
  if (wb == nullptr || name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_table_rename: NULL argument");
  }
  formulon::Workbook& book = wb->workbook();
  if (auto rc = check_index(index, book.tables().size(), "fm_workbook_table_rename", "index"); rc != 0) {
    return rc;
  }
  const std::string next_name = name;
  const std::string next_display_name = display_name != nullptr ? display_name : name;
  if (auto rc = check_table_names(book, next_name, next_display_name, index,
                                  "fm_workbook_table_rename: table name is empty or already exists");
      rc != 0) {
    return rc;
  }
  return finish_status(book.rename_table(index, next_name, next_display_name));
}

extern "C" fm_status_t fm_workbook_table_set_column_totals(fm_workbook_t* wb, size_t index, size_t column,
                                                           const char* function, const char* label,
                                                           const char* formula) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_table_set_column_totals: wb is NULL");
  }
  formulon::Workbook& book = wb->workbook();
  auto& tables = book.mutable_tables();
  if (auto rc = check_index(index, tables.size(), "fm_workbook_table_set_column_totals", "index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(column, tables[index].columns.size(), "fm_workbook_table_set_column_totals", "column");
      rc != 0) {
    return rc;
  }
  const std::string_view token = function != nullptr ? function : "none";
  if (find_totals_function(token) == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_set_column_totals: unknown totals-row function",
                             "function=" + std::string(token));
  }
  std::string stored_formula;
  if (token == "custom") {
    if (formula == nullptr || formula[0] == '\0') {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_table_set_column_totals: custom function needs a formula");
    }
    if (auto rc = check_stored_formula("fm_workbook_table_set_column_totals", formula,
                                       "column=" + std::to_string(column), &stored_formula);
        rc != 0) {
      return rc;
    }
  }
  formulon::TableColumn& col = tables[index].columns[column];
  col.totals_function = token == "none" ? std::string() : std::string(token);
  col.totals_label = label != nullptr ? label : "";
  col.totals_row_formula = std::move(stored_formula);
  const formulon::TableMetadata& table = tables[index];
  return table.totals_row ? write_totals_cell(book, table, column) : 0;
}
