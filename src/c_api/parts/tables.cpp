//
// C ABI - workbook tables: read-side iteration and create / update / remove.

#include <algorithm>
#include <cctype>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "io/xml_escape.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::clear_last_error;
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
  std::uint32_t first_row = 0;
  std::uint32_t first_col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
  const std::size_t colon = ref.find(':');
  if (colon == std::string_view::npos) {
    return formulon::a1::parse_a1_ref(ref, &first_row, &first_col) ? 1U : 0U;
  }
  if (!formulon::a1::parse_a1_ref(ref.substr(0, colon), &first_row, &first_col) ||
      !formulon::a1::parse_a1_ref(ref.substr(colon + 1), &last_row, &last_col)) {
    return 0U;
  }
  if (last_col < first_col || last_row < first_row) {
    return 0U;
  }
  return static_cast<std::size_t>(last_col - first_col) + 1U;
}

std::string table_style_xml(std::string_view style_name) {
  if (style_name.empty()) {
    return {};
  }
  return "<tableStyleInfo name=\"" + xml_attr_escape(style_name) +
         "\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>";
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
  for (std::size_t i = name_end; i + 3U <= opening_end; ++i) {
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
  if (idx >= tables.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_table_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(tables.size()));
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
  const std::string name_key = ascii_lower(name);
  for (const formulon::TableMetadata& existing : book.tables()) {
    if (ascii_lower(existing.name) == name_key || ascii_lower(existing.display_name) == name_key ||
        ascii_lower(existing.name) == ascii_lower(display_name) ||
        ascii_lower(existing.display_name) == ascii_lower(display_name)) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_table_create: table name already exists");
    }
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
  for (size_t i = 0; i < column_count; ++i) {
    if (column_names[i] == nullptr || column_names[i][0] == '\0') {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_table_create: column name is empty");
    }
    const std::string column_key = ascii_lower(column_names[i]);
    for (const formulon::TableColumn& existing : table.columns) {
      if (ascii_lower(existing.name) == column_key) {
        return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                                 "fm_workbook_table_create: column name is duplicated");
      }
    }
    formulon::TableColumn column;
    column.id = static_cast<uint32_t>(i + 1U);
    column.name = column_names[i];
    table.columns.push_back(std::move(column));
  }
  table.auto_filter_xml = "<autoFilter ref=\"" + xml_attr_escape(table.ref) + "\"/>";
  table.table_style_info_xml = table_style_xml(style_name != nullptr ? style_name : "");
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
    next_auto_filter = table_auto_filter_xml(next_auto_filter.xml(), next_ref);
  }
  const std::string next_style_xml = style_name == nullptr ? table.table_style_info_xml : table_style_xml(style_name);
  const bool next_header_row = header_row < 0 ? table.header_row : header_row > 0;
  const bool next_totals_row = totals_row < 0 ? table.totals_row : totals_row > 0;

  table.ref = next_ref;
  table.header_row = next_header_row;
  table.totals_row = next_totals_row;
  table.auto_filter_xml = std::move(next_auto_filter);
  table.table_style_info_xml = next_style_xml;
  // The `ref` just changed, so every structured reference into this table
  // resolves to a different rectangle (`eval/dep_extractor.cpp`); reindex
  // the formulas that name it rather than leave their dep-graph edges and
  // cached values pointing at the pre-edit extent.
  wb->workbook().reindex_formulas_for_table_change({table.name, table.display_name});
  wb->workbook().mark_row_visibility_dependents_dirty();
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
  tables.erase(tables.begin() + static_cast<std::ptrdiff_t>(index));
  wb->workbook().reindex_formulas_for_table_change({removed_name, removed_display_name});
  // The removed table's AutoFilter may have been the sheet's only one with
  // criteria.
  wb->workbook().mark_row_visibility_dependents_dirty();
  return 0;
}
