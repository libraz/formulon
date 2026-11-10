//
// C ABI - workbook tables: read-side iteration and create / update / remove.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <tuple>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "c_api/parts/table_style_defaults.h"
#include "drawing/slicer_edit.h"
#include "eval/text_format/display_text.h"
#include "io/auto_filter_xml.h"
#include "io/xml_escape.h"
#include "sheet_name.h"
#include "table.h"
#include "utils/error.h"
#include "utils/formula_text.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::c_str_or_null;
using formulon::c_api::parts::check_entered_cell_formula;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_patch_op;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::display_formula_text;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::kTotalsRowFunctionTokens;
using formulon::c_api::parts::local_table_index;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::null_field;
using formulon::c_api::parts::ordinal_of;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::resolve_sheet_table;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::to_cell_rect;
using formulon::c_api::parts::to_merge_range;
using formulon::c_api::parts::token_of;
using formulon::c_api::parts::with_formula_prefix;

namespace {

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

/// The A1 `ref` of a table covering `rect`, which must be ordered and inside the grid.
fm_status_t table_ref_text(const fm_cell_rect_t& rect, std::string* out) {
  if (auto rc = check_sheet_rect(rect); rc != 0) {
    return rc;
  }
  out->clear();
  formulon::append_a1_rectangle(*out, to_merge_range(rect));
  return 0;
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

/// Column names of a new table over `rect`. Without `column_names` each name is
/// its header cell's displayed text, empty for a blank header (every column of
/// a table without a header row), which `Workbook::add_table` then names. With
/// them, a non-blank header cell must already read as its column's name.
/// `*blank` marks the header cells the names are to be written into.
fm_status_t table_column_names(const formulon::Workbook& book, std::size_t sheet_index, const fm_cell_rect_t& rect,
                               bool header_row, const char* const* column_names, std::size_t column_count,
                               std::vector<std::string>* names, std::vector<bool>* blank) {
  const std::size_t width = static_cast<std::size_t>(rect.last_col - rect.first_col) + 1U;
  if (column_names != nullptr) {
    names->reserve(column_count);
    for (std::size_t i = 0; i < column_count; ++i) {
      names->push_back(text_or_empty(column_names[i]));
    }
  } else {
    names->assign(width, std::string());
  }
  blank->assign(names->size(), false);
  if (header_row) {
    const formulon::Sheet& sheet = book.sheet(sheet_index);
    for (std::size_t i = 0; i < names->size() && i < width; ++i) {
      std::string header = formulon::text_format::format_cell_for_display(
                               book, sheet, rect.first_row, rect.first_col + static_cast<std::uint32_t>(i))
                               .text;
      if (header.empty()) {
        (*blank)[i] = true;
      } else if (column_names == nullptr) {
        (*names)[i] = std::move(header);
      } else if ((*names)[i] != header) {
        return invalid_argument(("column_names[" + std::to_string(i) + "] must match its header cell").c_str(),
                                "column_names[" + std::to_string(i) + "]=" + (*names)[i] + " header=" + header);
      }
    }
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_table_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_count = static_cast<std::size_t>(
      std::count_if(wb->workbook().tables().begin(), wb->workbook().tables().end(),
                    [sheet_index](const formulon::TableMetadata& table) { return table.sheet_index == sheet_index; }));
  return 0;
}

extern "C" fm_status_t fm_sheet_get_table_at(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                             fm_table_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  const formulon::TableMetadata& table = wb->workbook().tables()[index];
  const std::optional<formulon::MergeRange> rect = formulon::parse_a1_rectangle(table.ref);
  if (!rect) {
    return invalid_argument("ref is not a range", "ref=" + table.ref);
  }
  const formulon::TableStyleInfo& style = table.style_info;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_table_info_t{};
  out->struct_size = struct_size;
  out->name = table.display_name.c_str();
  out->range = to_cell_rect(*rect);
  out->style_name = style.present && !style.name.empty() ? style.name.c_str() : nullptr;
  out->show_first_column = style.present && style.show_first_column ? 1 : 0;
  out->show_last_column = style.present && style.show_last_column ? 1 : 0;
  out->show_row_stripes = style.present && style.show_row_stripes ? 1 : 0;
  out->show_column_stripes = style.present && style.show_column_stripes ? 1 : 0;
  out->header_row_count = table.header_row ? 1U : 0U;
  out->totals_row_count = table.totals_row ? 1U : 0U;
  out->column_count = static_cast<uint32_t>(table.columns.size());
  return 0;
}

extern "C" fm_status_t fm_workbook_find_table(const fm_workbook_t* wb, const char* name, int32_t* out_present,
                                              size_t* out_sheet_index, size_t* out_table_index) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || out_present == nullptr || out_sheet_index == nullptr ||
      out_table_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const auto& tables = wb->workbook().tables();
  const std::size_t i =
      formulon::sheet_names::resolve_by(tables.size(), [&](std::size_t at, formulon::sheet_names::Fold fold) {
        return formulon::sheet_names::same(tables[at].display_name, name, fold);
      });
  if (i != formulon::sheet_names::kNotFound) {
    *out_sheet_index = tables[i].sheet_index;
    *out_table_index = local_table_index(wb->workbook(), i);
    *out_present = 1;
    return 0;
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_table_runtime_id(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                     uint32_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  std::size_t global = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &global); rc != 0) {
    return rc;
  }
  formulon::RuntimeId& id = wb->workbook().tables()[global].runtime_id;
  if (id.value() == 0U) {
    id.assign(wb->workbook().allocate_runtime_id());
  }
  *out = id.value();
  return 0;
}

extern "C" fm_status_t fm_workbook_resolve_table_runtime_id(const fm_workbook_t* wb, uint32_t id,
                                                            size_t* out_sheet_index, size_t* out_table_index) {
  clear_last_error();
  if (wb == nullptr || out_sheet_index == nullptr || out_table_index == nullptr) {
    return null_argument();
  }
  const auto& tables = wb->workbook().tables();
  // Zero is the id of a table nobody has asked about, so it names none.
  for (std::size_t i = 0; id != 0U && i < tables.size(); ++i) {
    if (tables[i].runtime_id.value() == id) {
      *out_sheet_index = tables[i].sheet_index;
      *out_table_index = local_table_index(wb->workbook(), i);
      return 0;
    }
  }
  return not_found(("id " + std::to_string(id) + " names no table").c_str(), "id=" + std::to_string(id));
}

extern "C" fm_status_t fm_sheet_add_table(fm_workbook_t* wb, size_t sheet_index, const fm_table_spec_t* spec,
                                          size_t* out_table_index) {
  clear_last_error();
  if (wb == nullptr || spec == nullptr || out_table_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*spec); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (spec->name == nullptr) {
    return null_argument("name");
  }
  const fm_cell_rect_t& rect = spec->range;
  std::string ref;
  if (auto rc = table_ref_text(rect, &ref); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({spec->name, spec->style_name}); rc != 0) {
    return rc;
  }
  const char* const* column_names = spec->column_names;
  for (size_t i = 0; column_names != nullptr && i < spec->column_count; ++i) {
    if (column_names[i] == nullptr) {
      continue;
    }
    if (auto rc = check_utf8(column_names[i], "column", i); rc != 0) {
      return rc;
    }
  }
  for (const auto& [value, field] : {std::pair<int32_t, const char*>{spec->styled, "styled"},
                                     {spec->header_row, "header_row"},
                                     {spec->totals_row, "totals_row"},
                                     {spec->show_first_column, "show_first_column"},
                                     {spec->show_last_column, "show_last_column"},
                                     {spec->show_row_stripes, "show_row_stripes"},
                                     {spec->show_column_stripes, "show_column_stripes"}}) {
    if (auto rc = check_enum_domain(value, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, field); rc != 0) {
      return rc;
    }
  }
  if (spec->style_name != nullptr && spec->style_name[0] == '\0') {
    return invalid_argument("style_name must not be empty", "style_name=");
  }
  const bool styled = spec->styled != FM_TRISTATE_FALSE;
  if (!styled && spec->style_name != nullptr) {
    return invalid_argument("style_name must be NULL when styled is false", "styled=false");
  }
  const bool header_row = spec->header_row != FM_TRISTATE_FALSE;
  formulon::Workbook& book = wb->workbook();
  std::vector<std::string> names;
  std::vector<bool> blank_headers;
  if (auto rc = table_column_names(book, sheet_index, rect, header_row, column_names, spec->column_count, &names,
                                   &blank_headers);
      rc != 0) {
    return rc;
  }
  formulon::TableMetadata table;
  table.sheet_index = sheet_index;
  table.name = spec->name;
  table.display_name = spec->name;
  table.ref = ref;
  table.header_row = header_row;
  table.totals_row = spec->totals_row == FM_TRISTATE_TRUE;
  table.columns.reserve(names.size());
  for (const std::string& column_name : names) {
    formulon::TableColumn column;
    column.name = column_name;
    table.columns.push_back(std::move(column));
  }
  if (styled) {
    table.style_info =
        default_style_info(spec->style_name != nullptr ? std::string_view(spec->style_name)
                                                       : formulon::c_api::parts::default_table_style_name(book));
  }
  const std::pair<int32_t, bool formulon::TableStyleInfo::*> toggles[] = {
      {spec->show_first_column, &formulon::TableStyleInfo::show_first_column},
      {spec->show_last_column, &formulon::TableStyleInfo::show_last_column},
      {spec->show_row_stripes, &formulon::TableStyleInfo::show_row_stripes},
      {spec->show_column_stripes, &formulon::TableStyleInfo::show_column_stripes}};
  for (const auto& [value, member] : toggles) {
    if (value == FM_TRISTATE_DEFAULT) {
      continue;
    }
    // A stated toggle on an unstyled table creates the element, as the style-options patch does.
    if (!table.style_info.present) {
      table.style_info.present = true;
      table.style_info.show_row_stripes = true;
    }
    table.style_info.*member = value == FM_TRISTATE_TRUE;
  }
  auto added = book.add_table(std::move(table));
  if (!added) {
    return formulon::c_api::parts::set_last_error(added.error());
  }
  for (std::size_t i = 0; i < blank_headers.size(); ++i) {
    if (!blank_headers[i]) {
      continue;
    }
    const std::string& name = book.tables()[added.value()].columns[i].name;
    if (auto rc = finish_status(
            book.set_cell_text(sheet_index, rect.first_row, rect.first_col + static_cast<std::uint32_t>(i), name));
        rc != 0) {
      return rc;
    }
  }
  *out_table_index = local_table_index(book, added.value());
  return 0;
}

extern "C" fm_status_t fm_sheet_update_table(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                             const fm_table_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || patch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  // The whole patch is checked before anything changes; no field is clearable.
  for (const auto& [op, field] : {std::pair<fm_patch_op_t, const char*>{patch->range_op, "range"},
                                  {patch->header_row_op, "header_row"},
                                  {patch->totals_row_op, "totals_row"}}) {
    if (auto rc = check_patch_op(op, field, false); rc != 0) {
      return rc;
    }
  }
  for (const auto& [op, value, field] :
       {std::tuple<fm_patch_op_t, int32_t, const char*>{patch->header_row_op, patch->header_row, "header_row"},
        {patch->totals_row_op, patch->totals_row, "totals_row"}}) {
    if (op == FM_PATCH_OP_SET) {
      if (auto rc = check_enum_domain(value, 0, 1, field); rc != 0) {
        return rc;
      }
    }
  }
  const bool range_set = patch->range_op == FM_PATCH_OP_SET;
  // Without a range, the totals toggle resizes the table; it checks everything
  // before changing anything, so a refusal leaves the whole patch unapplied.
  if (!range_set && patch->totals_row_op == FM_PATCH_OP_SET) {
    if (auto rc = finish_status(wb->workbook().set_table_totals_row(index, patch->totals_row == 1)); rc != 0) {
      return rc;
    }
  }
  formulon::TableMetadata& table = wb->workbook().mutable_tables()[index];
  std::string next_ref = table.ref;
  formulon::AutoFilterSlot next_auto_filter = table.auto_filter_xml;
  if (range_set) {
    if (auto rc = table_ref_text(patch->range, &next_ref); rc != 0) {
      return rc;
    }
    if (!table.columns.empty() && table_ref_column_count(next_ref) != table.columns.size()) {
      return invalid_argument("range width must match the table's column count",
                              "range=" + next_ref + " column_count=" + std::to_string(table.columns.size()));
    }
    const std::optional<formulon::MergeRange> next_range = formulon::parse_a1_rectangle(next_ref);
    if (next_range) {
      if (auto checked = wb->workbook().check_table_resize(index, *next_range); !checked) {
        return finish_status(checked);
      }
    }
    if (formulon::AutoFilter* filter = next_auto_filter.get();
        filter != nullptr && !filter->is_opaque() && next_range.has_value()) {
      filter->range = *next_range;
    } else {
      next_auto_filter.set(formulon::io::auto_filter_from_xml(
          table_auto_filter_xml(formulon::io::auto_filter_xml(next_auto_filter.get()), next_ref)));
    }
  }
  const bool next_header_row = patch->header_row_op == FM_PATCH_OP_SET ? patch->header_row == 1 : table.header_row;
  const bool next_totals_row = patch->totals_row_op == FM_PATCH_OP_SET ? patch->totals_row == 1 : table.totals_row;
  if (!range_set && next_header_row == table.header_row) {
    return 0;
  }
  const bool totals_turned_on = range_set && next_totals_row && !table.totals_row;

  table.ref = std::move(next_ref);
  table.header_row = next_header_row;
  table.totals_row = next_totals_row;
  if (totals_turned_on) {
    // Shown now, so the "never shown" marker no longer applies.
    table.totals_row_shown.reset();
  }
  if (range_set) {
    table.auto_filter_xml = std::move(next_auto_filter);
  }
  // The table's extent or row layout changed, so every structured reference
  // into it may resolve differently (`eval/dep_extractor.cpp`); reindex the
  // formulas that name it rather than leave their dep-graph edges and cached
  // values pointing at the pre-edit shape.
  wb->workbook().reindex_formulas_for_table_change({table.name, table.display_name});
  wb->workbook().mark_row_visibility_dependents_dirty();
  if (totals_turned_on) {
    for (std::size_t column = 0; column < table.columns.size(); ++column) {
      const formulon::TableColumn& col = table.columns[column];
      if (col.totals_function.empty() && col.totals_label.empty()) {
        continue;
      }
      if (auto rc = finish_status(wb->workbook().write_table_totals_cell(index, column)); rc != 0) {
        return rc;
      }
    }
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_table(fm_workbook_t* wb, size_t sheet_index, size_t table_index) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  auto& tables = wb->workbook().mutable_tables();
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

extern "C" fm_status_t fm_sheet_get_table_column_at(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                    size_t column_index, fm_table_column_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  const auto& tables = wb->workbook().tables();
  const auto& columns = tables[index].columns;
  if (auto rc = check_index(column_index, columns.size(), "column_index"); rc != 0) {
    return rc;
  }
  const formulon::TableColumn& col = columns[column_index];
  out->name = col.name.c_str();
  out->totals_row_function =
      ordinal_of(kTotalsRowFunctionTokens, col.totals_function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  out->totals_row_label = c_str_or_null(col.totals_label);
  // The formula sits in the table's totals row, so it is shown from inside the table.
  out->totals_row_formula =
      col.totals_row_formula.empty()
          ? nullptr
          : publish_scratch(wb, with_formula_prefix(display_formula_text(wb->workbook(), col.totals_row_formula,
                                                                         tables[index].display_name)));
  return 0;
}

extern "C" fm_status_t fm_sheet_get_table_style_options(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                        int32_t* out_present, fm_table_style_options_t* out) {
  clear_last_error();
  if (wb == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  const formulon::TableStyleInfo& style = wb->workbook().tables()[index].style_info;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_table_style_options_t{};
  out->struct_size = struct_size;
  out->style_name = c_str_or_null(style.name);
  *out_present = style.present ? 1 : 0;
  out->show_first_column = style.present && style.show_first_column ? 1 : 0;
  out->show_last_column = style.present && style.show_last_column ? 1 : 0;
  out->show_row_stripes = style.present && style.show_row_stripes ? 1 : 0;
  out->show_column_stripes = style.present && style.show_column_stripes ? 1 : 0;
  return 0;
}

namespace {

/// Checks one toggle of a table-style patch: a known op, no `CLEAR`, a 0 / 1 value under `SET`.
fm_status_t check_toggle_patch(fm_patch_op_t op, int32_t value, const char* field) {
  if (auto rc = check_patch_op(op, field, false); rc != 0) {
    return rc;
  }
  if (op == FM_PATCH_OP_SET) {
    return check_enum_domain(value, 0, 1, field);
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_update_table_style_options(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                           const fm_table_style_options_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || patch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  // The whole patch is checked before anything changes.
  if (auto rc = check_patch_op(patch->style_name_op, "style_name", true); rc != 0) {
    return rc;
  }
  if (patch->style_name_op == FM_PATCH_OP_SET) {
    if (patch->style_name == nullptr) {
      return null_field("style_name");
    }
    if (patch->style_name[0] == '\0') {
      return invalid_argument("style_name must not be empty", "style_name=");
    }
    if (auto rc = check_utf8({patch->style_name}); rc != 0) {
      return rc;
    }
  }
  for (const auto& [op, value, field] :
       {std::tuple<fm_patch_op_t, int32_t, const char*>{patch->show_first_column_op, patch->show_first_column,
                                                        "show_first_column"},
        {patch->show_last_column_op, patch->show_last_column, "show_last_column"},
        {patch->show_row_stripes_op, patch->show_row_stripes, "show_row_stripes"},
        {patch->show_column_stripes_op, patch->show_column_stripes, "show_column_stripes"}}) {
    if (auto rc = check_toggle_patch(op, value, field); rc != 0) {
      return rc;
    }
  }
  if (patch->style_name_op == FM_PATCH_OP_KEEP && patch->show_first_column_op == FM_PATCH_OP_KEEP &&
      patch->show_last_column_op == FM_PATCH_OP_KEEP && patch->show_row_stripes_op == FM_PATCH_OP_KEEP &&
      patch->show_column_stripes_op == FM_PATCH_OP_KEEP) {
    return 0;
  }
  formulon::TableStyleInfo& style = wb->workbook().mutable_tables()[index].style_info;
  if (!style.present) {
    style = default_style_info("");
    style.present = true;
    style.show_row_stripes = true;
  }
  if (patch->style_name_op == FM_PATCH_OP_SET) {
    style.name = patch->style_name;
  } else if (patch->style_name_op == FM_PATCH_OP_CLEAR) {
    style.name.clear();
  }
  if (patch->show_first_column_op == FM_PATCH_OP_SET) {
    style.show_first_column = patch->show_first_column != 0;
  }
  if (patch->show_last_column_op == FM_PATCH_OP_SET) {
    style.show_last_column = patch->show_last_column != 0;
  }
  if (patch->show_row_stripes_op == FM_PATCH_OP_SET) {
    style.show_row_stripes = patch->show_row_stripes != 0;
  }
  if (patch->show_column_stripes_op == FM_PATCH_OP_SET) {
    style.show_column_stripes = patch->show_column_stripes != 0;
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_rename_table(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                             const char* name) {
  clear_last_error();
  if (wb == nullptr || name == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  formulon::Workbook& book = wb->workbook();
  const std::string next_name = name;
  return finish_status(book.rename_table(index, next_name, next_name));
}

extern "C" fm_status_t fm_sheet_rename_table_column(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                    size_t column_index, const char* name, const char** out_name) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || out_name == nullptr) {
    return null_argument();
  }
  formulon::Workbook& book = wb->workbook();
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(column_index, book.tables()[index].columns.size(), "column_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  auto renamed = book.rename_table_column(index, column_index, name);
  if (!renamed) {
    return set_last_error(renamed.error());
  }
  *out_name = publish_scratch(wb, std::move(renamed.value()));
  return 0;
}

extern "C" fm_status_t fm_sheet_insert_table_rows(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                  uint32_t first, uint32_t last) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().insert_table_rows(sheet_index, index, first, last));
}

extern "C" fm_status_t fm_sheet_insert_table_columns(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                     uint32_t first, uint32_t last) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().insert_table_columns(sheet_index, index, first, last));
}

extern "C" fm_status_t fm_sheet_set_table_column_totals(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                        size_t column_index, int32_t function, const char* label,
                                                        const char* formula) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  formulon::Workbook& book = wb->workbook();
  std::size_t index = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &index); rc != 0) {
    return rc;
  }
  auto& tables = book.mutable_tables();
  if (auto rc = check_index(column_index, tables[index].columns.size(), "column_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({label}); rc != 0) {
    return rc;
  }
  if ((label != nullptr && label[0] == '\0') || (formula != nullptr && formula[0] == '\0')) {
    return invalid_argument("label and formula must not be empty");
  }
  if (auto rc = check_enum_domain(function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM,
                                  "function");
      rc != 0) {
    return rc;
  }
  const std::string_view token = token_of(kTotalsRowFunctionTokens, function);
  std::string stored_formula;
  if (token == "custom") {
    if (formula == nullptr) {
      return invalid_argument("formula is required for a custom function");
    }
    // Entered at the column's totals cell, so `[Qty]` and `[@Qty]` bind to this table.
    const std::optional<formulon::MergeRange> rect = formulon::parse_a1_rectangle(tables[index].ref);
    if (!rect) {
      return invalid_argument("ref is not a range", "ref=" + tables[index].ref);
    }
    std::string bound;
    if (auto rc = check_entered_cell_formula(wb, tables[index].sheet_index, rect->last_row,
                                             rect->first_col + static_cast<std::uint32_t>(column_index), formula,
                                             "field=formula", &bound);
        rc != 0) {
      return rc;
    }
    stored_formula = book.bind_self_book_names(
        formulon::Workbook::normalize_formula_text(std::move(bound), formulon::Workbook::FormulaSource::kEntered));
    book.bind_external_books(std::string_view(stored_formula));
    stored_formula = std::string(formulon::strip_formula_prefix(stored_formula));
  }
  formulon::TableColumn& col = tables[index].columns[column_index];
  col.totals_function = token == "none" ? std::string() : std::string(token);
  col.totals_label = text_or_empty(label);
  col.totals_row_formula = std::move(stored_formula);
  const formulon::TableMetadata& table = tables[index];
  return table.totals_row ? finish_status(book.write_table_totals_cell(index, column_index)) : 0;
}
