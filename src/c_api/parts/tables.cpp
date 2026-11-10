//
// C ABI - workbook tables: read-side iteration and create / update / remove.

#include <algorithm>
#include <cassert>
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
using formulon::c_api::parts::fresh_read_scratch;
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
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::TextStore;
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

/// One column patch after its checks: what `fm_sheet_update_table` merges into the column's settings once the table's
/// shape is settled and any totals defaults are written.
struct ColumnPlan {
  std::size_t column = 0;
  /// Each totals field's op and the value a `FM_PATCH_OP_SET` gives it (an empty function is `none`); a set formula is
  /// bound where it is stored.
  fm_patch_op_t function_op = FM_PATCH_OP_KEEP;
  std::string totals_function;
  fm_patch_op_t label_op = FM_PATCH_OP_KEEP;
  std::string totals_label;
  fm_patch_op_t formula_op = FM_PATCH_OP_KEEP;
  std::string totals_formula;
  fm_patch_op_t calculated_formula_op = FM_PATCH_OP_KEEP;
  std::string calculated_formula;
  /// The mark `calculated_formula` gives the column, as the engine stores it.
  std::string calculated_mark;
};

/// The table's range and header / totals rows once the rest of the patch is applied.
struct NextShape {
  formulon::MergeRange rect;
  bool header_row = false;
  bool totals_row = false;
  /// Whether showing the totals row replaces every column's totals settings with the defaults first.
  bool totals_defaults = false;
};

/// Checks `patch->columns` against the table at global index `index` and its shape after the patch, with each entry
/// taken as merged into its column's settings (the defaults' when `next.totals_defaults`), and fills `plans`; nothing
/// changes. A formula is parsed at the next totals cell (or the next last data row without a totals row); it is bound
/// where it is stored, once the table has its final shape.
fm_status_t plan_column_patches(const fm_workbook_t* wb, std::size_t index, const fm_table_patch_t* patch,
                                const NextShape& next, std::vector<ColumnPlan>* plans) {
  if (patch->columns_op != FM_PATCH_OP_SET || patch->columns_count == 0U) {
    return 0;
  }
  if (patch->columns == nullptr) {
    return null_field("columns");
  }
  const formulon::Workbook& book = wb->workbook();
  const formulon::TableMetadata& table = book.tables()[index];
  if (patch->columns_count > table.columns.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kIndexOutOfRange,
        ("columns_count " + std::to_string(patch->columns_count) + " is out of range (" +
         std::to_string(table.columns.size()) + " columns)")
            .c_str(),
        "columns_count=" + std::to_string(patch->columns_count) + " count=" + std::to_string(table.columns.size()));
  }
  const std::uint32_t first_data_row = next.rect.first_row + (next.header_row ? 1U : 0U);
  const std::uint32_t data_end = next.rect.last_row + 1U - (next.totals_row ? 1U : 0U);
  // A table with no data row passes the empty span 1..0.
  const std::uint32_t span_first = data_end > first_data_row ? first_data_row : 1U;
  const std::uint32_t span_last = data_end > first_data_row ? data_end - 1U : 0U;
  plans->reserve(patch->columns_count);
  for (std::size_t i = 0; i < patch->columns_count; ++i) {
    const fm_table_column_patch_t& entry = patch->columns[i];
    const formulon::TableColumn& current = table.columns[i];
    const std::string at = "columns[" + std::to_string(i) + "].";
    for (const auto& [op, field] :
         {std::pair<fm_patch_op_t, const char*>{entry.totals_row_function_op, "totals_row_function"},
          {entry.totals_row_label_op, "totals_row_label"},
          {entry.totals_row_formula_op, "totals_row_formula"},
          {entry.calculated_formula_op, "calculated_formula"}}) {
      if (auto rc = check_patch_op(op, (at + field).c_str(), true); rc != 0) {
        return rc;
      }
    }
    ColumnPlan plan;
    plan.column = i;
    plan.function_op = entry.totals_row_function_op;
    plan.label_op = entry.totals_row_label_op;
    plan.formula_op = entry.totals_row_formula_op;
    if (entry.totals_row_function_op == FM_PATCH_OP_SET) {
      if (auto rc = check_enum_domain(entry.totals_row_function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE,
                                      FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, (at + "totals_row_function").c_str());
          rc != 0) {
        return rc;
      }
      const std::string_view token = token_of(kTotalsRowFunctionTokens, entry.totals_row_function);
      plan.totals_function = token == "none" ? std::string() : std::string(token);
    }
    // A text to set is not `NULL` and not empty; clearing is the op, not an empty string.
    const auto text_to_set = [&](fm_patch_op_t op, const char* text, const char* field,
                                 std::string* out) -> fm_status_t {
      if (op == FM_PATCH_OP_CLEAR) {
        out->clear();
        return 0;
      }
      if (op != FM_PATCH_OP_SET) {
        return 0;
      }
      if (text == nullptr) {
        return null_field(at + field);
      }
      if (auto rc = check_utf8({text}); rc != 0) {
        return rc;
      }
      if (text[0] == '\0') {
        return invalid_argument((at + field + " must not be empty").c_str(), std::string(field) + "=");
      }
      *out = text;
      return 0;
    };
    if (auto rc =
            text_to_set(entry.totals_row_label_op, entry.totals_row_label, "totals_row_label", &plan.totals_label);
        rc != 0) {
      return rc;
    }
    if (auto rc = text_to_set(entry.totals_row_formula_op, entry.totals_row_formula, "totals_row_formula",
                              &plan.totals_formula);
        rc != 0) {
      return rc;
    }
    // The defaults set no custom function and no formula, so under them the settings read as empty.
    bool custom = !next.totals_defaults && current.totals_function == "custom";
    bool has_formula = !next.totals_defaults && !current.totals_row_formula.empty();
    if (plan.function_op != FM_PATCH_OP_KEEP) {
      custom = plan.function_op == FM_PATCH_OP_SET && plan.totals_function == "custom";
    }
    if (plan.formula_op != FM_PATCH_OP_KEEP) {
      has_formula = plan.formula_op == FM_PATCH_OP_SET;
    }
    if (custom) {
      if (!has_formula) {
        return invalid_argument((at + "totals_row_formula is required for a custom function").c_str(),
                                "totals_row_formula=");
      }
      if (plan.formula_op == FM_PATCH_OP_SET) {
        std::string checked;
        if (auto rc = check_entered_cell_formula(wb, table.sheet_index, next.rect.last_row,
                                                 next.rect.first_col + static_cast<std::uint32_t>(i),
                                                 plan.totals_formula, "field=" + at + "totals_row_formula", &checked);
            rc != 0) {
          return rc;
        }
      }
    }
    plan.calculated_formula_op = entry.calculated_formula_op;
    std::string calculated;
    if (auto rc = text_to_set(entry.calculated_formula_op, entry.calculated_formula, "calculated_formula", &calculated);
        rc != 0) {
      return rc;
    }
    if (entry.calculated_formula_op == FM_PATCH_OP_SET) {
      auto mark = book.check_table_calculated_formula(table.sheet_index, index, i, calculated, span_first, span_last);
      if (!mark) {
        return set_last_error(mark.error());
      }
      plan.calculated_formula = std::move(calculated);
      plan.calculated_mark = std::move(mark.value());
    }
    plans->push_back(std::move(plan));
  }
  return 0;
}

/// Applies the column patches `plan_column_patches` accepted, once the table has its final shape, each merged into its
/// column's settings as they stand then. A setting that equals the column's current one changes nothing.
fm_status_t apply_column_plans(fm_workbook_t* wb, std::size_t index, const std::vector<ColumnPlan>& plans) {
  formulon::Workbook& book = wb->workbook();
  const auto merged = [](fm_patch_op_t op, const std::string& value, const std::string& current) {
    return op == FM_PATCH_OP_SET ? value : op == FM_PATCH_OP_CLEAR ? std::string() : current;
  };
  for (const ColumnPlan& plan : plans) {
    formulon::TableMetadata& table = book.mutable_tables()[index];
    formulon::TableColumn& col = table.columns[plan.column];
    const std::string totals_function = merged(plan.function_op, plan.totals_function, col.totals_function);
    const std::string totals_label = merged(plan.label_op, plan.totals_label, col.totals_label);
    // A function other than `custom` keeps no formula.
    std::string stored_formula =
        totals_function == "custom" ? merged(plan.formula_op, plan.totals_formula, col.totals_row_formula) : "";
    if (totals_function == "custom" && plan.formula_op == FM_PATCH_OP_SET) {
      // Entered at the column's cell on the table's last row, so `[Qty]` and `[@Qty]` bind to this table.
      const std::optional<formulon::MergeRange> rect = formulon::parse_a1_rectangle(table.ref);
      if (!rect) {
        return invalid_argument("ref is not a range", "ref=" + table.ref);
      }
      std::string bound;
      if (auto rc = check_entered_cell_formula(wb, table.sheet_index, rect->last_row,
                                               rect->first_col + static_cast<std::uint32_t>(plan.column),
                                               plan.totals_formula, "field=totals_row_formula", &bound);
          rc != 0) {
        return rc;
      }
      stored_formula = book.bind_self_book_names(
          formulon::Workbook::normalize_formula_text(std::move(bound), formulon::Workbook::FormulaSource::kEntered));
      book.bind_external_books(std::string_view(stored_formula));
      stored_formula = std::string(formulon::strip_formula_prefix(stored_formula));
    }
    if (col.totals_function != totals_function || col.totals_label != totals_label ||
        col.totals_row_formula != stored_formula) {
      col.totals_function = totals_function;
      col.totals_label = totals_label;
      col.totals_row_formula = std::move(stored_formula);
      if (table.totals_row) {
        if (auto rc = finish_status(book.write_table_totals_cell(index, plan.column)); rc != 0) {
          return rc;
        }
      }
    }
    if (plan.calculated_formula_op == FM_PATCH_OP_SET) {
      if (plan.calculated_mark != col.calculated_column_formula) {
        if (auto rc = finish_status(
                book.set_table_calculated_column(table.sheet_index, index, plan.column, plan.calculated_formula));
            rc != 0) {
          return rc;
        }
      }
    } else if (plan.calculated_formula_op == FM_PATCH_OP_CLEAR && !col.calculated_column_formula.empty()) {
      if (auto rc = finish_status(book.clear_table_calculated_column(table.sheet_index, index, plan.column)); rc != 0) {
        return rc;
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
                                  {patch->totals_row_op, "totals_row"},
                                  {patch->columns_op, "columns"}}) {
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
  formulon::Workbook& book = wb->workbook();
  const formulon::TableMetadata& table = book.tables()[index];
  const std::optional<formulon::MergeRange> current_rect = formulon::parse_a1_rectangle(table.ref);
  if (!current_rect) {
    return invalid_argument("ref is not a range", "ref=" + table.ref);
  }
  std::optional<formulon::MergeRange> range;
  std::string range_ref;
  if (patch->range_op == FM_PATCH_OP_SET) {
    if (auto rc = table_ref_text(patch->range, &range_ref); rc != 0) {
      return rc;
    }
    if (!table.columns.empty() && table_ref_column_count(range_ref) != table.columns.size()) {
      return invalid_argument("range width must match the table's column count",
                              "range=" + range_ref + " column_count=" + std::to_string(table.columns.size()));
    }
    range = to_merge_range(patch->range);
  }
  const bool header_shown = patch->header_row == 1;
  const bool totals_shown = patch->totals_row == 1;
  const bool header_toggle = patch->header_row_op == FM_PATCH_OP_SET && header_shown != table.header_row;
  const bool totals_toggle = patch->totals_row_op == FM_PATCH_OP_SET && totals_shown != table.totals_row;

  // Every refusal is checked before the first change, in application order: the header toggle, the totals toggle,
  // the range, the column patches. Each check takes the shape and the cell shifts the steps before it leave.
  formulon::TableShape shape{*current_rect, table.header_row, table.totals_row};
  std::vector<formulon::PendingShift> pending;
  if (header_toggle) {
    auto shift = book.check_table_header_row(index, header_shown, shape, pending);
    if (!shift) {
      return set_last_error(shift.error());
    }
    if (!header_shown) {
      shape.rect.first_row += 1U;
    } else if (shift.value().has_value()) {
      shape.rect.last_row += 1U;
      pending.push_back(*shift.value());
    } else {
      shape.rect.first_row -= 1U;
    }
    shape.header_row = header_shown;
  }
  if (totals_toggle) {
    auto shift = book.check_table_totals_row(index, totals_shown, shape, pending);
    if (!shift) {
      return set_last_error(shift.error());
    }
    shape.rect.last_row = totals_shown ? shape.rect.last_row + 1U : shape.rect.last_row - 1U;
    if (shift.value().has_value()) {
      pending.push_back(*shift.value());
    }
    shape.totals_row = totals_shown;
  }
  // The range is the table's range after the toggles, whose first row they may have moved.
  if (range.has_value()) {
    if (range->first_row != shape.rect.first_row) {
      return invalid_argument("patch.range must start on the table's first row after the header toggle",
                              "range=" + range_ref + " first_row=" + std::to_string(shape.rect.first_row));
    }
    if (auto checked = book.check_table_resize(index, *range, shape, pending); !checked) {
      return finish_status(checked);
    }
  }
  NextShape next_shape;
  next_shape.rect = range.value_or(shape.rect);
  next_shape.header_row = shape.header_row;
  next_shape.totals_row = shape.totals_row;
  next_shape.totals_defaults = totals_toggle && totals_shown && !table.totals_row_shown.value_or(true);
  std::vector<ColumnPlan> column_plans;
  if (auto rc = plan_column_patches(wb, index, patch, next_shape, &column_plans); rc != 0) {
    return rc;
  }

  // Past the checks every step succeeds; a failure here is a check that missed a refusal.
  if (header_toggle) {
    const formulon::Expected<void, formulon::Error> toggled = book.set_table_header_row(index, header_shown);
    assert(toggled.has_value());
    if (auto rc = finish_status(toggled); rc != 0) {
      return rc;
    }
  }
  if (totals_toggle) {
    const formulon::Expected<void, formulon::Error> toggled = book.set_table_totals_row(index, totals_shown);
    assert(toggled.has_value());
    if (auto rc = finish_status(toggled); rc != 0) {
      return rc;
    }
  }
  if (range.has_value() && !(*range == shape.rect)) {
    // The first row the new range adds below the data rows the toggles left.
    const std::uint32_t first_new_row = shape.rect.last_row + (shape.totals_row ? 0U : 1U);
    const formulon::Expected<void, formulon::Error> resized = book.resize_table(index, *range);
    assert(resized.has_value());
    if (auto rc = finish_status(resized); rc != 0) {
      return rc;
    }
    // An AutoFilter the model could not parse keeps its text; only its ref follows the header and data rows.
    formulon::TableMetadata& shaped = book.mutable_tables()[index];
    if (const formulon::AutoFilter* filter = shaped.auto_filter_xml.get(); filter != nullptr && filter->is_opaque()) {
      formulon::MergeRange filter_rect = *range;
      if (shaped.totals_row && filter_rect.last_row > filter_rect.first_row) {
        filter_rect.last_row -= 1U;
      }
      std::string filter_ref;
      formulon::append_a1_rectangle(filter_ref, filter_rect);
      shaped.auto_filter_xml.set(
          formulon::io::auto_filter_from_xml(table_auto_filter_xml(formulon::io::auto_filter_xml(filter), filter_ref)));
    }
    const formulon::Expected<void, formulon::Error> inferred =
        book.infer_table_calculated_columns(shaped.sheet_index, index, first_new_row);
    assert(inferred.has_value());
    if (auto rc = finish_status(inferred); rc != 0) {
      return rc;
    }
  }
  const fm_status_t applied = apply_column_plans(wb, index, column_plans);
  assert(applied == 0);
  return applied;
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
  // Both formulas are shown from inside the table: the totals formula as in its totals row, the calculated-column
  // formula as in the column's first data row.
  TextStore& scratch = fresh_read_scratch(wb);
  const std::string& own_table = tables[index].display_name;
  out->totals_row_formula =
      col.totals_row_formula.empty()
          ? nullptr
          : store_cstr(scratch,
                       with_formula_prefix(display_formula_text(wb->workbook(), col.totals_row_formula, own_table)));
  out->calculated_formula = col.calculated_column_formula.empty()
                                ? nullptr
                                : store_cstr(scratch, with_formula_prefix(display_formula_text(
                                                          wb->workbook(), col.calculated_column_formula, own_table)));
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
