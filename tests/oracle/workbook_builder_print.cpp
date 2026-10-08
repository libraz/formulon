
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
  // Hidden lines are emitted before the size maps on purpose: both the
  // pagination fast path and `ColumnWidthChars` resolve a track against the
  // *first* matching span, so a `column_widths` span covering a hidden
  // column would otherwise mask it and silently paginate the case as if
  // nothing were hidden.
  if (const JsonValue* hidden_v = spec.find("hidden_columns"); hidden_v != nullptr && !hidden_v->is_null()) {
    if (!hidden_v->is_array()) {
      return invalid("'hidden_columns' must be an array");
    }
    for (const JsonValue& item : hidden_v->as_array()) {
      if (!item.is_string()) {
        return invalid("hidden_columns: entries must be column letters");
      }
      const std::string& key = item.as_string();
      std::size_t pos = 0;
      std::uint32_t col1 = 0;
      if (!a1::parse_column_letters(key, &pos, &col1) || pos != key.size()) {
        return invalid("hidden_columns/" + key + ": malformed column key");
      }
      ColumnLayout col;
      col.first = col1 - 1U;
      col.last = col1 - 1U;
      col.hidden = true;
      sheet->mutable_layout().columns.push_back(col);
    }
  }

  if (const JsonValue* hidden_v = spec.find("hidden_rows"); hidden_v != nullptr && !hidden_v->is_null()) {
    if (!hidden_v->is_array()) {
      return invalid("'hidden_rows' must be an array");
    }
    for (const JsonValue& item : hidden_v->as_array()) {
      if (!item.is_string()) {
        return invalid("hidden_rows: entries must be 1-based row numbers");
      }
      const std::string& key = item.as_string();
      std::size_t pos = 0;
      std::uint32_t row1 = 0;
      if (!a1::parse_uint(key, &pos, &row1) || pos != key.size() || row1 == 0U) {
        return invalid("hidden_rows/" + key + ": malformed 1-based row key");
      }
      RowLayout row;
      row.row = row1 - 1U;
      row.hidden = true;
      sheet->mutable_layout().row_overrides.push_back(row);
    }
  }

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
      if (!a1::parse_column_letters(lhs, &p_lhs, &first) || p_lhs != lhs.size() ||
          !a1::parse_column_letters(rhs, &p_rhs, &last) || p_rhs != rhs.size()) {
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
      if (!a1::parse_uint(key, &pos, &row1) || pos != key.size() || row1 == 0U) {
        return invalid("row_heights/" + key + ": malformed 1-based row key");
      }
      RowLayout row;
      row.row = row1 - 1U;
      row.height = value.as_number();
      row.has_height = true;
      row.custom_height = true;
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

Expected<BuiltPrint, Error> build_print_from_spec(const JsonValue& spec, ExcelProfile profile) {
  if (!spec.is_object()) {
    return invalid("workbook spec must be an object");
  }
  const JsonValue* print_v = spec.find("print");
  if (print_v == nullptr || !print_v->is_object()) {
    return invalid("workbook spec has no 'print' block");
  }
  const JsonValue& print = *print_v;

  ASSIGN_OR_RETURN(std::unique_ptr<Workbook> workbook, build_workbook(spec, profile));

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
        if (!a1::parse_column_letters(letters, &pos, &col1) || pos != letters.size() || col1 == 0U) {
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
  std::vector<DefinedName> defined_names = workbook->defined_names();
  if (const JsonValue* area_v = print.find("print_area"); area_v != nullptr && !area_v->is_null()) {
    if (!area_v->is_string()) {
      return invalid("print 'print_area' must be a string");
    }
    DefinedName dn;
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
      DefinedName dn;
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
