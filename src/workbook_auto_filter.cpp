//
// Workbook AutoFilter mutation for sheets and tables, keeping the hidden
// `_FilterDatabase` defined name and the row-visibility dependents in step.

#include <cstddef>
#include <cstdint>
#include <mutex>
#include <string>
#include <string_view>
#include <utility>

#include "auto_filter.h"
#include "defined_name.h"
#include "eval/recalc_engine.h"
#include "io/auto_filter_xml.h"
#include "parser/ast_format.h"
#include "sheet.h"
#include "table.h"
#include "utils/a1_column.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"
#include "workbook_sheet_mutation.h"

namespace formulon {
namespace {
using detail::find_defined_name;
using detail::kFilterDatabaseName;

// `op`-prefixed range check for a mutator addressing one table.
Expected<void, Error> check_table_index(const char* op, std::size_t table_index, std::size_t table_count) {
  if (table_index >= table_count) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op) + ": table_index out of range",
                      "table_index=" + std::to_string(table_index));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

namespace {

/// `Sheet1!$B$2:$F$12`, quoting the sheet name where Excel would.
std::string filter_database_formula(std::string_view sheet_name, const MergeRange& range) {
  std::string out;
  parser::append_sheet_name(sheet_name, false, out);
  out.push_back('!');
  const auto append_cell = [&out](std::uint32_t row, std::uint32_t col) {
    out.push_back('$');
    a1::append_column_letters(out, col);
    out.push_back('$');
    out.append(std::to_string(row + 1U));
  };
  append_cell(range.first_row, range.first_col);
  out.push_back(':');
  append_cell(range.last_row, range.last_col);
  return out;
}

}  // namespace

Expected<void, Error> Workbook::set_sheet_auto_filter(std::size_t sheet_index, AutoFilter filter) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_sheet_index("set_sheet_auto_filter", sheet_index, sheets_.size()));
  RETURN_IF_ERROR(validate_auto_filter(filter));
  std::string formula = filter_database_formula(sheets_[sheet_index].name(), filter.range);
  sheets_[sheet_index].set_auto_filter(std::move(filter));
  const auto scope = static_cast<std::int32_t>(sheet_index);
  const auto it = find_defined_name(defined_names_, kFilterDatabaseName, scope);
  if (it != defined_names_.end()) {
    it->formula = std::move(formula);
    it->hidden = true;
  } else {
    DefinedName entry;
    entry.name = std::string(kFilterDatabaseName);
    entry.formula = std::move(formula);
    entry.local_sheet_id = scope;
    entry.hidden = true;
    defined_names_.push_back(std::move(entry));
  }
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  reindex_formulas_referencing_name(sheets_, mutator, *this, kFilterDatabaseName);
  mark_row_visibility_dependents_dirty_locked(sheets_, mutator, defined_names_);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_sheet_auto_filter(std::size_t sheet_index) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_sheet_index("remove_sheet_auto_filter", sheet_index, sheets_.size()));
  sheets_[sheet_index].clear_auto_filter_model();
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  if (erase_filter_database_name(defined_names_, sheet_index)) {
    reindex_formulas_referencing_name(sheets_, mutator, *this, kFilterDatabaseName);
  }
  mark_row_visibility_dependents_dirty_locked(sheets_, mutator, defined_names_);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_sheet_auto_filter_xml(std::size_t sheet_index, std::string_view xml) {
  if (xml.empty()) {
    return remove_sheet_auto_filter(sheet_index);
  }
  Expected<AutoFilter, Error> parsed = io::parse_auto_filter_xml(xml);
  if (parsed && validate_auto_filter(parsed.value())) {
    return set_sheet_auto_filter(sheet_index, std::move(parsed.value()));
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_sheet_index("set_sheet_auto_filter_xml", sheet_index, sheets_.size()));
  sheets_[sheet_index].set_auto_filter(io::auto_filter_from_xml(xml));
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator(), defined_names_);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_table_auto_filter(std::size_t table_index, AutoFilter filter) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_table_index("set_table_auto_filter", table_index, tables_.size()));
  RETURN_IF_ERROR(validate_auto_filter(filter));
  tables_[table_index].auto_filter_xml.set(std::move(filter));
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator(), defined_names_);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_table_auto_filter(std::size_t table_index) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_table_index("remove_table_auto_filter", table_index, tables_.size()));
  tables_[table_index].auto_filter_xml.reset();
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator(), defined_names_);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
