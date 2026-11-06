//
// Defined names for `Workbook`: replacing and editing the name table with the
// formula re-indexing that follows, and the erasers for the hidden names an
// AutoFilter or slicer keeps.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "eval/recalc_engine.h"
#include "io/xlsb/table_records.h"
#include "parser/ast.h"
#include "parser/external_book_spelling.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/name_syntax.h"
#include "utils/strings.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"
#include "workbook_ref_rewrite.h"
#include "workbook_sheet_mutation.h"

namespace formulon {

namespace {

using detail::find_defined_name;
using detail::kFilterDatabaseName;

}  // namespace

/// Removes the `_FilterDatabase` name scoped to `sheet_index`. Returns true
/// when one existed.
bool erase_filter_database_name(std::vector<DefinedName>& names, std::size_t sheet_index) {
  const auto it = find_defined_name(names, kFilterDatabaseName, static_cast<std::int32_t>(sheet_index));
  if (it == names.end()) {
    return false;
  }
  names.erase(it);
  return true;
}

/// Removes the workbook-scoped names listed in `cache_names` (slicer cache
/// names, matched case-insensitively). Returns true when any was removed.
bool erase_slicer_cache_names(std::vector<DefinedName>& names, const std::vector<std::string>& cache_names) {
  const auto it = std::remove_if(names.begin(), names.end(), [&cache_names](const DefinedName& entry) {
    return entry.local_sheet_id == -1 &&
           std::any_of(cache_names.begin(), cache_names.end(), [&entry](const std::string& cache_name) {
             return strings::case_insensitive_eq(entry.name, cache_name);
           });
  });
  const bool removed = it != names.end();
  names.erase(it, names.end());
  return removed;
}

void Workbook::set_defined_names(std::vector<DefinedName> names) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (!file_name_.empty()) {
    for (DefinedName& entry : names) {
      entry.formula = parser::bind_self_book_names(entry.formula, file_name_);
    }
  }
  std::vector<std::string> changed;
  {
    // Declaration order matters for duplicate definitions. Metadata and
    // case-only spelling changes do not alter name resolution.
    using Definitions = std::vector<std::pair<std::int32_t, std::string_view>>;
    const auto group = [](const std::vector<DefinedName>& entries) {
      std::unordered_map<std::string, Definitions> grouped;
      for (const DefinedName& entry : entries) {
        grouped[strings::to_ascii_lower(entry.name)].emplace_back(entry.local_sheet_id, entry.formula);
      }
      return grouped;
    };
    const auto before = group(defined_names_);
    const auto after = group(names);
    for (const auto& [name, definitions] : before) {
      const auto found = after.find(name);
      if (found == after.end() || definitions != found->second) {
        changed.push_back(name);
      }
    }
    for (const auto& [name, definitions] : after) {
      if (before.find(name) == before.end()) {
        changed.push_back(name);
      }
    }
  }
  for (const DefinedName& entry : names) {
    bind_external_books_unlocked(std::string_view(entry.formula));
  }
  defined_names_ = std::move(names);
  if (changed.empty()) {
    return;
  }
  const auto affected = close_affected_names(defined_names_, std::move(changed));
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  reindex_formulas_if(sheets_, mutator, *this,
                      [&](const parser::AstNode& root) { return references_any_name(root, affected); });
}

Expected<void, Error> Workbook::set_defined_name(std::string name, std::string formula) {
  return set_defined_name_scoped(std::move(name), std::move(formula), -1);
}

Expected<void, Error> Workbook::set_defined_name_scoped(std::string name, std::string formula,
                                                        std::int32_t local_sheet_id) {
  // Validation, mutation, and the dependent graph rebuild must share one
  // critical section. Otherwise a concurrent sheet removal can change the
  // count between validation and insertion, or a recalc can observe the
  // new definition before its graph is rebuilt.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (name.empty()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_defined_name_scoped: name is empty");
  }
  if (local_sheet_id < -1 || (local_sheet_id >= 0 && static_cast<std::size_t>(local_sheet_id) >= sheets_.size())) {
    return make_error(
        FormulonErrorCode::kInvalidArgument, "set_defined_name_scoped: local_sheet_id out of range",
        "local_sheet_id=" + std::to_string(local_sheet_id) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (!file_name_.empty()) {
    formula = parser::bind_self_book_names(formula, file_name_);
  }
  bind_external_books_unlocked(std::string_view(formula));
  // Case-insensitive lookup: Excel resolves defined names case-folded.
  // Restrict the search to the requested scope.
  if (const auto it = find_defined_name(defined_names_, name, local_sheet_id); it != defined_names_.end()) {
    if (formula.empty()) {
      defined_names_.erase(it);
    } else {
      it->formula = std::move(formula);
    }
    // Retargeting or removing an existing name changes what every formula
    // that references it (directly or through another name that expands
    // to it) resolves to, so their dep-graph edges and cached values are
    // now stale; a formula that never mentions this name is unaffected.
    // Re-register and dirty just that subset, and clear spill geometry
    // only where one of them anchors a spill, rather than the workbook-
    // wide rebuild `reindex_all_formulas` would force.
    const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
    reindex_formulas_referencing_name(sheets_, mutator, *this, name);
    return Expected<void, Error>::Ok();
  }
  // No existing entry; appending only makes sense if a formula is
  // supplied. An empty-formula "remove" against a non-existent name is
  // a successful no-op.
  if (formula.empty()) {
    return Expected<void, Error>::Ok();
  }
  const std::vector<io::xlsb::XlsbRetainedTable> retained = io::xlsb::xlsb_retained_tables(*this);
  const bool retained_table = std::any_of(retained.begin(), retained.end(), [&name](const auto& table) {
    return strings::case_insensitive_eq(table.name, name) || strings::case_insensitive_eq(table.display_name, name);
  });
  if (!is_valid_defined_name(name) || is_table_name(name) || retained_table) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_defined_name_scoped: invalid or taken name",
                      "name=" + name);
  }
  DefinedName entry;
  entry.name = std::move(name);
  entry.formula = std::move(formula);
  entry.local_sheet_id = local_sheet_id;
  defined_names_.push_back(std::move(entry));
  // Adding a name can make an already-calculated #NAME? formula resolvable.
  // Only a formula that actually mentions this name (or another name that
  // expands to it) could have produced that stale #NAME?, so scope the
  // re-register/dirty pass to that subset rather than every formula.
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  reindex_formulas_referencing_name(sheets_, mutator, *this, defined_names_.back().name);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_defined_name_hidden(std::string_view name, std::int32_t local_sheet_id,
                                                        bool hidden) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (const auto it = find_defined_name(defined_names_, name, local_sheet_id); it != defined_names_.end()) {
    it->hidden = hidden;
    return Expected<void, Error>::Ok();
  }
  return make_error(FormulonErrorCode::kInvalidArgument, "set_defined_name_hidden: no such defined name",
                    "name=" + std::string(name) + " local_sheet_id=" + std::to_string(local_sheet_id));
}

Expected<void, Error> Workbook::set_defined_name_comment(std::string_view name, std::int32_t local_sheet_id,
                                                         std::string comment) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (const auto it = find_defined_name(defined_names_, name, local_sheet_id); it != defined_names_.end()) {
    it->comment = std::move(comment);
    return Expected<void, Error>::Ok();
  }
  return make_error(FormulonErrorCode::kInvalidArgument, "set_defined_name_comment: no such defined name",
                    "name=" + std::string(name) + " local_sheet_id=" + std::to_string(local_sheet_id));
}

}  // namespace formulon
