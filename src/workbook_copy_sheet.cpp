//
// `Workbook::copy_sheet`: duplicates a worksheet the way Excel's Move or Copy
// does. Every fallible step (validation, naming, table and name planning,
// drawing snapshots) runs before the first mutation, so a refusal leaves the
// workbook untouched. The copy is appended, its formulas re-pointed through a
// `SheetCopyTransform`, then moved into place under one hold of the engine
// mutex.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "drawing/drawing_clone.h"
#include "drawing/slicer_edit.h"
#include "eval/recalc_engine.h"
#include "parser/ref_transforms.h"
#include "sheet.h"
#include "sheet_copy_names.h"
#include "sheet_name.h"
#include "table.h"
#include "unknown_relationship.h"
#include "utils/derived_guid.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"
#include "workbook_ref_rewrite.h"

namespace formulon {
namespace {

Error unsupported(const char* what) {
  return make_error(FormulonErrorCode::kNotImplemented, "sheet content cannot be copied", what);
}

/// A worksheet's `.xlsb` binary index: optional, rebuilt by Excel, so a copy goes without it.
bool is_binary_index(const UnknownRelationship& rel) {
  constexpr std::string_view kPrefix = "xl/worksheets/binaryIndex";
  return !rel.target_external && rel.target.size() > kPrefix.size() && strings::starts_with(rel.target, kPrefix);
}

std::string rewritten(const std::string& formula, const parser::RefTransform& transform) {
  FormulaRewriteResult result = rewrite_formula(formula, transform);
  return result.changed ? std::move(result.text) : formula;
}

}  // namespace

Expected<std::size_t, Error> Workbook::copy_sheet(std::size_t source, std::size_t insert_before,
                                                  std::optional<std::string> new_name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (source >= sheets_.size() || insert_before > sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "index out of range",
                      "source=" + std::to_string(source) + " insert_before=" + std::to_string(insert_before) +
                          " sheet_count=" + std::to_string(sheets_.size()));
  }
  RETURN_IF_ERROR(check_sheet_capacity());
  const Sheet& src = sheets_[source];
  if (src.view().visibility() == SheetVisibility::kVeryHidden) {
    return make_error(FormulonErrorCode::kPreconditionFailed,
                      "source_index names a very hidden sheet, which cannot be copied",
                      "source=" + std::to_string(source));
  }
  if (src.is_opaque_ooxml_sheet()) {
    return unsupported("not a worksheet");
  }
  if (!std::all_of(src.unknown_relationships().begin(), src.unknown_relationships().end(), is_binary_index)) {
    return unsupported("unrecognised relationship");
  }
  if (!framed_slicers(*this, source).empty()) {
    return unsupported("slicer");
  }

  std::string copy_name;
  if (new_name.has_value()) {
    RETURN_IF_ERROR(check_entered_new_sheet_name(*new_name));
    copy_name = std::move(*new_name);
  } else {
    std::vector<std::string_view> existing;
    existing.reserve(sheets_.size());
    for (const Sheet& sheet : sheets_) {
      existing.push_back(sheet.name());
    }
    copy_name = copy_sheet_default_name(src.name(), existing);
  }
  const std::string source_name = src.name();
  const std::size_t n = sheets_.size();

  // Tables one at a time in workbook order, each name counting the clones before it.
  std::vector<TableMetadata> table_copies;
  std::vector<std::pair<std::string, std::string>> table_map;
  std::uint32_t next_id = 0;
  struct NameScope {
    const Workbook* workbook;
    const std::vector<TableMetadata>* table_copies;
  };
  NameScope scope{this, &table_copies};
  const auto name_taken = [](void* context, std::string_view candidate) {
    const auto& [workbook, copies] = *static_cast<const NameScope*>(context);
    const auto table_named = [candidate](const TableMetadata& table) {
      return sheet_names::write_collides(table.name, candidate) ||
             sheet_names::write_collides(table.display_name, candidate);
    };
    return std::any_of(workbook->tables_.begin(), workbook->tables_.end(), table_named) ||
           std::any_of(copies->begin(), copies->end(), table_named) ||
           std::any_of(
               workbook->defined_names_.begin(), workbook->defined_names_.end(),
               [candidate](const DefinedName& entry) { return sheet_names::write_collides(entry.name, candidate); });
  };
  for (const TableMetadata& table : tables_) {
    if (table.sheet_index != source) {
      continue;
    }
    TableMetadata copy = table;
    const std::string renamed =
        copy_table_name(table.display_name, tables_.size() + table_copies.size(), name_taken, &scope);
    copy.name = renamed;
    copy.display_name = renamed;
    next_id = next_free_table_id(next_id);
    copy.id = next_id;
    copy.sheet_index = n;
    rederive_attr_guid(copy.root_extra_attrs, "xr:uid", copy_name);
    for (TableColumn& column : copy.columns) {
      rederive_attr_guid(column.extra_attrs, "xr3:uid", copy_name);
    }
    table_map.emplace_back(table.display_name, renamed);
    table_copies.push_back(std::move(copy));
  }
  const sheet_names::NameList sheet_list = detail::sheet_name_list(sheets_);
  const sheet_names::NameList table_list = sheet_names::table_name_list(tables_);
  const parser::SheetCopyTransform transform(source_name, copy_name, std::move(table_map), &sheet_list, &table_list);

  // Source-local names, then workbook names reaching the source that no local of the same name shadows.
  std::vector<DefinedName> name_copies;
  const auto source_id = static_cast<std::int32_t>(source);
  for (const DefinedName& entry : defined_names_) {
    if (entry.local_sheet_id == source_id) {
      DefinedName copy = entry;
      copy.formula = rewritten(entry.formula, transform);
      copy.local_sheet_id = static_cast<std::int32_t>(n);
      name_copies.push_back(std::move(copy));
    }
  }
  for (const DefinedName& entry : defined_names_) {
    if (entry.local_sheet_id >= 0) {
      continue;
    }
    FormulaRewriteResult result = rewrite_formula(entry.formula, transform);
    if (!result.changed) {
      continue;
    }
    const bool shadowed = sheet_names::resolve_by(defined_names_.size(), [&](std::size_t i, sheet_names::Fold fold) {
                            return defined_names_[i].local_sheet_id == source_id &&
                                   sheet_names::same(defined_names_[i].name, entry.name, fold);
                          }) != sheet_names::kNotFound;
    if (shadowed) {
      continue;
    }
    DefinedName copy = entry;
    copy.formula = std::move(result.text);
    copy.local_sheet_id = static_cast<std::int32_t>(n);
    name_copies.push_back(std::move(copy));
  }

  auto drawing_plan = plan_drawing_copy(*this, source);
  if (!drawing_plan) {
    return drawing_plan.error();
  }

  // Mutation starts here.
  sheets_.push_back(sheets_[source].clone_as(copy_name));
  {
    std::vector<UnknownRelationship> rels = sheets_[n].unknown_relationships();
    rels.erase(std::remove_if(rels.begin(), rels.end(), is_binary_index), rels.end());
    sheets_[n].set_unknown_relationships(std::move(rels));
  }
  for (TableMetadata& table : table_copies) {
    tables_.push_back(std::move(table));
  }

  struct CellUpdate {
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    std::string formula;
  };
  std::vector<CellUpdate> cell_updates;
  for (const auto& [row, cells] : sheets_[n].rows()) {
    for (const auto& [col, cell] : cells) {
      if (cell.formula_text.empty()) {
        continue;
      }
      FormulaRewriteResult result = rewrite_formula(cell.formula_text, transform);
      if (result.changed) {
        cell_updates.push_back(CellUpdate{row, col, std::move(result.text)});
      }
    }
  }
  for (CellUpdate& update : cell_updates) {
    sheets_[n].set_cell_formula(update.row, update.col, std::move(update.formula));
  }

  const parser::IdentityTransform identity;
  std::vector<const parser::RefTransform*> per_sheet(sheets_.size(), &identity);
  per_sheet[n] = &transform;
  rewrite_sheet_metadata_formulas(sheets_, per_sheet, tables_, pivot_caches_, {}, {}, {}, identity);

  for (DefinedName& entry : name_copies) {
    defined_names_.push_back(std::move(entry));
  }

  if (auto applied = apply_drawing_copy(*this, n, drawing_plan.value(), transform); !applied) {
    (void)remove_sheet_locked(static_cast<std::uint32_t>(n));
    return applied.error();
  }

  // Either path ends in one full reindex, which also dirties every formula the new names may reach.
  if (insert_before != n) {
    move_sheet_locked(static_cast<std::uint32_t>(n), static_cast<std::uint32_t>(insert_before));
  } else {
    reindex_all_formulas(sheets_, engine_->locked_mutator(), *this);
  }
  return insert_before;
}

}  // namespace formulon
