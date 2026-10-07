//
// Dependency-graph re-indexing shared by the `Workbook` mutators: every
// structural change (sheet add/rename/remove/move, defined-name and table
// edits, row/column edits) re-registers the formulas it can affect.
// Internal to the workbook implementation; not part of the public API.
//
// Precondition for every function taking a `LockedMutator`: the caller holds
// the engine mutex for the mutator's lifetime.

#ifndef FORMULON_WORKBOOK_FORMULA_INDEX_H_
#define FORMULON_WORKBOOK_FORMULA_INDEX_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "defined_name.h"
#include "eval/dep_graph.h"
#include "eval/recalc_engine.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"

namespace formulon {

class Workbook;

parser::AstNode* parse_indexable_formula(std::string_view body, Arena& arena);
parser::AstNode* parse_stored_formula(std::string_view text, Arena& arena);
eval::CellNodeId make_node(std::size_t sheet_index, std::uint32_t row, std::uint32_t col);

void reindex_all_formulas(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                          const Workbook& workbook);
bool references_any_name(const parser::AstNode& root, const std::vector<std::string>& candidates);
std::vector<std::string> close_affected_names(const std::vector<DefinedName>& defined_names,
                                              std::vector<std::string> affected);

// Re-registers and dirties only the formula cells whose value can change
// when `changed_name` (and every name that transitively references it, per
// `collect_affected_names`) is added, retargeted, or removed. Unlike
// `reindex_all_formulas`, this leaves every unaffected formula's dep-graph
// edges, cached value, and spill geometry untouched, so setting a print
// area/print titles/unrelated name no longer wipes spills workbook-wide or
// forces every formula to recompute. Shared by every scoped-reindex caller
// (defined names, table structure) via `affected`, which tests the
// already-parsed AST; the parse itself cannot be skipped for the unaffected
// majority (no per-name/per-table dependent index exists in the recalc
// engine to look this up directly), so cost is proportional to the
// workbook's formula count regardless of how few are actually affected. It
// no longer re-registers dep-graph edges, resets the graph, or clears
// spills for that unaffected majority, unlike `reindex_all_formulas`.
template <typename Affected>
void reindex_formulas_if(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                         const Workbook& workbook, const Affected& affected) {
  Arena parser_arena;
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    Sheet& sheet = sheets[sheet_idx];
    for (const auto& [row, cells] : sheet.rows()) {
      for (std::size_t col = 0; col < cells.size(); ++col) {
        const Cell& cell = cells[col];
        if (cell.formula_text.empty()) {
          continue;
        }
        parser::AstNode* root = parse_stored_formula(cell.formula_text, parser_arena);
        if (root == nullptr || !affected(*root)) {
          continue;
        }
        const eval::CellNodeId node{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)};
        // No-op when `(row, col)` is not a spill anchor.
        sheet.clear_spill(row, static_cast<std::uint32_t>(col));
        mutator.register_formula(node, *root, workbook);
        mutator.mark_dirty(node);
      }
    }
  }
}

void reindex_formulas_referencing_name(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                                       const Workbook& workbook, std::string_view changed_name);
void reindex_formulas_referencing_sheet(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                                        const Workbook& workbook, std::string_view changed_sheet);
void reindex_formulas_referencing_table(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                                        const Workbook& workbook, const std::vector<std::string>& table_names);
void mark_row_visibility_dependents_dirty_locked(const std::vector<Sheet>& sheets,
                                                 const eval::RecalcEngine::LockedMutator& mutator);

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_FORMULA_INDEX_H_
