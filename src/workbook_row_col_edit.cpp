//
// Row and column insert/delete for `Workbook`: the formula rewrite and
// dependency-graph re-keying around `Sheet`'s cell-store shift, plus the
// spill, AutoFilter and defined-name bookkeeping that follows it.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "defined_name.h"
#include "drawing/drawing_edit.h"
#include "eval/dep_graph.h"
#include "eval/recalc_engine.h"
#include "io/auto_filter_xml.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/ref_transforms.h"
#include "sheet.h"
#include "sheet_name.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_ref_rewrite.h"
#include "workbook_sheet_mutation.h"

namespace formulon {

namespace {

// Computes the post-edit coordinates of a cell on the edited sheet.
// `kept == false` means the cell is dropped by the upcoming
// `Sheet::delete_rows/cols` (deletion fully covers the cell) or pushed
// past the sheet bound by an insert.
struct CellShift {
  bool kept = true;
  std::uint32_t new_row = 0;
  std::uint32_t new_col = 0;
};

CellShift shift_cell_coords_for_row_col_edit(parser::RowColAxis axis, parser::RowColEdit edit, std::uint32_t index,
                                             std::uint32_t count, std::uint32_t row, std::uint32_t col) {
  CellShift r;
  r.new_row = row;
  r.new_col = col;
  std::uint32_t& target = (axis == parser::RowColAxis::kRow) ? r.new_row : r.new_col;
  const std::uint32_t coord = target;
  const std::uint32_t bound = (axis == parser::RowColAxis::kRow) ? Sheet::kMaxRows : Sheet::kMaxCols;

  if (edit == parser::RowColEdit::kInsert) {
    if (coord >= index) {
      const std::uint64_t shifted = static_cast<std::uint64_t>(coord) + count;
      if (shifted >= bound) {
        r.kept = false;
      } else {
        target = static_cast<std::uint32_t>(shifted);
      }
    }
    return r;
  }

  if (coord >= index + count) {
    target = coord - count;
  } else if (coord >= index) {
    r.kept = false;
  }
  return r;
}

// Walks every formula cell in `sheets`, applying a row/col shift
// transform. The transform is rebuilt per sheet so its
// `local_means_target` flag tracks whether the current sheet's
// unqualified references should be rewritten — a local reference on
// the target sheet is in scope, but a local reference on any other
// sheet refers to that other sheet and must be left alone.
//
// The transformed AST drives both the rewritten formula text and the
// dep-graph re-registration; we keep the arena alive across both
// operations so `register_formula` can read the same nodes the
// formatter just emitted, instead of re-parsing the formatted output.
//
// Coordinate handling: cells on the target sheet have their keys shifted
// by the same transform as the AST refs (the physical move happens
// later in `Sheet::insert_rows` / `delete_rows`; here we re-key the
// dep-graph entries and dirty marks so they match the post-shift cell
// position). Without this re-key the dirty marks land at stale
// coordinates and recalc evaluates the wrong cell — visible as blank
// cached values on the band edge and `#REF!` on aggregators that span
// the band.
//
// Precondition: caller holds the engine mutex for the lifetime of the
// `LockedMutator&`. The helper drives a long per-cell loop of dep-graph
// mutations and must share the critical section that the surrounding
// `apply_row_col_edit_operation` takes, so the eventual
// `Sheet::insert_rows` / `delete_rows` does not race against a
// concurrent `recalc_parallel`.
void rewrite_formulas_for_row_col_edit(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                                       const Workbook& workbook, std::string_view target_sheet, parser::RowColAxis axis,
                                       parser::RowColEdit edit, std::uint32_t index, std::uint32_t count) {
  struct FormulaUpdate {
    Sheet* sheet = nullptr;
    eval::CellNodeId old_node;
    eval::CellNodeId new_node;
    bool dropped = false;
    std::string formula;
  };

  // Collect the complete mutation set before touching the graph. A single
  // structural edit can map an old key onto a key that is still occupied by
  // another formula (for example, delete row 1 moves B3 -> B2 while B2 is
  // not yet unregistered). DepGraph::remove_node is coordinate based, so
  // interleaving unregister/register can then delete the newly registered
  // node's edges. The two phases below make every old key absent before any
  // new key is registered, independent of unordered_map iteration order.
  std::vector<FormulaUpdate> updates;
  Arena parser_arena;
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    Sheet& sheet = sheets[sheet_idx];
    const bool local_means_target = sheet_names::equal(sheet.name(), target_sheet);
    const parser::RowColShiftTransform transform(target_sheet, axis, edit, index, count, local_means_target);

    std::vector<std::uint32_t> row_keys;
    row_keys.reserve(sheet.rows().size());
    for (const auto& kv : sheet.rows()) {
      row_keys.push_back(kv.first);
    }
    for (std::uint32_t row : row_keys) {
      const auto it = sheet.rows().find(row);
      if (it == sheet.rows().end()) {
        continue;
      }
      const RowCells& cells = it->second;
      for (std::size_t col = 0; col < cells.size(); ++col) {
        const Cell& cell = cells[col];
        if (cell.formula_text.empty()) {
          continue;
        }

        // Resolve the cell's post-shift coordinates before parsing the formula.
        // An opaque formula still moves with its cell; skipping it on parse
        // failure would leave the dependency graph keyed at the old address.
        std::uint32_t new_row = row;
        std::uint32_t new_col = static_cast<std::uint32_t>(col);
        bool dropped = false;
        if (local_means_target) {
          const CellShift sr =
              shift_cell_coords_for_row_col_edit(axis, edit, index, count, row, static_cast<std::uint32_t>(col));
          if (!sr.kept) {
            dropped = true;
          } else {
            new_row = sr.new_row;
            new_col = sr.new_col;
          }
        }
        const bool cell_moves = (new_row != row) || (new_col != static_cast<std::uint32_t>(col));

        std::string_view body = cell.formula_text;
        bool had_equals = false;
        if (!body.empty() && body.front() == '=') {
          body = body.substr(1);
          had_equals = true;
        }

        bool ast_changed = false;
        std::string effective_formula = cell.formula_text;
        parser_arena.reset();
        if (!body.empty()) {
          parser::AstNode* root = parse_indexable_formula(body, parser_arena);
          if (root != nullptr) {
            const parser::AstNode* shifted = parser::shift_refs(*root, parser_arena, transform);
            ast_changed = (shifted != nullptr && shifted != root);
            if (ast_changed) {
              effective_formula.clear();
              if (had_equals) {
                effective_formula.push_back('=');
              }
              effective_formula.append(parser::format_formula(*shifted));
            }
          }
        }

        if (!ast_changed && !cell_moves && !dropped) {
          continue;  // Cell unaffected by the edit.
        }

        // Keep the rewritten text alongside the key transition. It is
        // applied only after every disappearing / moving old key has been
        // unregistered. Parse failures and an empty body remain opaque and
        // therefore retain their original text exactly.

        const eval::CellNodeId old_node{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)};
        const eval::CellNodeId new_node{static_cast<std::uint16_t>(sheet_idx), new_row, new_col};
        updates.push_back(FormulaUpdate{&sheet, old_node, new_node, dropped, std::move(effective_formula)});
      }
    }
  }

  // Phase 1: no old graph key may survive a coordinate transition.
  for (const FormulaUpdate& update : updates) {
    if (update.dropped || update.old_node != update.new_node) {
      mutator.unregister_formula(update.old_node);
    }
  }

  // Phase 2: apply formula text and create the post-edit graph. Re-parse the
  // stored text because the first-pass AST arenas intentionally expired
  // before phase 1; this keeps memory bounded for large workbooks.
  for (FormulaUpdate& update : updates) {
    if (update.dropped) {
      continue;
    }
    const Cell* current = update.sheet->cell_at(update.old_node.row, update.old_node.col);
    if (current == nullptr || current->formula_text != update.formula) {
      update.sheet->set_cell_formula(update.old_node.row, update.old_node.col, std::move(update.formula));
    }
    const Cell* formula_cell = update.sheet->cell_at(update.old_node.row, update.old_node.col);
    if (formula_cell == nullptr || formula_cell->formula_text.empty()) {
      continue;
    }
    std::string_view body = formula_cell->formula_text;
    if (!body.empty() && body.front() == '=') {
      body.remove_prefix(1);
    }
    parser_arena.reset();
    parser::AstNode* root = parse_indexable_formula(body, parser_arena);
    if (root == nullptr) {
      // The cell has moved, but its formula is opaque. Keep the text as-is
      // and wake the new coordinate without inventing dependency edges.
      mutator.mark_dirty(update.new_node);
      continue;
    }
    mutator.register_formula(update.new_node, *root, workbook);
    mutator.mark_dirty(update.new_node);
  }
}

// Re-register the owners of Ref3D spans after the target sheet's cell store
// has completed its physical move. Their shared inner coordinates intentionally
// remain unchanged, so the ordinary AST rewrite cannot wake them. The owner
// snapshot is taken before rewrite/unregister; map only owners that survive
// the edit, then parse the final formula text against the final workbook
// coordinates so defined-name and nested-Lambda expansion follows the same
// extractor path as ordinary registration.
void reregister_three_d_span_owners_after_row_col_edit(const std::vector<eval::CellNodeId>& owners,
                                                       std::vector<Sheet>& sheets,
                                                       const eval::RecalcEngine::LockedMutator& mutator,
                                                       const Workbook& workbook, std::size_t edited_sheet_index,
                                                       parser::RowColAxis axis, parser::RowColEdit edit,
                                                       std::uint32_t index, std::uint32_t count) {
  Arena parser_arena;
  for (const eval::CellNodeId owner : owners) {
    eval::CellNodeId mapped = owner;
    if (owner.sheet_id == edited_sheet_index) {
      const CellShift shift = shift_cell_coords_for_row_col_edit(axis, edit, index, count, owner.row, owner.col);
      if (!shift.kept) {
        continue;
      }
      mapped.row = shift.new_row;
      mapped.col = shift.new_col;
    }
    if (mapped.sheet_id >= sheets.size()) {
      continue;
    }
    const Cell* formula_cell = sheets[mapped.sheet_id].cell_at(mapped.row, mapped.col);
    if (formula_cell == nullptr || formula_cell->formula_text.empty()) {
      continue;
    }
    parser::AstNode* root = parse_stored_formula(formula_cell->formula_text, parser_arena);
    if (root == nullptr) {
      continue;
    }
    mutator.register_formula(mapped, *root, workbook);
    mutator.mark_dirty(mapped);
  }
}

std::vector<BlockedSpillFootprint> remap_blocked_spill_footprints(const std::vector<BlockedSpillFootprint>& footprints,
                                                                  parser::RowColAxis axis, parser::RowColEdit edit,
                                                                  std::uint32_t index, std::uint32_t count) {
  std::vector<BlockedSpillFootprint> mapped;
  mapped.reserve(footprints.size());
  for (const BlockedSpillFootprint& original : footprints) {
    const CellShift shift =
        shift_cell_coords_for_row_col_edit(axis, edit, index, count, original.anchor_row, original.anchor_col);
    if (!shift.kept) {
      continue;
    }
    BlockedSpillFootprint next = original;
    next.anchor_row = shift.new_row;
    next.anchor_col = shift.new_col;
    mapped.push_back(next);
  }
  return mapped;
}

Expected<void, Error> apply_row_col_edit(Workbook& wb, std::size_t sheet_index, parser::RowColAxis axis,
                                         parser::RowColEdit edit, std::uint32_t origin, std::uint32_t count,
                                         const char* op_name) {
  RETURN_IF_ERROR(check_sheet_index(op_name, sheet_index, wb.sheet_count()));
  if (count == 0U) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op_name) + ": count must be >= 1");
  }
  const std::uint32_t bound = (axis == parser::RowColAxis::kRow) ? Sheet::kMaxRows : Sheet::kMaxCols;
  if (origin >= bound) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op_name) + ": origin out of bounds",
                      "origin=" + std::to_string(origin));
  }
  // Deletion helpers use `origin + count` as their exclusive endpoint.
  // Validate against the remaining grid with subtraction so an unsigned
  // count received through the C/WASM ABI cannot wrap and turn a delete
  // into a corrupting row/column move.
  if (edit == parser::RowColEdit::kDelete && count > bound - origin) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op_name) + ": count exceeds sheet bounds",
                      "origin=" + std::to_string(origin) + " count=" + std::to_string(count));
  }
  return Expected<void, Error>::Ok();
}

// Precondition: caller holds the engine mutex for the lifetime of the
// `LockedMutator&`. The compound edit (per-cell dep-graph re-keys plus
// the eventual `Sheet::insert_rows` / `delete_rows` / `insert_cols` /
// `delete_cols`) runs entirely under that single critical section so
// a concurrent `recalc_parallel` either sees the pre-edit state or
// the fully patched one, never a half-applied edit.
// `rewrite_defined_names` does not touch the engine but is included
// here to preserve the "defined-name table matches dep-graph state"
// invariant for any other reader that consults both under the same
// lock, and because the dep-graph re-index it can trigger has to run
// against the rewritten table.
[[gnu::noinline]] Expected<void, Error> apply_row_col_edit_operation(Workbook& wb, std::vector<Sheet>& sheets,
                                                                     const eval::RecalcEngine::LockedMutator& mutator,
                                                                     std::vector<DefinedName>& defined_names,
                                                                     std::size_t sheet_index, parser::RowColAxis axis,
                                                                     parser::RowColEdit edit, std::uint32_t origin,
                                                                     std::uint32_t count, const char* op_name) {
  RETURN_IF_ERROR(apply_row_col_edit(wb, sheet_index, axis, edit, origin, count, op_name));
  const std::string target_sheet_name = sheets[sheet_index].name();
  // `apply_row_col_edit` above rejected an out-of-range `sheet_index`, and
  // `Workbook::kMaxSheets` bounds `sheet_count()`, so the narrowing to the
  // dep graph's 16-bit sheet id keeps the index intact.
  const std::vector<eval::CellNodeId> three_d_owners =
      mutator.three_d_span_owners_covering_sheet(static_cast<std::uint16_t>(sheet_index));
  // Defined names are rewritten before cell formulas, and a change to any of
  // them forces a full re-index afterwards. A formula that reaches a shifted
  // range only through a name — `=MyRef*1` — is textually unchanged by the
  // shift, because a `NameRef` node is the identity case for the transform.
  // The per-formula rewriter only re-registers formulas whose text changed,
  // so without the re-index that formula keeps dep-graph edges pointing at
  // the range `MyRef` used to cover. `set_defined_name_scoped` and
  // `remove_sheet` already take this fallback for the same reason.
  const parser::RowColShiftTransform name_transform(target_sheet_name, axis, edit, origin, count);
  const parser::RowColShiftTransform local_transform(target_sheet_name, axis, edit, origin, count,
                                                     /*local_means_target=*/true);
  std::vector<const parser::RefTransform*> per_sheet_transforms(sheets.size(), &name_transform);
  per_sheet_transforms[sheet_index] = &local_transform;
  bool names_changed = rewrite_defined_names(defined_names, per_sheet_transforms, name_transform);
  // Snapshot the pending-footprint records before formula text rewrites.
  // `Sheet::set_cell_formula` quite correctly clears a user-overwritten
  // blocked anchor; a structural rewrite of that same formula is not a user
  // overwrite, so restore the mapped record after the physical move below.
  // Keep the other sheets too: a name rewrite can force the full graph reset,
  // whose spill invalidation must not silently discard unrelated pending
  // producers.
  std::vector<std::vector<BlockedSpillFootprint>> blocked_before_all;
  blocked_before_all.reserve(sheets.size());
  for (const Sheet& sheet : sheets) {
    blocked_before_all.push_back(sheet.blocked_spill_footprints());
  }
  Sheet& target = sheets[sheet_index];
  const std::vector<BlockedSpillFootprint> blocked_before = blocked_before_all[sheet_index];
  // The edit clears committed spills; snapshot their anchors so those formulas are re-dirtied.
  const std::vector<SpillFootprint> committed_before = target.committed_spill_footprints();
  // AutoFilters on the edited sheet, serialized before the move, to detect a
  // change that alters which hidden rows count as filtered.
  const bool had_sheet_filter = target.has_auto_filter();
  std::vector<std::string> filters_before;
  filters_before.push_back(io::auto_filter_xml(target.auto_filter()));
  for (const TableMetadata& table : wb.tables()) {
    if (table.sheet_index == sheet_index) {
      filters_before.push_back(io::auto_filter_xml(table.auto_filter_xml.get()));
    }
  }
  rewrite_formulas_for_row_col_edit(sheets, mutator, wb, target_sheet_name, axis, edit, origin, count);
  // Metadata formulas follow the same per-sheet policy as cell formulas: a
  // qualified reference to the edited sheet shifts no matter which sheet owns
  // the rule, while an unqualified one is in scope only on the edited sheet.
  // The text is rewritten against pre-edit coordinates, like the cells above
  // and unlike the sqref/anchor rectangles, which `Sheet::insert_rows` and
  // friends move afterwards.
  std::vector<std::uint32_t> ignored_cache_ids;
  rewrite_sheet_metadata_formulas(sheets, per_sheet_transforms, wb.mutable_tables(), wb.mutable_pivot_caches(), {}, {},
                                  {}, ignored_cache_ids, name_transform);
  if (axis == parser::RowColAxis::kRow) {
    if (edit == parser::RowColEdit::kInsert) {
      target.insert_rows(origin, count);
    } else {
      target.delete_rows(origin, count);
    }
  } else if (edit == parser::RowColEdit::kInsert) {
    target.insert_cols(origin, count);
  } else {
    target.delete_cols(origin, count);
  }
  shift_retained_extension_ranges(target, origin, count, edit == parser::RowColEdit::kDelete,
                                  axis == parser::RowColAxis::kRow);
  shift_drawing_anchors(wb, sheet_index, origin, count, edit == parser::RowColEdit::kDelete,
                        axis == parser::RowColAxis::kRow);
  std::vector<std::string> filters_after;
  filters_after.push_back(io::auto_filter_xml(target.auto_filter()));
  for (TableMetadata& table : wb.mutable_tables()) {
    if (table.sheet_index != sheet_index) {
      continue;
    }
    if (AutoFilter* filter = table.auto_filter_xml.get();
        filter != nullptr && !shift_auto_filter(*filter, origin, count, edit == parser::RowColEdit::kDelete,
                                                axis == parser::RowColAxis::kRow, /*header_delete_removes=*/false)) {
      table.auto_filter_xml.reset();
    }
    filters_after.push_back(io::auto_filter_xml(table.auto_filter_xml.get()));
  }
  // A sheet AutoFilter the edit removed takes its `_FilterDatabase` name
  // with it; the full re-index below covers formulas naming it.
  if (had_sheet_filter && !target.has_auto_filter() && erase_filter_database_name(defined_names, sheet_index)) {
    names_changed = true;
  }
  if (filters_after != filters_before) {
    mark_row_visibility_dependents_dirty_locked(sheets, mutator);
  }
  // Re-index only after the physical move. Formula text is rewritten while
  // cells still occupy their pre-edit coordinates, so rebuilding the graph
  // before the move would register moved owners (including Ref3D owners) at
  // stale coordinates and leave duplicate registry entries behind.
  if (names_changed) {
    reindex_all_formulas(sheets, mutator, wb);
    for (std::size_t preserved_sheet = 0; preserved_sheet < sheets.size(); ++preserved_sheet) {
      if (preserved_sheet == sheet_index) {
        continue;
      }
      sheets[preserved_sheet].restore_blocked_spill_footprints(std::move(blocked_before_all[preserved_sheet]));
    }
  }
  reregister_three_d_span_owners_after_row_col_edit(three_d_owners, sheets, mutator, wb, sheet_index, axis, edit,
                                                    origin, count);
  // Map each snapshotted anchor through the edit; deleted anchors and non-formula cells are skipped.
  for (const SpillFootprint& footprint : committed_before) {
    const CellShift shift =
        shift_cell_coords_for_row_col_edit(axis, edit, origin, count, footprint.anchor_row, footprint.anchor_col);
    if (!shift.kept) {
      continue;
    }
    const Cell* cell = target.cell_at(shift.new_row, shift.new_col);
    if (cell == nullptr || cell->formula_text.empty()) {
      continue;
    }
    mutator.mark_dirty(make_node(sheet_index, shift.new_row, shift.new_col));
  }
  std::vector<BlockedSpillFootprint> blocked_mapped =
      remap_blocked_spill_footprints(blocked_before, axis, edit, origin, count);
  // Do not resurrect a record whose formula anchor was deleted or moved past
  // the grid. The cell store has already completed the physical move, so a
  // formula check is stable while the engine mutex is held.
  blocked_mapped.erase(std::remove_if(blocked_mapped.begin(), blocked_mapped.end(),
                                      [&](const auto& footprint) {
                                        const Cell* cell = target.cell_at(footprint.anchor_row, footprint.anchor_col);
                                        return cell == nullptr || cell->formula_text.empty();
                                      }),
                       blocked_mapped.end());
  target.restore_blocked_spill_footprints(std::move(blocked_mapped));
  // Sheet-local row/column edits remap the pending spill-footprint reverse
  // index alongside formula cells.  Dirty the surviving anchors at their new
  // coordinates; the next recalc retries each formula against the moved
  // blocker/footprint. Anchors deleted or shifted past the grid were dropped
  // by the Sheet remap and therefore do not appear here.
  for (const CellAddress anchor : target.blocked_spill_anchors()) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<void, Error> Workbook::insert_rows(std::size_t sheet_index, std::uint32_t row, std::uint32_t count) {
  // The whole edit — including the per-cell dep-graph re-keys done by
  // `rewrite_formulas_for_row_col_edit` and the eventual
  // `Sheet::insert_rows` — must run under a single hold of the engine
  // mutex. A concurrent `recalc_parallel` holds the same mutex for its
  // full pass, so the compound mutation either runs entirely before
  // or entirely after a recalc, never half-applied alongside it. The
  // `LockedMutator` facade routes through the engine's `*_locked` API
  // and assumes this lock is already held.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index, parser::RowColAxis::kRow,
                                      parser::RowColEdit::kInsert, row, count, "insert_rows");
}

Expected<void, Error> Workbook::delete_rows(std::size_t sheet_index, std::uint32_t row, std::uint32_t count) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index, parser::RowColAxis::kRow,
                                      parser::RowColEdit::kDelete, row, count, "delete_rows");
}

Expected<void, Error> Workbook::insert_cols(std::size_t sheet_index, std::uint32_t col, std::uint32_t count) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index, parser::RowColAxis::kCol,
                                      parser::RowColEdit::kInsert, col, count, "insert_cols");
}

Expected<void, Error> Workbook::delete_cols(std::size_t sheet_index, std::uint32_t col, std::uint32_t count) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index, parser::RowColAxis::kCol,
                                      parser::RowColEdit::kDelete, col, count, "delete_cols");
}

}  // namespace formulon
