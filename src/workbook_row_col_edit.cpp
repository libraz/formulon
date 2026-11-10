//
// Row and column insert/delete for `Workbook`: the formula rewrite and
// dependency-graph re-keying around `Sheet`'s cell-store shift, plus the
// spill, AutoFilter and defined-name bookkeeping that follows it.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "chart/chart_ref.h"
#include "defined_name.h"
#include "drawing/drawing_edit.h"
#include "drawing/slicer_edit.h"
#include "eval/dep_graph.h"
#include "eval/recalc_engine.h"
#include "io/auto_filter_xml.h"
#include "io/worksheet_child_refs.h"
#include "new_workbook_defaults.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/ref_transforms.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "sheet_name.h"
#include "sheet_passthrough.h"
#include "styles.h"
#include "table.h"
#include "utils/a1_ref.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "utils/status_macros.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"
#include "workbook_ref_rewrite.h"
#include "workbook_sheet_mutation.h"
#include "workbook_table_loss.h"

namespace formulon {

namespace {

// Computes the post-edit coordinates of a cell on the edited sheet.
// `kept == false` means the cell is dropped by the upcoming cell-store move
// (deletion fully covers the cell) or pushed past the sheet bound by an
// insert. A cell outside a band edit's band stays where it is.
struct ShiftedCell {
  bool kept = true;
  std::uint32_t new_row = 0;
  std::uint32_t new_col = 0;
};

ShiftedCell shift_cell_coords_for_row_col_edit(const StructuralEdit& edit, std::uint32_t row, std::uint32_t col) {
  ShiftedCell r;
  r.new_row = row;
  r.new_col = col;
  std::uint32_t& target = edit.row_axis ? r.new_row : r.new_col;
  const std::uint32_t across = edit.row_axis ? col : row;
  if (edit.band.has_value() && (across < edit.band->first || across > edit.band->last)) {
    return r;
  }
  const std::uint32_t coord = target;
  const std::uint32_t bound = a1::axis_extent(edit.row_axis);
  const std::uint32_t index = edit.index;
  const std::uint32_t count = edit.count;

  if (!edit.is_delete) {
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

// The reference transform of `edit` on `target_sheet`, confined to its band
// when it has one.
parser::RowColShiftTransform shift_transform(std::string_view target_sheet, const StructuralEdit& edit,
                                             bool local_means_target, const sheet_names::NameList* sheet_list) {
  const parser::RowColAxis axis = edit.row_axis ? parser::RowColAxis::kRow : parser::RowColAxis::kCol;
  const parser::RowColEdit kind = edit.is_delete ? parser::RowColEdit::kDelete : parser::RowColEdit::kInsert;
  if (edit.band.has_value()) {
    return parser::RowColShiftTransform(target_sheet, axis, kind, edit.index, edit.count, *edit.band,
                                        local_means_target, sheet_list);
  }
  return parser::RowColShiftTransform(target_sheet, axis, kind, edit.index, edit.count, local_means_target, sheet_list);
}

// @size-budget: 11 KB
// Whole row/column edits' table, data-table and pivot checks; Insert/Delete Cells band edits are budgeted at
// `Workbook::insert_cells`.
std::vector<TableLoss> table_losses(const std::vector<TableMetadata>& tables, std::size_t sheet_index,
                                    const StructuralEdit& edit) {
  std::vector<TableLoss> losses;
  if (!edit.is_delete) {
    return losses;
  }
  const std::uint64_t end = static_cast<std::uint64_t>(edit.index) + edit.count;  // exclusive
  for (const TableMetadata& table : tables) {
    const std::optional<MergeRange> ref = detail::table_rect_on_sheet(table, sheet_index);
    // A band edit reaches only a table its band covers.
    if (!ref.has_value() || (edit.band.has_value() && band_fit(*ref, edit) != BandFit::kInside)) {
      continue;
    }
    const std::uint32_t first = edit.row_axis ? ref->first_row : ref->first_col;
    const std::uint32_t last = edit.row_axis ? ref->last_row : ref->last_col;
    if (end <= first || edit.index > last) {
      continue;
    }
    TableLoss loss{table.name, table.display_name, edit.index <= first && end > last, {}, {}};
    if (!loss.whole && !edit.row_axis) {
      const std::uint64_t stop = std::min<std::uint64_t>(end, last + 1ULL);
      for (std::uint64_t col = std::max(edit.index, first); col < stop; ++col) {
        const auto offset = static_cast<std::size_t>(col - first);
        if (offset < table.columns.size()) {
          loss.columns.push_back(table.columns[offset].name);
        }
      }
      if (!loss.columns.empty()) {
        for (const TableColumn& column : table.columns) {
          loss.all_columns.push_back(column.name);
        }
      }
    }
    if (loss.whole || !loss.columns.empty()) {
      losses.push_back(std::move(loss));
    }
  }
  return losses;
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
                                       const Workbook& workbook, std::string_view target_sheet,
                                       const StructuralEdit& edit, const std::vector<TableLoss>& table_losses,
                                       const sheet_names::NameList* table_list) {
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
  const sheet_names::NameList sheet_list = detail::sheet_name_list(sheets);
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    Sheet& sheet = sheets[sheet_idx];
    const bool local_means_target = sheet_names::equal(sheet.name(), target_sheet);
    const parser::RowColShiftTransform shift = shift_transform(target_sheet, edit, local_means_target, &sheet_list);
    const TableLossTransform transform(shift, table_losses, table_list);

    for (const auto& [row, cells] : sheet.rows()) {
      for (const auto& [col, cell] : cells) {
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
          const ShiftedCell sr = shift_cell_coords_for_row_col_edit(edit, row, static_cast<std::uint32_t>(col));
          if (!sr.kept) {
            dropped = true;
          } else {
            new_row = sr.new_row;
            new_col = sr.new_col;
          }
        }
        const bool cell_moves = (new_row != row) || (new_col != static_cast<std::uint32_t>(col));

        const std::string_view body = strip_formula_prefix(cell.formula_text);
        const bool had_equals = body.size() != cell.formula_text.size();

        bool ast_changed = false;
        std::string rewritten;
        parser_arena.reset();
        if (!body.empty()) {
          parser::AstNode* root = parse_indexable_formula(body, parser_arena);
          if (root != nullptr) {
            const parser::AstNode* shifted = parser::shift_refs(*root, parser_arena, transform);
            ast_changed = (shifted != nullptr && shifted != root);
            if (ast_changed) {
              if (had_equals) {
                rewritten.push_back('=');
              }
              rewritten.append(parser::format_formula(*shifted));
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

        const eval::CellNodeId old_node = make_node(sheet_idx, row, static_cast<std::uint32_t>(col));
        const eval::CellNodeId new_node = make_node(sheet_idx, new_row, new_col);
        if (!ast_changed) {
          rewritten = cell.formula_text;
        }
        updates.push_back(FormulaUpdate{&sheet, old_node, new_node, dropped, std::move(rewritten)});
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
    const std::string_view body = strip_formula_prefix(formula_cell->formula_text);
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
                                                       const StructuralEdit& edit) {
  Arena parser_arena;
  for (const eval::CellNodeId owner : owners) {
    eval::CellNodeId mapped = owner;
    if (owner.sheet_id == edited_sheet_index) {
      const ShiftedCell shift = shift_cell_coords_for_row_col_edit(edit, owner.row, owner.col);
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
                                                                  const StructuralEdit& edit) {
  std::vector<BlockedSpillFootprint> mapped;
  mapped.reserve(footprints.size());
  for (const BlockedSpillFootprint& original : footprints) {
    const ShiftedCell shift = shift_cell_coords_for_row_col_edit(edit, original.anchor_row, original.anchor_col);
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

// Gives a table the column entries its rewritten `ref` spans: a column delete
// removes the deleted columns' entries, and an insert inside the table adds
// entries named from their left neighbour (`detail::add_table_columns_locked`).
// `before` is the table's extent before the edit.
void update_table_columns(TableMetadata& table, const MergeRange& before, const MergeRange& after, Sheet& sheet,
                          std::size_t sheet_index, const eval::RecalcEngine::LockedMutator& mutator,
                          const StructuralEdit& edit, std::string_view column_word) {
  const std::uint64_t edit_end = static_cast<std::uint64_t>(edit.index) + edit.count;  // exclusive
  if (edit.is_delete) {
    // The deleted columns inside the table, as offsets from its first column.
    const std::uint64_t lo = std::max<std::uint64_t>(edit.index, before.first_col);
    const std::uint64_t hi = std::min<std::uint64_t>(edit_end, before.last_col + 1ULL);
    const std::size_t size = table.columns.size();
    if (lo < hi && lo - before.first_col < size) {
      table.columns.erase(
          table.columns.begin() + static_cast<std::ptrdiff_t>(lo - before.first_col),
          table.columns.begin() + static_cast<std::ptrdiff_t>(std::min<std::uint64_t>(hi - before.first_col, size)));
    }
    return;
  }
  if (edit.index <= before.first_col || edit.index > before.last_col) {
    return;
  }
  detail::add_table_columns_locked(table, edit.index - before.first_col, edit.count, after, sheet, sheet_index, mutator,
                                   column_word);
}

// Brings the tables on the edited sheet in line with the `ref` their rewrite
// gave them: column entries follow a column edit inside the table, the sort
// state moves the way a worksheet one does, and a table whose whole extent was
// deleted goes. A row insert that grew a table fills its calculated columns
// into the inserted data rows. `refs_before` holds each table's `ref` before
// the edit, by index. Returns the names of every table whose extent changed,
// for the re-index of the structured references that resolve through it.
std::vector<std::string> update_tables_for_row_col_edit(const Workbook& wb, std::vector<TableMetadata>& tables,
                                                        Sheet& sheet, std::size_t sheet_index,
                                                        const std::vector<std::string>& refs_before,
                                                        const eval::RecalcEngine::LockedMutator& mutator,
                                                        const StructuralEdit& edit, std::string_view column_word) {
  struct Fill {
    std::size_t table;
    std::uint32_t first_row;
    std::uint32_t last_row;
  };
  std::vector<Fill> fills;
  std::vector<std::string> changed;
  std::vector<TableMetadata> retained;
  retained.reserve(tables.size());
  for (std::size_t i = 0; i < tables.size(); ++i) {
    TableMetadata& table = tables[i];
    if (table.sheet_index != sheet_index || table.ref == refs_before[i]) {
      retained.push_back(std::move(table));
      continue;
    }
    changed.push_back(table.name);
    changed.push_back(table.display_name);
    const std::optional<MergeRange> after = parse_a1_rectangle(table.ref);
    if (!after.has_value()) {
      continue;
    }
    const std::optional<MergeRange> before = parse_a1_rectangle(refs_before[i]);
    if (before.has_value() && !edit.row_axis) {
      update_table_columns(table, *before, *after, sheet, sheet_index, mutator, edit, column_word);
    }
    if (before.has_value() && edit.row_axis && !edit.is_delete && before->first_row == after->first_row &&
        after->last_row > before->last_row) {
      const std::uint64_t first = std::max<std::uint64_t>(edit.index, after->first_row + (table.header_row ? 1U : 0U));
      const std::uint64_t last =
          std::min<std::uint64_t>(edit.index + static_cast<std::uint64_t>(edit.count) - 1U,
                                  static_cast<std::uint64_t>(after->last_row) - (table.totals_row ? 1U : 0U));
      // Rows inserted at the first data row moved the formulas' anchor cell down; the anchor stays put.
      if (table.header_row && edit.index == after->first_row + 1U) {
        detail::shift_calculated_column_anchor(table, -static_cast<std::int32_t>(edit.count));
      }
      if (first <= last) {
        fills.push_back(Fill{retained.size(), static_cast<std::uint32_t>(first), static_cast<std::uint32_t>(last)});
      }
    }
    if (!table.sort_state_xml.empty()) {
      WorksheetRawExtensions sort_state{WorksheetRawChild{
          static_cast<std::uint32_t>(worksheet_child::slot_of("sortState")), std::move(table.sort_state_xml)}};
      io::remap_worksheet_children(sort_state, edit);
      table.sort_state_xml = sort_state.empty() ? std::string() : std::move(sort_state.front().xml);
    }
    retained.push_back(std::move(table));
  }
  tables = std::move(retained);
  // Filled once every table is back in place, as structured references resolve through them.
  for (const Fill& fill : fills) {
    detail::fill_calculated_columns_locked(wb, sheet, mutator, tables[fill.table], fill.first_row, fill.last_row, 0U,
                                           tables[fill.table].columns.size() - 1U);
  }
  return changed;
}

// Moves each shared-formula group of `sheet` through `transform` the way a
// table's `ref` moves; a group the edit deletes or leaves non-rectangular goes.
void update_shared_formula_groups(Sheet& sheet, const parser::RefTransform& transform) {
  std::vector<SharedFormulaGroup>& groups = sheet.mutable_shared_formula_groups();
  std::vector<SharedFormulaGroup> kept;
  kept.reserve(groups.size());
  for (const SharedFormulaGroup& group : groups) {
    std::string ref;
    if (!append_a1_rectangle(ref, group.ref)) {
      continue;
    }
    const FormulaRewriteResult result = rewrite_formula(ref, transform);
    const std::optional<MergeRange> moved = result.error ? std::nullopt : parse_a1_rectangle(result.text);
    if (moved.has_value()) {
      kept.push_back(SharedFormulaGroup{*moved});
    }
  }
  groups = std::move(kept);
}

// Whether cell format `xf_index` draws something: a fill or a border. Fill
// records 0 and 1 are the reserved none / gray125 entries no cell shows.
bool xf_is_visible(const StylesTable& styles, std::uint32_t xf_index) {
  if (xf_index == 0U || xf_index >= styles.cell_xfs.size()) {
    return false;
  }
  const CellXf& xf = styles.cell_xfs[xf_index];
  if (xf.fill_index > 1U && xf.fill_index < styles.fills.size()) {
    const FillRecord& fill = styles.fills[xf.fill_index];
    if (fill.has_gradient || fill.pattern != 0U) {
      return true;
    }
  }
  if (xf.border_index < styles.borders.size()) {
    const BorderRecord& border = styles.borders[xf.border_index];
    const bool diagonal = (border.diagonal_up || border.diagonal_down) && border.diagonal.style != 0U;
    if (border.left.style != 0U || border.right.style != 0U || border.top.style != 0U || border.bottom.style != 0U ||
        diagonal) {
      return true;
    }
  }
  return false;
}

// What an insert would push past the grid's edge that blocks it, or null: a
// value or formula, a note, a fill or border, or a merge, as Excel refuses to
// shift non-empty cells off the sheet. A cell format that draws nothing (a
// font or number format), row and column formats, hyperlinks, validations and
// conditional formats are dropped there instead.
const char* insert_pushes_content_off(const Sheet& sheet, const StylesTable& styles, const StructuralEdit& edit) {
  if (edit.is_delete) {
    return nullptr;
  }
  const std::uint32_t bound = a1::axis_extent(edit.row_axis);
  const std::uint32_t first_lost = std::max(edit.index, edit.count >= bound ? 0U : bound - edit.count);
  const std::uint32_t across_first = edit.band.has_value() ? edit.band->first : 0U;
  const std::uint32_t across_last = edit.band.has_value() ? edit.band->last : a1::axis_extent(!edit.row_axis) - 1U;
  // The strip that falls off the grid, as an inclusive rectangle.
  const MergeRange lost = edit.row_axis ? MergeRange{first_lost, across_first, bound - 1U, across_last}
                                        : MergeRange{across_first, first_lost, across_last, bound - 1U};
  const auto in_lost = [&lost](std::uint32_t row, std::uint32_t col) {
    return lost.first_row <= row && row <= lost.last_row && lost.first_col <= col && col <= lost.last_col;
  };
  bool formatted = false;
  for (const auto& [row, cells] : sheet.rows()) {
    if (row < lost.first_row || row > lost.last_row) {
      continue;
    }
    for (const RowCells::Entry& entry : cells.entries_in(lost.first_col, lost.last_col + 1ULL)) {
      if (!entry.cell.formula_text.empty() || !entry.cell.cached_value.is_blank()) {
        return "a value or formula";
      }
      formatted = formatted || xf_is_visible(styles, entry.cell.xf_index);
    }
  }
  for (const CellComment& comment : sheet.comments()) {
    if (in_lost(comment.row, comment.col)) {
      return "a note";
    }
  }
  for (const ThreadedComment& thread : sheet.threaded_comments()) {
    if (in_lost(thread.row, thread.col)) {
      return "a note";
    }
  }
  if (formatted) {
    return "a fill or border";
  }
  for (const MergeRange& merge : sheet.merges()) {
    if (merge.first_row <= lost.last_row && lost.first_row <= merge.last_row && merge.first_col <= lost.last_col &&
        lost.first_col <= merge.last_col) {
      return "a merge";
    }
  }
  return nullptr;
}

// The structure-intersection preflight every row/column edit passes before
// anything moves. A data table, a legacy array block and a PivotTable each move
// or go whole, so an edit cutting into one is refused, as a write into it is.
// A table grows and shrinks with an edit inside it instead. An insert that
// would push any of them, or a value or formula, past the grid edge is
// refused, as Excel refuses to shift non-empty cells off the sheet. A band
// edit is also refused when its moving region meets one of these, or a
// table, across a band edge, and when it lies inside a PivotTable. A band
// insert is refused when it would move a PivotTable at all, and when it cuts
// into a table without spanning exactly that table's columns (rows).
Expected<void, Error> check_structure_edit(const Workbook& wb, std::size_t sheet_index, const StructuralEdit& edit) {
  const auto refuse = [&](StructureEditCut cut, const char* what) {
    std::string message(cut == StructureEditCut::kPushesOff ? "Cannot shift " : "Cannot change part of ");
    message.append(what);
    if (cut == StructureEditCut::kPushesOff) {
      message.append(" off the sheet");
    }
    return make_error(FormulonErrorCode::kPreconditionFailed, std::move(message),
                      "origin=" + std::to_string(edit.index) + " count=" + std::to_string(edit.count));
  };
  const auto refused = [](StructureEditCut cut) {
    return cut == StructureEditCut::kSplits || cut == StructureEditCut::kPushesOff;
  };
  // A whole row/column edit meets every structure along its axis alone.
  const auto fit_of = [&edit](const MergeRange& rect) {
    return edit.band.has_value() ? band_fit(rect, edit) : BandFit::kInside;
  };
  const Sheet& sheet = wb.sheet(sheet_index);
  for (const DataTable& table : sheet.data_tables()) {
    const BandFit fit = fit_of(data_table_block(table));
    if (fit == BandFit::kStraddles) {
      return refuse(StructureEditCut::kSplits, "a data table");
    }
    if (const StructureEditCut cut = data_table_structure_cut(table, edit); fit == BandFit::kInside && refused(cut)) {
      return refuse(cut, "a data table");
    }
  }
  for (const CellAddress address : sheet.all_formula_cells()) {
    const Cell* cell = sheet.cell_at(address.row, address.col);
    if (cell == nullptr || cell->array_block_rows == 0U) {
      continue;
    }
    const MergeRange block{address.row, address.col, address.row + cell->array_block_rows - 1U,
                           address.col + cell->array_block_cols - 1U};
    const BandFit fit = fit_of(block);
    if (fit == BandFit::kStraddles) {
      return refuse(StructureEditCut::kSplits, "an array");
    }
    const StructureEditCut cut = edit.row_axis ? structure_edit_cut(block.first_row, block.last_row, edit)
                                               : structure_edit_cut(block.first_col, block.last_col, edit);
    if (fit == BandFit::kInside && refused(cut)) {
      return refuse(cut, "an array");
    }
  }
  for (const std::unique_ptr<pivot::PivotTable>& pivot : sheet.pivot_tables()) {
    if (pivot == nullptr) {
      continue;
    }
    const MergeRange location = pivot_location(*pivot);
    const BandFit fit = fit_of(location);
    if (fit == BandFit::kStraddles) {
      return refuse(StructureEditCut::kSplits, "a PivotTable");
    }
    // Insert Cells never shifts a PivotTable along.
    if (edit.band.has_value() && !edit.is_delete && fit == BandFit::kInside) {
      return refuse(StructureEditCut::kSplits, "a PivotTable");
    }
    // Cells selected inside a report are the report's own to change.
    if (edit.band.has_value()) {
      const std::uint32_t last = edit.index + edit.count - 1U;
      const MergeRange selected = edit.row_axis ? MergeRange{edit.index, edit.band->first, last, edit.band->last}
                                                : MergeRange{edit.band->first, edit.index, edit.band->last, last};
      const bool within = location.first_row <= selected.first_row && selected.last_row <= location.last_row &&
                          location.first_col <= selected.first_col && selected.last_col <= location.last_col;
      if (within && !(edit.is_delete && selected == location)) {
        return refuse(StructureEditCut::kSplits, "a PivotTable");
      }
    }
    if (const StructureEditCut cut = pivot_structure_cut(*pivot, edit); fit == BandFit::kInside && refused(cut)) {
      return refuse(cut, "a PivotTable");
    }
  }
  for (const TableMetadata& table : wb.tables()) {
    const std::optional<MergeRange> ref = detail::table_rect_on_sheet(table, sheet_index);
    if (!ref.has_value()) {
      continue;
    }
    const BandFit fit = fit_of(*ref);
    if (fit == BandFit::kStraddles) {
      return refuse(StructureEditCut::kSplits, "a table");
    }
    const StructureEditCut cut = edit.row_axis ? structure_edit_cut(ref->first_row, ref->last_row, edit)
                                               : structure_edit_cut(ref->first_col, ref->last_col, edit);
    if (fit == BandFit::kInside && cut == StructureEditCut::kPushesOff) {
      return refuse(cut, "a table");
    }
    // Insert Cells growing a table spans exactly its columns (rows), totals row included.
    if (edit.band.has_value() && !edit.is_delete && fit == BandFit::kInside && cut == StructureEditCut::kSplits &&
        (edit.band->first != (edit.row_axis ? ref->first_col : ref->first_row) ||
         edit.band->last != (edit.row_axis ? ref->last_col : ref->last_row))) {
      return refuse(cut, "a table");
    }
  }
  if (const char* blocker = insert_pushes_content_off(sheet, wb.styles(), edit); blocker != nullptr) {
    return refuse(StructureEditCut::kPushesOff, blocker);
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> apply_row_col_edit(Workbook& wb, std::size_t sheet_index, const StructuralEdit& edit) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, wb.sheet_count()));
  if (edit.count == 0U) {
    return make_error(FormulonErrorCode::kInvalidArgument, "count must be >= 1");
  }
  const std::uint32_t bound = a1::axis_extent(edit.row_axis);
  if (edit.index >= bound) {
    return make_error(FormulonErrorCode::kInvalidArgument, "first is out of bounds",
                      "origin=" + std::to_string(edit.index));
  }
  // Deletion helpers use `origin + count` as their exclusive endpoint.
  // Validate against the remaining grid with subtraction so an unsigned
  // count received through the C/WASM ABI cannot wrap and turn a delete
  // into a corrupting row/column move.
  if (edit.is_delete && edit.count > bound - edit.index) {
    return make_error(FormulonErrorCode::kInvalidArgument, "count exceeds the sheet bounds",
                      "origin=" + std::to_string(edit.index) + " count=" + std::to_string(edit.count));
  }
  return check_structure_edit(wb, sheet_index, edit);
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
//
// `edit` carries a band for Insert / Delete Cells: then only the band's cells
// move, references follow the band rule of `RowColShiftTransform`, and the row
// and column geometry (drawing anchors included) stays.
[[gnu::noinline]] Expected<void, Error> apply_row_col_edit_operation(Workbook& wb, std::vector<Sheet>& sheets,
                                                                     const eval::RecalcEngine::LockedMutator& mutator,
                                                                     std::vector<DefinedName>& defined_names,
                                                                     std::size_t sheet_index,
                                                                     const StructuralEdit& edit) {
  RETURN_IF_ERROR(apply_row_col_edit(wb, sheet_index, edit));
  const std::uint32_t origin = edit.index;
  const std::uint32_t count = edit.count;
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
  const sheet_names::NameList sheet_list = detail::sheet_name_list(sheets);
  const parser::RowColShiftTransform name_transform =
      shift_transform(target_sheet_name, edit, /*local_means_target=*/false, &sheet_list);
  const parser::RowColShiftTransform local_transform =
      shift_transform(target_sheet_name, edit, /*local_means_target=*/true, &sheet_list);
  const sheet_names::NameList table_list = sheet_names::table_name_list(wb.tables());
  const std::vector<TableLoss> losses = table_losses(wb.tables(), sheet_index, edit);
  const TableLossTransform name_rewrite(name_transform, losses, &table_list);
  const TableLossTransform local_rewrite(local_transform, losses, &table_list);
  std::vector<const parser::RefTransform*> per_sheet_transforms(sheets.size(), &name_rewrite);
  per_sheet_transforms[sheet_index] = &local_rewrite;
  bool names_changed = rewrite_defined_names(defined_names, per_sheet_transforms, name_rewrite);
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
  // Every table's `ref` before the rewrite, by index, to tell which ones the edit moved or resized.
  std::vector<std::string> table_refs_before;
  table_refs_before.reserve(wb.tables().size());
  for (const TableMetadata& table : wb.tables()) {
    table_refs_before.push_back(table.ref);
  }
  rewrite_formulas_for_row_col_edit(sheets, mutator, wb, target_sheet_name, edit, losses, &table_list);
  // Metadata formulas follow the same per-sheet policy as cell formulas: a
  // qualified reference to the edited sheet shifts no matter which sheet owns
  // the rule, while an unqualified one is in scope only on the edited sheet.
  // The text is rewritten against pre-edit coordinates, like the cells above
  // and unlike the sqref/anchor rectangles, which `Sheet::insert_rows` and
  // friends move afterwards.
  std::vector<std::uint32_t> ignored_cache_ids;
  rewrite_sheet_metadata_formulas(sheets, per_sheet_transforms, wb.mutable_tables(), wb.mutable_pivot_caches(), {}, {},
                                  {}, ignored_cache_ids, name_rewrite);
  // Chart refs are always sheet-qualified, so the unowned policy is the right one.
  RETURN_IF_ERROR(chart_ref::rewrite_all(wb, name_rewrite));
  if (edit.band.has_value()) {
    target.edit_cells_in_band(edit);
    shift_retained_extension_ranges(target, edit);
  } else {
    if (edit.row_axis) {
      if (edit.is_delete) {
        target.delete_rows(origin, count);
      } else {
        target.insert_rows(origin, count);
      }
    } else if (edit.is_delete) {
      target.delete_cols(origin, count);
    } else {
      target.insert_cols(origin, count);
    }
    shift_retained_extension_ranges(target, edit);
    const std::vector<std::string> framed_slicers_before = framed_slicers(wb, sheet_index);
    shift_drawing_anchors(wb, sheet_index, origin, count, edit.is_delete, edit.row_axis);
    // A slicer whose frame the edit deleted goes with its cache; their `Slicer_*` names leave here, under the lock.
    if (erase_slicer_cache_names(defined_names, remove_unframed_slicers(wb, sheet_index, framed_slicers_before))) {
      names_changed = true;
    }
  }
  // A table the edit deleted whole takes its slicers and their caches with it.
  for (std::size_t i = 0; i < wb.tables().size(); ++i) {
    const TableMetadata& table = wb.tables()[i];
    if (table.sheet_index == sheet_index && table.ref != table_refs_before[i] &&
        !parse_a1_rectangle(table.ref).has_value() &&
        erase_slicer_cache_names(defined_names, remove_table_slicers_locked(wb, table.id))) {
      names_changed = true;
    }
  }
  const std::vector<std::string> changed_tables =
      update_tables_for_row_col_edit(wb, wb.mutable_tables(), target, sheet_index, table_refs_before, mutator, edit,
                                     new_workbook_defaults(wb.excel_profile().locale).column_word);
  update_shared_formula_groups(target, local_rewrite);
  sync_table_spill_barriers_locked(sheets, mutator, wb.tables());
  std::vector<std::string> filters_after;
  filters_after.push_back(io::auto_filter_xml(target.auto_filter()));
  for (TableMetadata& table : wb.mutable_tables()) {
    if (table.sheet_index != sheet_index) {
      continue;
    }
    if (AutoFilter* filter = table.auto_filter_xml.get();
        filter != nullptr && !shift_auto_filter(*filter, edit, /*header_delete_removes=*/false)) {
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
    mark_row_visibility_dependents_dirty_locked(sheets, mutator, defined_names);
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
  } else if (!changed_tables.empty()) {
    // A structured reference resolves through its table's `ref`, so a formula
    // naming a moved, resized or removed table keeps its text but not its edges.
    reindex_formulas_referencing_table(sheets, mutator, wb, changed_tables);
  }
  reregister_three_d_span_owners_after_row_col_edit(three_d_owners, sheets, mutator, wb, sheet_index, edit);
  // Compact range owners over the moved region read cells that moved or vanished without any cell write.
  {
    eval::CellRangeDependency moved;
    moved.sheet_id = static_cast<std::uint16_t>(sheet_index);
    const std::uint32_t across_first = edit.band.has_value() ? edit.band->first : 0U;
    moved.row_first = edit.row_axis ? origin : across_first;
    moved.row_last = edit.row_axis || !edit.band.has_value() ? Sheet::kMaxRows - 1U : edit.band->last;
    moved.col_first = edit.row_axis ? across_first : origin;
    moved.col_last = !edit.row_axis || !edit.band.has_value() ? Sheet::kMaxCols - 1U : edit.band->last;
    mutator.mark_range_dependents_dirty(moved);
  }
  // Map each snapshotted anchor through the edit; deleted anchors and non-formula cells are skipped.
  for (const SpillFootprint& footprint : committed_before) {
    const ShiftedCell shift = shift_cell_coords_for_row_col_edit(edit, footprint.anchor_row, footprint.anchor_col);
    if (!shift.kept) {
      continue;
    }
    const Cell* cell = target.cell_at(shift.new_row, shift.new_col);
    if (cell == nullptr || cell->formula_text.empty()) {
      continue;
    }
    mutator.mark_dirty(make_node(sheet_index, shift.new_row, shift.new_col));
  }
  std::vector<BlockedSpillFootprint> blocked_mapped = remap_blocked_spill_footprints(blocked_before, edit);
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
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index,
                                      StructuralEdit{row, count, false, true});
}

Expected<void, Error> Workbook::delete_rows(std::size_t sheet_index, std::uint32_t row, std::uint32_t count) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index,
                                      StructuralEdit{row, count, true, true});
}

Expected<void, Error> Workbook::insert_cols(std::size_t sheet_index, std::uint32_t col, std::uint32_t count) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index,
                                      StructuralEdit{col, count, false, false});
}

Expected<void, Error> Workbook::delete_cols(std::size_t sheet_index, std::uint32_t col, std::uint32_t count) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index,
                                      StructuralEdit{col, count, true, false});
}

// @size-budget: 13 KB
// Insert/Delete Cells: the band entry below and the band paths of the shared row/column edit.
Expected<void, Error> Workbook::insert_cells(std::size_t sheet_index, const MergeRange& range, CellShift shift) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  return apply_band_edit_locked(sheet_index, range, shift, /*is_delete=*/false);
}

Expected<void, Error> Workbook::delete_cells(std::size_t sheet_index, const MergeRange& range, CellShift shift) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  return apply_band_edit_locked(sheet_index, range, shift, /*is_delete=*/true);
}

Expected<void, Error> Workbook::apply_band_edit_locked(std::size_t sheet_index, const MergeRange& range,
                                                       CellShift shift, bool is_delete) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, sheets_.size()));
  const bool known =
      shift == CellShift::kDown || shift == CellShift::kRight || shift == CellShift::kUp || shift == CellShift::kLeft;
  if (!known || (shift == CellShift::kDown || shift == CellShift::kRight) == is_delete) {
    return make_error(FormulonErrorCode::kInvalidArgument,
                      is_delete ? "shift must be up or left" : "shift must be down or right");
  }
  if (range.first_row > range.last_row || range.first_col > range.last_col || range.last_row >= Sheet::kMaxRows ||
      range.last_col >= Sheet::kMaxCols) {
    return make_error(FormulonErrorCode::kInvalidArgument, "range is outside the grid");
  }
  const bool vertical = shift == CellShift::kDown || shift == CellShift::kUp;
  const bool whole_cols = range.first_row == 0U && range.last_row == Sheet::kMaxRows - 1U;
  const bool whole_rows = range.first_col == 0U && range.last_col == Sheet::kMaxCols - 1U;
  if (vertical ? whole_cols : whole_rows) {
    return make_error(FormulonErrorCode::kInvalidArgument,
                      vertical ? "range covers whole columns, which cannot shift vertically"
                               : "range covers whole rows, which cannot shift horizontally");
  }
  StructuralEdit edit{vertical ? range.first_row : range.first_col,
                      vertical ? range.last_row - range.first_row + 1U : range.last_col - range.first_col + 1U,
                      is_delete, vertical};
  // A band over the whole perpendicular extent is the whole row/column edit.
  if (!(vertical ? whole_rows : whole_cols)) {
    edit.band = vertical ? Band{range.first_col, range.last_col} : Band{range.first_row, range.last_row};
  }
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index, edit);
}

}  // namespace formulon
