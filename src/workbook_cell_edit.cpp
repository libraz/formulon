//
// Workbook cell writes: literal, text and formula setters, range erase and
// the dirty-marking around them, all under one hold of the engine mutex.

#include <cstddef>
#include <cstdint>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/dep_graph.h"
#include "eval/recalc_engine.h"
#include "io/dynamic_array_formula.h"
#include "io/future_functions.h"
#include "io/xlsb/ptg_writer.h"
#include "parser/ast.h"
#include "parser/external_book_spelling.h"
#include "parser/formula_prefix.h"
#include "parser/literal_spelling.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "utils/index_sort.h"
#include "utils/status_macros.h"
#include "utils/text_ops.h"
#include "value.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"

namespace formulon {
namespace {
using detail::check_cell_target;
using detail::MayNameExternalBook;

// Refusal of a write into part of a legacy array formula
// block: a block is only ever replaced as a whole, so a write into one of
// its non-anchor cells leaves it unchanged. A write to the anchor replaces
// the whole block.
Expected<void, Error> check_not_part_of_array(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  const std::optional<SpillFootprint> block = sheet.array_block_covering(row, col);
  if (block.has_value() && (block->anchor_row != row || block->anchor_col != col)) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "Cannot change part of an array",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  return Expected<void, Error>::Ok();
}

// Refusal of a write into a What-If data table's result range:
// unlike an array, the whole table is only ever replaced or removed as one.
Expected<void, Error> check_not_part_of_data_table(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  if (sheet.data_table_covering(row, col) != nullptr) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "Cannot change part of a data table",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  return Expected<void, Error>::Ok();
}

// Target and block checks shared by the cell value, text and formula setters.
Expected<void, Error> check_cell_writable(const std::vector<Sheet>& sheets, std::size_t sheet_index, std::uint32_t row,
                                          std::uint32_t col) {
  RETURN_IF_ERROR(check_cell_target(sheet_index, sheets.size(), row, col));
  RETURN_IF_ERROR(check_not_part_of_array(sheets[sheet_index], row, col));
  return check_not_part_of_data_table(sheets[sheet_index], row, col);
}

// When `(row, col)` is a *phantom* of a committed spill region (i.e. a
// spill target that is not the anchor), a write here clears the region via
// `Sheet::set_cell_value` / `set_cell_formula` but leaves the anchor's
// cached value untouched. The anchor is not a dep-graph dependent of the
// phantom, so nothing else dirties it; mark it dirty so the next recalc
// re-evaluates it — re-spilling if the write vacated the cell, or surfacing
// `#SPILL!` if it now blocks the footprint. Must be called BEFORE the sheet
// write, while the region still covers the cell. Precondition: caller holds
// the engine mutex for the lifetime of the `LockedMutator&`.
void mark_spill_anchor_dirty_if_covered(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                                        const std::vector<Sheet>& sheets, std::uint32_t row, std::uint32_t col) {
  const SpillRegion* covering = sheets[sheet_index].spill_region_covering(row, col);
  if (covering == nullptr) {
    return;
  }
  if (covering->anchor_row == row && covering->anchor_col == col) {
    return;  // Writing the anchor itself is already handled by the caller.
  }
  mutator.mark_dirty(make_node(sheet_index, covering->anchor_row, covering->anchor_col));
}

// A failed dynamic-array commit keeps its attempted rectangle even though no
// phantom cells were materialised.  A mutation that vacates any one cell in
// that rectangle must wake the blocked producer so it can retry on the next
// recalc pass.  The Sheet query returns coordinate copies while holding its
// own lock; marking happens after that lock is released, preserving the
// workbook's engine-mutex -> sheet/spill-mutex order.
void mark_blocked_spill_anchors_intersecting(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                                             const std::vector<Sheet>& sheets, const SpillFootprint& rectangle) {
  for (const CellAddress anchor : sheets[sheet_index].blocked_spill_anchors_intersecting(
           rectangle.anchor_row, rectangle.anchor_col, rectangle.rows, rectangle.cols)) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
}

void mark_blocked_spill_anchors_intersecting(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                                             const std::vector<Sheet>& sheets, std::uint32_t row, std::uint32_t col) {
  mark_blocked_spill_anchors_intersecting(mutator, sheet_index, sheets, SpillFootprint{row, col, 1U, 1U});
}

void mark_blocked_spill_anchors_released_by_cell(const eval::RecalcEngine::LockedMutator& mutator,
                                                 std::size_t sheet_index, const std::vector<Sheet>& sheets,
                                                 std::uint32_t row, std::uint32_t col) {
  // A write into an existing spill clears the entire committed rectangle,
  // not just the addressed cell. Snapshot that rectangle before the write and
  // wake every pending producer intersecting it.
  if (const auto rectangle = sheets[sheet_index].committed_spill_footprint_covering(row, col); rectangle.has_value()) {
    mark_blocked_spill_anchors_intersecting(mutator, sheet_index, sheets, *rectangle);
  }
  mark_blocked_spill_anchors_intersecting(mutator, sheet_index, sheets, row, col);
}

// Dirties the cell, its dependents and every spill anchor the write wakes,
// drops the cell's outgoing edges and its error detail. Precondition: caller
// holds the engine mutex; call before the sheet write.
void prepare_cell_overwrite(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                            std::vector<Sheet>& sheets, std::uint32_t row, std::uint32_t col) {
  sheets[sheet_index].clear_error_detail(row, col);
  const eval::CellNodeId node = make_node(sheet_index, row, col);
  mutator.mark_dirty(node);
  detail::mark_dependents_dirty(mutator, node);
  mark_blocked_spill_anchors_released_by_cell(mutator, sheet_index, sheets, row, col);
  mutator.clear_cell_dependencies(node);
  mark_spill_anchor_dirty_if_covered(mutator, sheet_index, sheets, row, col);
}

}  // namespace

Expected<void, Error> Workbook::set_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                               Value value) {
  RETURN_IF_ERROR(check_cell_writable(sheets_, sheet_index, row, col));
  put_cell_value(sheet_index, row, col, value);
  return Expected<void, Error>::Ok();
}

void Workbook::put_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, Value value) {
  // The cell is becoming a literal: drop outgoing edges (the old formula's
  // reads), but preserve incoming edges so other formulas that *read* this
  // cell continue to re-evaluate when the literal changes.
  // `clear_cell_dependencies` does exactly that — `unregister_formula`
  // would also remove the incoming edges, which is wrong here.
  //
  // Mark dependents *before* clearing the cell's outgoing edges so the
  // reverse-edge snapshot still describes the pre-clear graph. (Clearing
  // outgoing edges does not actually drop incoming edges, but doing the
  // mark first keeps the ordering robust against future API changes.)
  //
  // The entire compound mutation — three engine operations plus the
  // `Sheet` write — is performed under a single hold of the engine
  // mutex so a concurrent `recalc_parallel` (which holds the same
  // mutex for the duration of its pass) cannot observe a half-applied
  // edit. Going through the public engine API would release and
  // re-acquire the lock between every step and let the `Sheet` write
  // race against the recalc worker reading the same cell.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  put_cell_value_locked(sheet_index, row, col, value);
}

void Workbook::put_cell_value_locked(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, Value value) {
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  prepare_cell_overwrite(mutator, sheet_index, sheets_, row, col);
  sheets_[sheet_index].set_cell_value(row, col, value);
}

Expected<void, Error> Workbook::set_cell_text(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                              std::string_view text) {
  RETURN_IF_ERROR(check_cell_writable(sheets_, sheet_index, row, col));
  put_cell_text(sheet_index, row, col, text);
  return Expected<void, Error>::Ok();
}

void Workbook::put_cell_text(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, std::string_view text) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  put_cell_text_locked(sheet_index, row, col, text);
}

void Workbook::put_cell_text_locked(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                    std::string_view text) {
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  prepare_cell_overwrite(mutator, sheet_index, sheets_, row, col);
  sheets_[sheet_index].set_cell_text(row, col, text);
}

void Workbook::erase_cells_in(std::size_t sheet_index, const MergeRange& rect) {
  Sheet& sheet = sheets_[sheet_index];
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  const std::uint64_t col_end = static_cast<std::uint64_t>(rect.last_col) + 1U;
  std::vector<CellAddress> stored;
  std::vector<std::uint32_t> styled_rows;
  for (std::uint64_t row = rect.first_row; row <= rect.last_row; ++row) {
    const auto it = sheet.rows().find(static_cast<std::uint32_t>(row));
    if (it == sheet.rows().end()) {
      continue;
    }
    for (const RowCells::Entry entry : it->second.entries_in(rect.first_col, col_end)) {
      stored.push_back(CellAddress{static_cast<std::uint32_t>(row), entry.col});
    }
    if (it->second.run_count_in(rect.first_col, col_end) != 0U) {
      styled_rows.push_back(static_cast<std::uint32_t>(row));
    }
  }
  // Like a literal write, the erase re-evaluates whatever read the cell.
  for (const CellAddress address : stored) {
    prepare_cell_overwrite(mutator, sheet_index, sheets_, address.row, address.col);
    sheet.erase_cell(address.row, address.col);
  }
  // Formatted blanks carry no value, so nothing reads them.
  for (const std::uint32_t row : styled_rows) {
    sheet.clear_style_runs(row, rect.first_col, col_end);
  }
}

namespace {

/// `text` with an `@` inserted at each node `io::xlsb::legacy_intersections` names.
std::string InsertLegacyIntersections(std::string text, const io::xlsb::NameShapes& name_shapes) {
  const std::size_t body_at = !text.empty() && text.front() == '=' ? 1U : 0U;
  Arena arena;
  const parser::AstNode* root = parser::parse_strict(std::string_view(text).substr(body_at), arena);
  if (root == nullptr) {
    return text;
  }
  // Byte offsets into the body; a node's range counts UTF-16 units.
  const std::string_view body = std::string_view(text).substr(body_at);
  std::vector<std::uint32_t> at;
  for (const parser::AstNode* node : io::xlsb::legacy_intersections(*root, name_shapes)) {
    if (node->kind() != parser::NodeKind::ImplicitIntersection) {
      // A parenthesised operand takes the `@` inside its parentheses: `(@A1:A3)`.
      auto start = static_cast<std::uint32_t>(utf16_to_byte_offset(body, node->range().start));
      for (std::uint16_t pairs = node->paren_depth(); pairs > 0U && body_at + start < text.size(); ++start) {
        if (text[body_at + start] == '(') {
          --pairs;
        }
      }
      at.push_back(start);
    }
  }
  // Right to left, so each insertion leaves the earlier offsets valid.
  sort_ascending(at);
  for (auto it = at.rbegin(); it != at.rend(); ++it) {
    text.insert(body_at + *it, 1U, '@');
  }
  return text;
}

}  // namespace

void Workbook::apply_legacy_implicit_intersections() {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  for (std::size_t sheet_index = 0; sheet_index < sheets_.size(); ++sheet_index) {
    const io::xlsb::NameShapes name_shapes = io::name_shapes(*this, sheet_index);
    Sheet& sheet = sheets_[sheet_index];
    for (const CellAddress address : sheet.all_formula_cells()) {
      const Cell* cell = sheet.cell_at(address.row, address.col);
      // A dynamic-array formula already says where it intersects, and a CSE
      // block evaluates as an array.
      if (cell == nullptr || cell->dynamic_array || cell->array_block_rows != 0U) {
        continue;
      }
      std::string text = InsertLegacyIntersections(cell->formula_text, name_shapes);
      if (text != cell->formula_text) {
        sheet.set_cell_formula_text(address.row, address.col, std::move(text));
      }
    }
  }
}

std::string Workbook::with_legacy_intersections(std::string text, std::size_t sheet_index) const {
  return InsertLegacyIntersections(std::move(text), io::name_shapes(*this, sheet_index));
}

std::string Workbook::normalize_formula_text(std::string formula, FormulaSource source) {
  std::string text = parser::spell_storage_operators(parser::strip_storage_prefixes(formula, &io::has_storage_prefix));
  if (source == FormulaSource::kEntered) {
    text = parser::spell_entered_formula(text);
  }
  return text;
}

Expected<void, Error> Workbook::set_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                 std::string formula, FormulaSource source) {
  RETURN_IF_ERROR(check_cell_writable(sheets_, sheet_index, row, col));
  put_cell_formula(sheet_index, row, col, std::move(formula), source, 0U, 0U);
  return Expected<void, Error>::Ok();
}

void Workbook::put_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, std::string formula,
                                FormulaSource source, std::uint32_t block_rows, std::uint32_t block_cols) {
  // The compound mutation runs under a single hold of the engine mutex so a
  // concurrent `recalc_parallel` does not see a half-applied edit -- see the
  // comment in `put_cell_value` for the full rationale.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  put_cell_formula_locked(sheet_index, row, col, std::move(formula), source, block_rows, block_cols);
}

void Workbook::put_cell_formula_locked(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                       std::string formula, FormulaSource source, std::uint32_t block_rows,
                                       std::uint32_t block_cols) {
  // Normalize Excel's `_xlfn.` / `_xlfn._xlws.` / `_xlws.` / `_xlpm.`
  // storage prefixes to the canonical formula-bar form at the single
  // ingestion point every reader (OOXML DOM / SAX, XLSB) and binding
  // funnels through. This keeps the stored `formula_text` (and the
  // dependency-extraction parse below) matching what Excel's formula bar
  // shows, and lets LET / LAMBDA resolve their `_xlpm.`-prefixed
  // parameter names. The transform is a no-op on an already-canonical
  // formula, so hand-authored / test formulas are unaffected. The writer
  // re-applies the prefixes on save for Excel readability. The stored
  // `SINGLE(x)` / `ANCHORARRAY(x)` calls read back as the `@x` / `x#` the
  // formula bar shows. Entered text also takes the whitespace and number
  // literal spelling Excel stores.
  formula = normalize_formula_text(std::move(formula), source);
  if (source == FormulaSource::kEntered && !file_name().empty()) {
    formula = parser::bind_self_book_names(formula, file_name());
  }

  const eval::CellNodeId node = make_node(sheet_index, row, col);

  // Parse the formula in a throwaway arena to extract its dependency list.
  // Strip a leading '=' so the parser sees a bare expression (the tokenizer
  // accepts both shapes, but the dep extractor walks the AST regardless).
  const std::string_view src = strip_formula_prefix(formula);

  Arena tmp_arena;
  parser::AstNode* root = parse_indexable_formula(src, tmp_arena);
  const bool may_name_external_book = MayNameExternalBook(src);

  {
    const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
    // Query the covering spill BEFORE the sheet write clears it, so a write
    // into a live spill's phantom re-dirties the anchor.
    mark_spill_anchor_dirty_if_covered(mutator, sheet_index, sheets_, row, col);
    mark_blocked_spill_anchors_released_by_cell(mutator, sheet_index, sheets_, row, col);
    sheets_[sheet_index].clear_error_detail(row, col);
    // Persist the formula text on the sheet first so a later `recalc()`
    // reads what the user actually typed. This also resets `cached_value`
    // to blank.
    sheets_[sheet_index].set_cell_formula(row, col, std::move(formula));
    // Marked as Excel 365 marks a typed formula; a reader replaces this with
    // the file's own mark.
    sheets_[sheet_index].set_cell_dynamic_array(
        row, col, block_rows == 0U && root != nullptr && io::entered_as_dynamic_array(*this, sheet_index, *root));
    sheets_[sheet_index].set_cell_array_block(row, col, block_rows, block_cols);

    if (root != nullptr) {
      if (may_name_external_book) {
        bind_external_books_unlocked(*root);
      }
      mutator.register_formula(node, *root, *this);
    } else {
      // Hard parse failure, or a valid prefix trailed by unparseable
      // tokens. Drop any stale edges rather than register dependencies for
      // a recovered prefix that is not the whole formula; the cell surfaces
      // `#NAME?` at the next recalc (via the strict gate in cell_evaluator).
      mutator.unregister_formula(node);
    }

    // Mark the cell dirty and propagate to direct dependents.
    mutator.mark_dirty(node);
    detail::mark_dependents_dirty(mutator, node);
  }
}

void Workbook::mark_cell_dependents_dirty(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  if (sheet_index >= sheets_.size() || !Sheet::coord_in_grid(row, col)) {
    return;
  }
  const eval::CellNodeId node = make_node(sheet_index, row, col);
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  detail::mark_dependents_dirty(mutator, node);
}

}  // namespace formulon
