// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Workbook implementation. Wires the OOXML save path and the embedded
// recalc engine. The `RecalcEngine` is held via `unique_ptr` (PIMPL-style)
// so the public header does not need to include the heavyweight recalc /
// dep-graph headers.

#include "workbook.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/dep_graph.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "io/defined_names.h"
#include "io/external_links.h"
#include "io/ooxml_writer.h"
#include "io/passthrough_part.h"
#include "io/styles_reader.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "parser/ref_transforms.h"
#include "pivot/pivot_cache.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {

Workbook::Workbook() : engine_(std::make_unique<eval::RecalcEngine>()), kind_(io::WorkbookKind::kXlsx) {}
Workbook::Workbook(Workbook&&) noexcept = default;
Workbook& Workbook::operator=(Workbook&&) noexcept = default;
Workbook::~Workbook() = default;

Workbook Workbook::create() {
  Workbook wb;
  wb.sheets_.emplace_back(Sheet{std::string("Sheet1")});
  return wb;
}

Workbook Workbook::create_empty() {
  // No default sheet; callers are expected to populate via add_sheet().
  return Workbook{};
}

std::string_view Workbook::intern_text(std::string_view text) {
  text_storage_.emplace_back(text.data(), text.size());
  return std::string_view(text_storage_.back());
}

Sheet& Workbook::add_sheet(std::string name) {
  sheets_.emplace_back(Sheet{std::move(name)});
  return sheets_.back();
}

namespace {

// Excel's structural validation for sheet names: non-empty, ≤ 31
// characters, no `: \ / ? * [ ]`. This is a byte-level check (the
// forbidden set is ASCII), so it works correctly on UTF-8 sheet names
// because none of the disallowed code units appear as continuation
// bytes (all are < 0x80).
bool is_valid_sheet_name_chars(std::string_view name) noexcept {
  for (char byte : name) {
    switch (byte) {
      case ':':
      case '\\':
      case '/':
      case '?':
      case '*':
      case '[':
      case ']':
        return false;
      default:
        break;
    }
  }
  return true;
}

// Quotes `sheet` for use in a formula reference. Excel quotes a sheet
// name with single quotes when the name contains anything other than
// alphanumerics, underscores or periods. We err on the side of always
// considering both shapes (raw and quoted) when matching defined-name
// targets so a rename catches both.
std::string single_quote_sheet(std::string_view sheet) {
  std::string out;
  out.reserve(sheet.size() + 4U);
  out.push_back('\'');
  for (char byte : sheet) {
    if (byte == '\'') {
      // Embedded single quotes double up in OOXML formula syntax.
      out.append("''");
    } else {
      out.push_back(byte);
    }
  }
  out.push_back('\'');
  return out;
}

// Returns true if `byte` is part of an unquoted Excel identifier
// (letter, digit, underscore). UTF-8 continuation bytes (high bit set)
// are deliberately excluded so a sheet name whose preceding byte is a
// multi-byte CJK character still treats the sheet token as starting
// fresh.
constexpr unsigned char kAsciiHighBit = 0x80U;
constexpr bool is_id_byte(char byte) noexcept {
  const auto value = static_cast<unsigned char>(byte);
  if (value >= kAsciiHighBit) {
    return false;
  }
  return (byte >= 'A' && byte <= 'Z') || (byte >= 'a' && byte <= 'z') || (byte >= '0' && byte <= '9') || byte == '_';
}

// Returns true iff `formula` references the sheet `sheet_name` by name
// (either bare `Sheet1!` or quoted `'Sheet One'!`). Strict prefix match
// of the bang-suffixed sheet token followed by a `!`. The check is
// case-insensitive, mirroring Excel's sheet-name resolution semantics.
//
// This is a textual scan deliberately kept out of the parser: the
// AST-level reference shifter (a separate follow-up bundle) handles the
// general case. For this bundle we only need to decide whether a
// defined-name target string mentions the renamed sheet at all so the
// bulk-replace helper below knows whether the entry is affected.
bool formula_references_sheet(std::string_view formula, std::string_view sheet_name) noexcept {
  // Exhaustive bare-token scan: walk every position and check whether a
  // sheet token followed by `!` matches `sheet_name` (case-insensitive).
  // Quoted-sheet matches are handled by also scanning for the
  // single-quoted form.
  const std::string quoted = single_quote_sheet(sheet_name);
  for (std::size_t pos = 0; pos + sheet_name.size() < formula.size(); ++pos) {
    if (formula[pos + sheet_name.size()] != '!') {
      continue;
    }
    if (strings::case_insensitive_eq(formula.substr(pos, sheet_name.size()), sheet_name)) {
      // Reject matches that are part of a longer identifier (e.g. the
      // suffix of `OtherSheet1` should not match `Sheet1`). A sheet
      // token starts at pos==0 or after a non-identifier byte.
      const bool token_start = pos == 0 || !is_id_byte(formula[pos - 1]);
      if (token_start) {
        return true;
      }
    }
  }
  if (quoted.size() + 1U <= formula.size()) {
    for (std::size_t pos = 0; pos + quoted.size() < formula.size(); ++pos) {
      if (formula[pos + quoted.size()] != '!') {
        continue;
      }
      if (strings::case_insensitive_eq(formula.substr(pos, quoted.size()), quoted)) {
        return true;
      }
    }
  }
  return false;
}

// Rewrites every reference to sheet `old_name` in `formula` to
// `new_name` by parsing through the AST and re-formatting. Both bare
// and quoted forms are recognised by the parser; the formatter emits
// the canonical shape (quoted only when the new name's bytes demand
// it). On parse failure or arena exhaustion the input is returned
// unchanged so the caller's behaviour stays conservative — a formula
// we cannot understand is left alone rather than corrupted.
std::string replace_sheet_in_formula(std::string_view formula, std::string_view old_name, std::string_view new_name) {
  // Defined-name targets are typically authored without a leading `=`;
  // tolerate either by parsing the body and prepending the marker back
  // when the source had one.
  std::string_view body = formula;
  bool had_equals = false;
  if (!body.empty() && body.front() == '=') {
    body = body.substr(1);
    had_equals = true;
  }
  Arena arena;
  parser::Parser parser(body, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    return std::string(formula);
  }
  const parser::SheetRenameTransform transform(old_name, new_name);
  const parser::AstNode* shifted = parser::shift_refs(*root, arena, transform);
  if (shifted == nullptr) {
    return std::string(formula);
  }
  std::string out;
  if (had_equals) {
    out.push_back('=');
  }
  out.append(parser::format_formula(*shifted));
  return out;
}

}  // namespace

Expected<void, Error> Workbook::rename_sheet(std::uint32_t index, std::string new_name) {
  if (static_cast<std::size_t>(index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "rename_sheet: index out of range",
                      "index=" + std::to_string(index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  // Validate the new name. Excel rejects empty, > 31 chars, or any of
  // the forbidden ASCII separators.
  constexpr std::size_t kMaxSheetNameLength = 31U;
  if (new_name.empty() || new_name.size() > kMaxSheetNameLength || !is_valid_sheet_name_chars(new_name)) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "rename_sheet: invalid sheet name",
                      "name=\"" + new_name + "\"");
  }
  // Collision check (case-insensitive). A no-op rename — same case-fold
  // as the current name — bypasses the collision check so callers can
  // change only the casing.
  for (std::size_t idx = 0; idx < sheets_.size(); ++idx) {
    if (idx == static_cast<std::size_t>(index)) {
      continue;
    }
    if (strings::case_insensitive_eq(sheets_[idx].name(), new_name)) {
      return make_error(FormulonErrorCode::kInvalidSheetName, "rename_sheet: name collides with another sheet",
                        "name=\"" + new_name + "\" colliding_index=" + std::to_string(idx));
    }
  }

  const std::string old_name = sheets_[index].name();

  // Update workbook-scoped defined names whose target string mentions
  // the renamed sheet by name. Sheet-scoped names (local_sheet_id >= 0)
  // are unaffected: they reference sheets by ordinal index, not by
  // name, and Excel preserves the binding across renames.
  for (io::DefinedName& entry : defined_names_) {
    if (entry.local_sheet_id >= 0) {
      continue;
    }
    if (formula_references_sheet(entry.formula, old_name)) {
      entry.formula = replace_sheet_in_formula(entry.formula, old_name, new_name);
    }
  }
  // Move the rename into the sheet last so the loop above still has the
  // pre-move `new_name` value to read from.
  sheets_[index].set_name(std::move(new_name));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_sheet(std::uint32_t index) {
  if (static_cast<std::size_t>(index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "remove_sheet: index out of range",
                      "index=" + std::to_string(index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (sheets_.size() <= 1U) {
    return make_error(FormulonErrorCode::kCannotRemoveLastSheet, "remove_sheet: cannot remove the only sheet",
                      "sheet_count=" + std::to_string(sheets_.size()));
  }

  const std::string removed_name = sheets_[index].name();

  // Drop dep-graph nodes for every populated cell on the removed sheet,
  // then erase the sheet itself. Both halves run under a single hold of
  // the engine mutex so a concurrent `recalc_parallel` either sees the
  // sheet (and its graph nodes) fully present or fully gone, never a
  // half-erased intermediate where the graph still names a cell whose
  // sheet vector has already shifted.
  //
  // The graph stores reverse edges, so other sheets' formulas that read
  // into the removed sheet keep their edges — but those edges are now
  // dangling. The next recalc catches them naturally because the source
  // cell is gone.
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
    const Sheet& removed = sheets_[index];
    for (const auto& [row, cells] : removed.rows()) {
      for (std::size_t col = 0; col < cells.size(); ++col) {
        eval::CellNodeId node{static_cast<std::uint16_t>(index), row, static_cast<std::uint32_t>(col)};
        mutator.unregister_formula(node);
      }
    }
    sheets_.erase(sheets_.begin() + static_cast<std::ptrdiff_t>(index));
  }

  // Drop defined names that target the removed sheet. We compare against
  // the removed sheet's name verbatim; sheet-scoped names whose
  // local_sheet_id matches the removed index are also dropped (the
  // index they referenced no longer exists). Note: indices for
  // sheet-scoped names that point at sheets *after* the removed one are
  // shifted down by 1 to preserve the binding. This matches Excel's UI
  // observation that sheet-scoped names stay attached to the same sheet
  // as the workbook is rearranged.
  std::vector<io::DefinedName> retained;
  retained.reserve(defined_names_.size());
  for (io::DefinedName& entry : defined_names_) {
    if (entry.local_sheet_id == static_cast<std::int32_t>(index)) {
      continue;
    }
    if (entry.local_sheet_id < 0 && formula_references_sheet(entry.formula, removed_name)) {
      // Workbook-scoped name targeting the removed sheet: drop it.
      continue;
    }
    if (entry.local_sheet_id > static_cast<std::int32_t>(index)) {
      entry.local_sheet_id -= 1;
    }
    retained.push_back(std::move(entry));
  }
  defined_names_ = std::move(retained);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::move_sheet(std::uint32_t from_index, std::uint32_t to_index) {
  if (static_cast<std::size_t>(from_index) >= sheets_.size() || static_cast<std::size_t>(to_index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "move_sheet: index out of range",
                      "from=" + std::to_string(from_index) + " to=" + std::to_string(to_index) +
                          " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (from_index == to_index) {
    return Expected<void, Error>::Ok();
  }

  // `to_index` is the destination in the *post-removal* sheet list, which
  // matches Excel's UI semantics. Implementation: lift the sheet out,
  // then insert at the destination. The sheet vector mutation and the
  // subsequent per-cell `mark_dirty` loop run under a single hold of
  // the engine mutex so a concurrent `recalc_parallel` either sees the
  // pre-move workbook (sheets in their original order, dirty set
  // unchanged) or the fully patched one, never a half-applied move
  // where `sheets_` has been reordered but the dep-graph still indexes
  // cells by their pre-move sheet_id.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  Sheet moving = std::move(sheets_[from_index]);
  sheets_.erase(sheets_.begin() + static_cast<std::ptrdiff_t>(from_index));
  sheets_.insert(sheets_.begin() + static_cast<std::ptrdiff_t>(to_index), std::move(moving));

  // Patch sheet-scoped defined names so their `local_sheet_id` follows
  // the rearrangement. Workbook-scoped names reference sheets by name,
  // so they need no update.
  const auto from_id = static_cast<std::int32_t>(from_index);
  const auto to_id = static_cast<std::int32_t>(to_index);
  for (io::DefinedName& entry : defined_names_) {
    if (entry.local_sheet_id < 0) {
      continue;
    }
    if (entry.local_sheet_id == from_id) {
      entry.local_sheet_id = to_id;
      continue;
    }
    if (from_id < to_id && entry.local_sheet_id > from_id && entry.local_sheet_id <= to_id) {
      entry.local_sheet_id -= 1;
    } else if (from_id > to_id && entry.local_sheet_id >= to_id && entry.local_sheet_id < from_id) {
      entry.local_sheet_id += 1;
    }
  }

  // The recalc engine's `CellNodeId.sheet_id` is the workbook-relative
  // index, so a move invalidates every dep-graph edge whose endpoint
  // sits on a moved sheet. Rather than rebuild the graph we conservatively
  // mark every cell on every affected sheet (the whole `[min..max]`
  // window) dirty so the next `recalc()` re-registers their edges.
  // Cells outside the window are unaffected. This mirrors Excel's
  // post-rearrange behaviour where downstream formulas re-evaluate.
  const std::uint32_t window_lo = std::min(from_index, to_index);
  const std::uint32_t window_hi = std::max(from_index, to_index);
  for (std::uint32_t sheet_idx = window_lo; sheet_idx <= window_hi; ++sheet_idx) {
    const Sheet& target = sheets_[sheet_idx];
    for (const auto& [row, cells] : target.rows()) {
      for (std::size_t col = 0; col < cells.size(); ++col) {
        eval::CellNodeId node{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)};
        mutator.mark_dirty(node);
      }
    }
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_defined_name(std::string name, std::string formula) {
  return set_defined_name_scoped(std::move(name), std::move(formula), -1);
}

Expected<void, Error> Workbook::set_defined_name_scoped(std::string name, std::string formula,
                                                        std::int32_t local_sheet_id) {
  if (name.empty()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_defined_name_scoped: name is empty");
  }
  if (local_sheet_id < -1 || (local_sheet_id >= 0 && static_cast<std::size_t>(local_sheet_id) >= sheets_.size())) {
    return make_error(
        FormulonErrorCode::kInvalidArgument, "set_defined_name_scoped: local_sheet_id out of range",
        "local_sheet_id=" + std::to_string(local_sheet_id) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  // Case-insensitive lookup: Excel resolves defined names case-folded.
  // Restrict the search to the requested scope.
  for (auto it = defined_names_.begin(); it != defined_names_.end(); ++it) {
    if (it->local_sheet_id != local_sheet_id) {
      continue;
    }
    if (strings::case_insensitive_eq(it->name, name)) {
      if (formula.empty()) {
        defined_names_.erase(it);
      } else {
        it->formula = std::move(formula);
      }
      return Expected<void, Error>::Ok();
    }
  }
  // No existing entry; appending only makes sense if a formula is
  // supplied. An empty-formula "remove" against a non-existent name is
  // a successful no-op.
  if (formula.empty()) {
    return Expected<void, Error>::Ok();
  }
  io::DefinedName entry;
  entry.name = std::move(name);
  entry.formula = std::move(formula);
  entry.local_sheet_id = local_sheet_id;
  defined_names_.push_back(std::move(entry));
  return Expected<void, Error>::Ok();
}

const Sheet* Workbook::sheet_by_name(std::string_view name) const noexcept {
  for (const Sheet& s : sheets_) {
    if (strings::case_insensitive_eq(s.name(), name)) {
      return &s;
    }
  }
  return nullptr;
}

std::size_t Workbook::sheet_index_by_name(std::string_view name) const noexcept {
  for (std::size_t i = 0; i < sheets_.size(); ++i) {
    if (strings::case_insensitive_eq(sheets_[i].name(), name)) {
      return i;
    }
  }
  return static_cast<std::size_t>(-1);
}

void Workbook::add_pivot_cache(std::unique_ptr<pivot::PivotCache> cache) {
  if (cache == nullptr) {
    return;
  }
  pivot_caches_.push_back(std::move(cache));
}

const pivot::PivotCache* Workbook::find_pivot_cache(std::uint32_t cache_id) const noexcept {
  for (const std::unique_ptr<pivot::PivotCache>& c : pivot_caches_) {
    if (c != nullptr && c->cache_id() == cache_id) {
      return c.get();
    }
  }
  return nullptr;
}

Expected<std::vector<std::uint8_t>, Error> Workbook::save() const {
  return io::write_ooxml(*this);
}

namespace {

// Builds a CellNodeId for a workbook-relative coordinate. Sheet ids fit in
// uint16_t per the dep graph contract; Excel allows at most a few thousand
// sheets per workbook, well within range.
eval::CellNodeId make_node(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  return eval::CellNodeId{static_cast<std::uint16_t>(sheet_index), row, col};
}

// Eagerly marks every existing dependent of `cell` dirty in `mutator`'s
// engine. The next `recalc()` pass would discover them via BFS anyway,
// but eager marking keeps the dirty set self-consistent between
// mutations and matches what callers see when they introspect the
// engine via `recalc_engine()`.
//
// Precondition: caller holds the engine mutex for the lifetime of the
// `LockedMutator&`. The helper routes through the facade so the
// compound mutation in `set_cell_value` / `set_cell_formula` stays
// under a single critical section, which keeps the `Sheet` write that
// follows from racing against a concurrent `recalc_parallel`.
void mark_dependents_dirty(const eval::RecalcEngine::LockedMutator& mutator, eval::CellNodeId cell) {
  for (eval::CellNodeId dep : mutator.dep_graph().dependents_of(cell)) {
    mutator.mark_dirty(dep);
  }
}

}  // namespace

Expected<void, Error> Workbook::set_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                               Value value) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_value: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheets_.size()));
  }

  const eval::CellNodeId node = make_node(sheet_index, row, col);

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
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
    mutator.mark_dirty(node);
    mark_dependents_dirty(mutator, node);
    mutator.clear_cell_dependencies(node);
    sheets_[sheet_index].set_cell_value(row, col, value);
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                 std::string formula) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_formula: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheets_.size()));
  }

  const eval::CellNodeId node = make_node(sheet_index, row, col);

  // Parse the formula in a throwaway arena to extract its dependency list.
  // Strip a leading '=' so the parser sees a bare expression (the tokenizer
  // accepts both shapes, but the dep extractor walks the AST regardless).
  std::string_view src = formula;
  if (!src.empty() && src.front() == '=') {
    src.remove_prefix(1);
  }

  Arena tmp_arena;
  parser::Parser parser(src, tmp_arena);
  parser::AstNode* root = parser.parse();

  // The compound mutation runs under a single hold of the engine mutex
  // so a concurrent `recalc_parallel` does not see a half-applied
  // edit — see the comment in `set_cell_value` for the full rationale.
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
    // Persist the formula text on the sheet first so a later `recalc()`
    // reads what the user actually typed. This also resets `cached_value`
    // to blank.
    sheets_[sheet_index].set_cell_formula(row, col, std::move(formula));

    if (root != nullptr) {
      mutator.register_formula(node, *root, *this);
    } else {
      // Parser failed beyond recovery (typically empty input). Drop any
      // stale edges so we do not retain spurious dependencies; the cell
      // will surface `#NAME?` at the next recalc.
      mutator.unregister_formula(node);
    }

    // Mark the cell dirty and propagate to direct dependents.
    mutator.mark_dirty(node);
    mark_dependents_dirty(mutator, node);
  }
  return Expected<void, Error>::Ok();
}

Expected<eval::RecalcStats, Error> Workbook::recalc(const eval::FunctionRegistry& registry) {
  return engine_->recalc(*this, registry);
}

Expected<void, Error> Workbook::recalc_parallel(const eval::FunctionRegistry& registry,
                                                const eval::SchedulerConfig& cfg, eval::SchedulerStats* stats) {
  return eval::recalc_parallel(*this, registry, cfg, stats);
}

void Workbook::set_iterative_options(eval::IterativeOptions opts) {
  engine_->set_iterative_options(opts);
}

const eval::IterativeOptions& Workbook::iterative_options() const noexcept {
  return engine_->iterative_options();
}

Expected<eval::RecalcStats, Error> Workbook::partial_recalc(const eval::FunctionRegistry& registry,
                                                            const eval::SheetCellRange& viewport) {
  return engine_->partial_recalc(*this, registry, viewport);
}

void Workbook::set_iterative_progress(eval::IterativeProgressCb cb, void* user_data) noexcept {
  engine_->set_iterative_progress(cb, user_data);
}

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

// Rewrites a single formula text through `transform`. Returns the
// rewritten body alongside a flag indicating whether any reference was
// actually changed; an unchanged body lets the caller skip the
// dep-graph re-register and avoid touching the cell.
struct FormulaRewriteResult {
  std::string text;
  bool changed = false;
  bool error = false;  // Parse failure or arena exhaustion.
};

FormulaRewriteResult rewrite_formula(std::string_view formula, const parser::RefTransform& transform) {
  FormulaRewriteResult out;
  out.text.assign(formula);
  std::string_view body = formula;
  bool had_equals = false;
  if (!body.empty() && body.front() == '=') {
    body = body.substr(1);
    had_equals = true;
  }
  if (body.empty()) {
    return out;
  }
  Arena arena;
  parser::Parser parser(body, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    out.error = true;
    return out;
  }
  const parser::AstNode* shifted = parser::shift_refs(*root, arena, transform);
  if (shifted == nullptr) {
    out.error = true;
    return out;
  }
  if (shifted == root) {
    return out;  // Identity walk; no change required.
  }
  std::string rewritten;
  if (had_equals) {
    rewritten.push_back('=');
  }
  rewritten.append(parser::format_formula(*shifted));
  out.text = std::move(rewritten);
  out.changed = true;
  return out;
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
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    Sheet& sheet = sheets[sheet_idx];
    const bool local_means_target = strings::case_insensitive_eq(sheet.name(), target_sheet);
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
      const std::vector<Cell>& cells = it->second;
      for (std::size_t col = 0; col < cells.size(); ++col) {
        const Cell& cell = cells[col];
        if (cell.formula_text.empty()) {
          continue;
        }
        std::string_view body = cell.formula_text;
        bool had_equals = false;
        if (!body.empty() && body.front() == '=') {
          body = body.substr(1);
          had_equals = true;
        }
        if (body.empty()) {
          continue;
        }
        Arena arena;
        parser::Parser parser(body, arena);
        parser::AstNode* root = parser.parse();
        if (root == nullptr || !parser.errors().empty()) {
          continue;  // Unparseable formula; leave alone (matches Excel "carry through unchanged").
        }
        const parser::AstNode* shifted = parser::shift_refs(*root, arena, transform);
        const bool ast_changed = (shifted != nullptr && shifted != root);

        // Resolve the cell's post-shift coordinates. Only cells on the
        // target sheet move; off-target cells keep their (row, col).
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

        if (!ast_changed && !cell_moves && !dropped) {
          continue;  // Cell unaffected by the edit.
        }

        // Rewrite the formula text only when the AST actually changed.
        // Cells whose key shifts but whose refs are stable keep the
        // existing text + cached value; the physical move done later
        // by Sheet::insert_rows / delete_rows carries both into place.
        if (ast_changed) {
          std::string rewritten;
          if (had_equals) {
            rewritten.push_back('=');
          }
          rewritten.append(parser::format_formula(*shifted));
          sheet.set_cell_formula(row, static_cast<std::uint32_t>(col), std::move(rewritten));
        }

        const eval::CellNodeId old_node{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)};
        if (dropped) {
          // Cell will vanish from the sheet — drop its dep-graph node.
          mutator.unregister_formula(old_node);
          continue;
        }

        // Re-key the dep-graph entry from OLD to NEW coords. Even
        // when the AST is identity, the cell's key changes so the
        // engine must learn the new key; otherwise dirty marks and
        // dependent edges fire against an empty / wrong slot.
        if (cell_moves) {
          mutator.unregister_formula(old_node);
        }
        const eval::CellNodeId new_node{static_cast<std::uint16_t>(sheet_idx), new_row, new_col};
        const parser::AstNode& effective_ast = ast_changed ? *shifted : *root;
        mutator.register_formula(new_node, effective_ast, workbook);
        mutator.mark_dirty(new_node);
      }
    }
  }
}

void rewrite_defined_names(std::vector<io::DefinedName>& names, const parser::RefTransform& transform) {
  for (io::DefinedName& entry : names) {
    FormulaRewriteResult result = rewrite_formula(entry.formula, transform);
    if (result.changed) {
      entry.formula = std::move(result.text);
    }
  }
}

Expected<void, Error> apply_row_col_edit(Workbook& wb, std::size_t sheet_index, parser::RowColAxis axis,
                                         std::uint32_t origin, std::uint32_t count, const char* op_name) {
  if (sheet_index >= wb.sheet_count()) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op_name) + ": sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  if (count == 0U) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op_name) + ": count must be >= 1");
  }
  const std::uint32_t bound = (axis == parser::RowColAxis::kRow) ? Sheet::kMaxRows : Sheet::kMaxCols;
  if (origin >= bound) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op_name) + ": origin out of bounds",
                      "origin=" + std::to_string(origin));
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
// lock.
Expected<void, Error> apply_row_col_edit_operation(Workbook& wb, std::vector<Sheet>& sheets,
                                                   const eval::RecalcEngine::LockedMutator& mutator,
                                                   std::vector<io::DefinedName>& defined_names, std::size_t sheet_index,
                                                   parser::RowColAxis axis, parser::RowColEdit edit,
                                                   std::uint32_t origin, std::uint32_t count, const char* op_name) {
  RETURN_IF_ERROR(apply_row_col_edit(wb, sheet_index, axis, origin, count, op_name));
  const std::string target_sheet_name = sheets[sheet_index].name();
  rewrite_formulas_for_row_col_edit(sheets, mutator, wb, target_sheet_name, axis, edit, origin, count);
  const parser::RowColShiftTransform name_transform(target_sheet_name, axis, edit, origin, count);
  rewrite_defined_names(defined_names, name_transform);
  Sheet& target = sheets[sheet_index];
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

Expected<void, Error> Workbook::set_cell_xf_index(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                  std::uint32_t xf_index) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_xf_index: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  Sheet& sheet = sheets_[sheet_index];
  if (sheet.cell_at(row, col) == nullptr) {
    sheet.set_cell_value(row, col, Value::blank());
  }
  sheet.set_cell_xf_index(row, col, xf_index);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
