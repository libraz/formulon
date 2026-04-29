// Copyright 2026 libraz. Licensed under the MIT License.
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
#include "io/ooxml_writer.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {

Workbook::Workbook() : engine_(std::make_unique<eval::RecalcEngine>()) {}
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

Sheet& Workbook::add_sheet(std::string name) {
  sheets_.emplace_back(Sheet{std::move(name)});
  return sheets_.back();
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

Expected<std::vector<std::uint8_t>, Error> Workbook::save() const { return io::write_ooxml(*this); }

namespace {

// Builds a CellNodeId for a workbook-relative coordinate. Sheet ids fit in
// uint16_t per the dep graph contract; Excel allows at most a few thousand
// sheets per workbook, well within range.
eval::CellNodeId make_node(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  return eval::CellNodeId{static_cast<std::uint16_t>(sheet_index), row, col};
}

// Eagerly marks every existing dependent of `cell` dirty in `engine`. The
// next `recalc()` pass would discover them via BFS anyway, but eager
// marking keeps the dirty set self-consistent between mutations and
// matches what callers see when they introspect the engine via
// `recalc_engine()`.
void mark_dependents_dirty(eval::RecalcEngine& engine, eval::CellNodeId cell) {
  for (eval::CellNodeId dep : engine.dep_graph().dependents_of(cell)) {
    engine.mark_dirty(dep);
  }
}

}  // namespace

Expected<void, Error> Workbook::set_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                               Value value) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_value: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) +
                          " sheet_count=" + std::to_string(sheets_.size()));
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
  engine_->mark_dirty(node);
  mark_dependents_dirty(*engine_, node);
  engine_->clear_cell_dependencies(node);

  sheets_[sheet_index].set_cell_value(row, col, value);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                 std::string formula) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_formula: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) +
                          " sheet_count=" + std::to_string(sheets_.size()));
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

  // Persist the formula text on the sheet first so a later `recalc()`
  // reads what the user actually typed. This also resets `cached_value`
  // to blank.
  sheets_[sheet_index].set_cell_formula(row, col, std::move(formula));

  if (root != nullptr) {
    engine_->register_formula(node, *root, *this);
  } else {
    // Parser failed beyond recovery (typically empty input). Drop any
    // stale edges so we do not retain spurious dependencies; the cell
    // will surface `#NAME?` at the next recalc.
    engine_->unregister_formula(node);
  }

  // Mark the cell dirty and propagate to direct dependents.
  engine_->mark_dirty(node);
  mark_dependents_dirty(*engine_, node);
  return Expected<void, Error>::Ok();
}

Expected<eval::RecalcStats, Error> Workbook::recalc(const eval::FunctionRegistry& registry) {
  return engine_->recalc(*this, registry);
}

void Workbook::set_iterative_options(eval::IterativeOptions opts) { engine_->set_iterative_options(opts); }

const eval::IterativeOptions& Workbook::iterative_options() const noexcept { return engine_->iterative_options(); }

}  // namespace formulon
