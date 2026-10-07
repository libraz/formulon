//
// Implementation of `SpillProducerIndex`; see `spill_producer_index.h`.

#include "eval/spill_producer_index.h"

#include <string_view>

#include "cell.h"
#include "eval/builtin_names.h"
#include "eval/function_registry.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Cells probed straight above and straight left of a spill target for a blocker; none found admits conservatively.
constexpr std::uint32_t kSpillBlockerProbeCells = 64U;

}  // namespace

bool blocks_spill(const Cell* cell) noexcept {
  return cell != nullptr && (!cell->formula_text.empty() || !cell->cached_value.is_blank());
}

SpillTarget::SpillTarget(const Sheet& sheet, std::uint32_t target_row, std::uint32_t target_col) noexcept
    : row(target_row), col(target_col) {
  for (std::uint32_t step = 1U; step <= kSpillBlockerProbeCells && step <= row; ++step) {
    if (blocks_spill(sheet.cell_at(row - step, col))) {
      blocker_row = static_cast<std::int64_t>(row - step);
      break;
    }
  }
  for (std::uint32_t step = 1U; step <= kSpillBlockerProbeCells && step <= col; ++step) {
    if (blocks_spill(sheet.cell_at(row, col - step))) {
      blocker_col = static_cast<std::int64_t>(col - step);
      break;
    }
  }
}

bool SpillTarget::reachable_from(CellNodeId producer) const noexcept {
  if (producer.row > row || producer.col > col) {
    return false;
  }
  const auto producer_row = static_cast<std::int64_t>(producer.row);
  const auto producer_col = static_cast<std::int64_t>(producer.col);
  if (producer_row <= blocker_row && !(producer_row == blocker_row && producer.col == col)) {
    return false;
  }
  return !(producer_col <= blocker_col && !(producer_col == blocker_col && producer.row == row));
}

void SpillProducerIndex::update(CellNodeId cell, SpillPotential potential) {
  if (cell.sheet_id >= by_sheet_.size()) {
    if (potential == SpillPotential::kNever) {
      return;
    }
    by_sheet_.resize(static_cast<std::size_t>(cell.sheet_id) + 1U);
  }
  auto& producers = by_sheet_[cell.sheet_id];
  if (potential == SpillPotential::kNever) {
    producers.erase(cell);
  } else {
    producers.insert_or_assign(cell, potential);
  }
}

const SpillPotential* SpillProducerIndex::find(CellNodeId producer) const {
  if (producer.sheet_id >= by_sheet_.size()) {
    return nullptr;
  }
  const auto it = by_sheet_[producer.sheet_id].find(producer);
  return it == by_sheet_[producer.sheet_id].end() ? nullptr : &it->second;
}

bool SpillProducerIndex::admissible(const Workbook& workbook, const FunctionRegistry& registry, CellNodeId producer,
                                    SpillPotential static_potential, Arena& potential_arena) {
  if (producer.sheet_id >= workbook.sheet_count()) {
    return false;
  }
  SpillPotential candidate = static_potential;
  if (candidate == SpillPotential::kNeedsRegistry) {
    const Sheet& producer_sheet = workbook.sheet(producer.sheet_id);
    const Cell* producer_cell = producer_sheet.cell_at(producer.row, producer.col);
    if (producer_cell == nullptr || producer_cell->formula_text.empty()) {
      return false;
    }
    std::string_view formula = producer_cell->formula_text;
    if (!formula.empty() && formula.front() == '=') {
      formula.remove_prefix(1);
    }
    potential_arena.reset();
    parser::AstNode* producer_ast = parse_formula_entry(formula, potential_arena);
    if (producer_ast == nullptr) {
      // A stale/unparseable candidate cannot commit a spill; its cell will
      // surface the ordinary formula error if it is dirty.
      return false;
    }
    candidate = spill_potential(*producer_ast, registry);
  }
  return candidate != SpillPotential::kNever;
}

}  // namespace formulon::eval
