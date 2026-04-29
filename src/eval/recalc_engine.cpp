// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `RecalcEngine`. See `recalc_engine.h` for the public
// contract.

#include "eval/recalc_engine.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/dep_extractor.h"
#include "eval/dep_graph.h"
#include "eval/dirty_set.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/tree_walker.h"
#include "eval/volatile_tracker.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Returns true when the SCC `component` represents a real cycle: either
// more than one cell, or a singleton with a self-edge in the dep graph.
// Plain singletons evaluate normally; cyclic SCCs short-circuit to #REF!.
bool is_cyclic_component(const std::vector<CellNodeId>& component, const DepGraph& graph) {
  if (component.size() > 1U) {
    return true;
  }
  // Singleton: only a self-loop (cell depends on itself) is cyclic.
  const CellNodeId only = component.front();
  std::vector<CellNodeId> deps = graph.dependencies_of(only);
  return std::find(deps.begin(), deps.end(), only) != deps.end();
}

// Re-parses and evaluates the formula at `cell` on `sheet`. Returns the
// result (or an Excel error sentinel on parse failure). The caller is
// responsible for writing the result back to the cell.
//
// When `iterative_mode` is true the EvalContext is built WITHOUT an
// EvalState binding. The 3-arg `EvalContext::resolve_ref` overload then
// short-circuits every formula-cell read to `cell->cached_value` instead
// of recursing into the cell's formula text. This is exactly what the
// iterative solver wants: each member of the SCC reads its peers'
// most-recently-committed values rather than triggering re-entrant
// evaluation (which would surface `#REF!` for any back-edge inside the
// cycle and break the fixed-point search). Cross-sheet resolution
// still works because the workbook pointer is bound regardless.
Value evaluate_cell(Workbook& workbook, Sheet& sheet, const Cell& cell_data, std::uint32_t row, std::uint32_t col,
                    const FunctionRegistry& registry, Arena& arena, bool iterative_mode = false) {
  // Strip the leading '=' before parsing (the parser expects an expression,
  // not an assignment). `formula_text` is guaranteed non-empty by the
  // caller because we only feed formula cells through here.
  std::string_view src = cell_data.formula_text;
  if (!src.empty() && src.front() == '=') {
    src.remove_prefix(1);
  }

  parser::Parser parser(src, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr) {
    // Parser failure beyond panic-mode recovery (typically empty input or
    // arena exhaustion). #NAME? matches the existing recursive resolver
    // behaviour in `EvalContext::resolve_ref`.
    return Value::error(ErrorCode::Name);
  }

  // Build an EvalContext that authorises spill writes on `sheet` and
  // anchors the formula at the cell being evaluated. Outside of iterative
  // mode each cell gets its own EvalState so the recursion stack / memo
  // are scoped to one top-level evaluate() call — the dep graph already
  // handles the workbook-wide ordering.
  EvalState state;
  EvalContext ctx;
  if (iterative_mode) {
    // Workbook-bound, state-less context: formula refs short-circuit to
    // their cached values, which is what the solver iterates against.
    ctx = EvalContext::workbook_only(workbook, sheet)
              .with_mutable_sheet(sheet)
              .with_formula_cell(row, col);
  } else {
    ctx = EvalContext(workbook, sheet, state)
              .with_mutable_sheet(sheet)
              .with_formula_cell(row, col);
  }

  Value result = evaluate(*root, arena, registry, ctx);
  // If the top-level evaluator produced an Array (e.g. a SEQUENCE() at the
  // anchor), commit the spill and return the anchor scalar. Mirrors the
  // logic in `EvalContext::resolve_ref` for recursive Array results.
  if (result.is_array()) {
    result = ctx.dispatch_array_result(result);
  }
  return result;
}

}  // namespace

// ----------------------------------------------------------------------------
// RecalcEngine
// ----------------------------------------------------------------------------

RecalcEngine::RecalcEngine() : arena_(std::make_unique<Arena>()) {}
RecalcEngine::~RecalcEngine() = default;

void RecalcEngine::register_formula(CellNodeId cell, const parser::AstNode& ast, const Workbook& workbook) {
  // Drop the cell's previous outgoing edges so re-registration is a clean
  // rewrite (the new dependency set may differ from the old one).
  graph_.clear_dependencies_of(cell);
  // Same for the volatile flag — only re-register if the new AST is still
  // volatile.
  volatiles_.unregister_cell(cell);

  const ExtractedDeps deps = extract_deps(ast, cell.sheet_id, workbook);
  for (CellNodeId dep : deps.cell_deps) {
    graph_.add_dependency(cell, dep);
  }
  if (deps.is_volatile) {
    volatiles_.register_cell(cell);
  }
}

void RecalcEngine::unregister_formula(CellNodeId cell) {
  graph_.remove_node(cell);
  volatiles_.unregister_cell(cell);
}

void RecalcEngine::clear_cell_dependencies(CellNodeId cell) {
  graph_.clear_dependencies_of(cell);
  volatiles_.unregister_cell(cell);
}

void RecalcEngine::mark_dirty(CellNodeId cell) { dirty_.mark(cell); }

Expected<RecalcStats, Error> RecalcEngine::recalc(Workbook& workbook, const FunctionRegistry& registry) {
  RecalcStats stats;

  // ---- Phase 1: seed the dirty set with every volatile cell. ----
  // Volatile formulas re-execute every pass even if their inputs are
  // unchanged. Counting them here also reports how many of the eventually-
  // evaluated cells were forced by volatility.
  volatiles_.for_each([&](CellNodeId cell) {
    if (!dirty_.contains(cell)) {
      dirty_.mark(cell);
    }
    ++stats.volatile_cells;
  });

  // ---- Phase 2: BFS-propagate dirtiness through reverse edges. ----
  // Snapshot the seed list because `dirty_.mark` mutations during BFS
  // would otherwise shift the iteration target.
  std::vector<CellNodeId> bfs_queue;
  bfs_queue.reserve(dirty_.size());
  dirty_.for_each([&](CellNodeId c) { bfs_queue.push_back(c); });
  std::size_t bfs_head = 0;
  while (bfs_head < bfs_queue.size()) {
    const CellNodeId current = bfs_queue[bfs_head++];
    for (CellNodeId dependent : graph_.dependents_of(current)) {
      if (!dirty_.contains(dependent)) {
        dirty_.mark(dependent);
        bfs_queue.push_back(dependent);
      }
    }
  }

  // ---- Phase 3: Tarjan SCC over the entire dep graph. ----
  // Tarjan emits SCCs in reverse-topological order (leaves first), so the
  // evaluation walk below sees a cell's dependencies before the cell
  // itself.
  const std::vector<std::vector<CellNodeId>> sccs = graph_.tarjan_scc();

  // Index sheet pointers once so we can resolve `CellNodeId::sheet_id` to
  // a `Sheet*` without a per-cell lookup. Workbook sheet count is small
  // (typically <= a few dozen).
  const std::size_t sheet_count = workbook.sheet_count();

  // ---- Phase 4: evaluate every dirty SCC. ----
  // Also track which dirty cells have been visited via the SCC walk, so a
  // final sweep can pick up dirty cells with no graph edges (e.g. a
  // standalone `=NOW()` that reads nothing). Tarjan only emits nodes that
  // appear in the forward / reverse adjacency maps; isolated formula
  // cells are absent from both.
  std::unordered_set<CellNodeId, CellNodeIdHash> visited_in_sccs;
  for (const std::vector<CellNodeId>& component : sccs) {
    // Skip components whose intersection with the dirty set is empty —
    // their cells are already up to date.
    bool any_dirty = false;
    for (CellNodeId c : component) {
      if (dirty_.contains(c)) {
        any_dirty = true;
        break;
      }
    }
    if (!any_dirty) {
      continue;
    }

    if (is_cyclic_component(component, graph_)) {
      // Track which members the dispatcher visited regardless of the
      // resolution path so the standalone-dirty sweep does not re-touch
      // them.
      for (CellNodeId c : component) {
        visited_in_sccs.insert(c);
      }

      if (!iterative_.enabled) {
        // Cycle SCC, iterative calc disabled: surface #REF! on every
        // member. Excel's analogous behaviour pops a warning dialog and
        // leaves cells at zero; Formulon collapses the diagnostic into
        // an Excel-visible error sentinel since we have no UI to host a
        // banner.
        for (CellNodeId c : component) {
          if (c.sheet_id >= sheet_count) {
            continue;  // Defensive — should not happen.
          }
          Sheet& sheet = workbook.sheet(c.sheet_id);
          sheet.set_cell_cached_value(c.row, c.col, Value::error(ErrorCode::Ref));
          ++stats.cycle_cells;
        }
        continue;
      }

      // Iterative calc enabled: hand the SCC to the solver. The
      // `evaluate_one` lambda mirrors the singleton path's evaluator
      // glue: parse on the fly, dispatch through `evaluate()`, fold
      // dynamic-array spills back into a scalar anchor. The `commit`
      // lambda writes the new value into the cell store so the next
      // iteration's `evaluate_one` reads the freshest value back.
      auto evaluate_one = [&](CellNodeId c) -> Value {
        if (c.sheet_id >= sheet_count) {
          return Value::error(ErrorCode::Ref);
        }
        Sheet& sheet = workbook.sheet(c.sheet_id);
        const Cell* cell_data = sheet.cell_at(c.row, c.col);
        if (cell_data == nullptr || cell_data->formula_text.empty()) {
          // No formula text: nothing to evaluate. Treat as Blank so the
          // solver still has a value to compare against — this happens
          // only on logic bugs but we degrade gracefully.
          return Value::blank();
        }
        // Reset the bump arena per-evaluation; the previous iteration's
        // arena-owned text payloads have already been written into the
        // sheet's `cached_value` (which currently retains the pointer —
        // see the TODO note further down for the cross-pass text-storage
        // promotion).
        arena_->reset();
        return evaluate_cell(workbook, sheet, *cell_data, c.row, c.col, registry, *arena_,
                             /*iterative_mode=*/true);
      };
      auto commit = [&](CellNodeId c, Value v) {
        if (c.sheet_id >= sheet_count) {
          return;
        }
        Sheet& sheet = workbook.sheet(c.sheet_id);
        sheet.set_cell_cached_value(c.row, c.col, v);
      };

      const IterativeOutcome outcome = run_iterative_solve(component, iterative_, evaluate_one, commit);
      if (outcome.converged) {
        // Solver wrote the converged values into the cell store; count
        // each member as evaluated (singleton-style accounting) plus
        // tagged as iterative.
        for (CellNodeId c : component) {
          (void)c;
          ++stats.cells_evaluated;
          ++stats.iterative_cells;
        }
      } else {
        // Either divergence (solver wrote `#NUM!` on every member) or
        // iteration-limit exhaustion (solver left the last-iteration
        // values in place). The user-visible failure mode in both cases
        // is that the cycle did not resolve, so we count the members in
        // `cycle_cells` to mirror the disabled-iterative-calc accounting.
        // For the iteration-limit case, write `#NUM!` ourselves so the
        // cells do not retain misleading partial values that the user
        // might mistake for a converged result.
        if (!outcome.diverged) {
          for (CellNodeId c : component) {
            if (c.sheet_id >= sheet_count) {
              continue;
            }
            Sheet& sheet = workbook.sheet(c.sheet_id);
            sheet.set_cell_cached_value(c.row, c.col, Value::error(ErrorCode::Num));
          }
        }
        for (CellNodeId c : component) {
          (void)c;
          ++stats.cycle_cells;
        }
      }
      continue;
    }

    // Plain singleton: evaluate the cell.
    const CellNodeId only = component.front();
    visited_in_sccs.insert(only);
    if (only.sheet_id >= sheet_count) {
      continue;
    }
    Sheet& sheet = workbook.sheet(only.sheet_id);
    const Cell* cell_data = sheet.cell_at(only.row, only.col);
    if (cell_data == nullptr || cell_data->formula_text.empty()) {
      // The dep graph may carry pure-input cells (read-only literals that
      // someone reads via `add_dependency`). They have nothing to evaluate.
      continue;
    }
    // Reset the per-pass arena before each evaluate so the bump allocator
    // does not grow without bound across cells. Array results are deep-
    // copied into sheet-owned storage by `Sheet::commit_spill` (via
    // `EvalContext::dispatch_array_result`), so the spill table survives
    // the reset.
    //
    // TODO: scalar Text results currently reference arena-owned bytes;
    // they survive within the recalc pass but not across it. Promote
    // them into a sheet-owned string store before the next iterative
    // solver bundle wires in cross-pass cached values.
    arena_->reset();
    Value result = evaluate_cell(workbook, sheet, *cell_data, only.row, only.col, registry, *arena_);
    sheet.set_cell_cached_value(only.row, only.col, result);
    ++stats.cells_evaluated;
  }

  // ---- Phase 4b: pick up dirty cells with no graph edges. ----
  // Isolated formula cells (e.g. `=NOW()` reading nothing) never enter the
  // dep graph, so Tarjan does not include them in any SCC. Sweep the
  // dirty set for any such cells and evaluate them as plain singletons.
  // Snapshot the set first because evaluation does not mutate `dirty_`,
  // but copying keeps the loop body trivial.
  std::vector<CellNodeId> standalone_dirty;
  dirty_.for_each([&](CellNodeId c) {
    if (visited_in_sccs.count(c) == 0U) {
      standalone_dirty.push_back(c);
    }
  });
  for (CellNodeId c : standalone_dirty) {
    if (c.sheet_id >= sheet_count) {
      continue;
    }
    Sheet& sheet = workbook.sheet(c.sheet_id);
    const Cell* cell_data = sheet.cell_at(c.row, c.col);
    if (cell_data == nullptr || cell_data->formula_text.empty()) {
      continue;
    }
    arena_->reset();
    Value result = evaluate_cell(workbook, sheet, *cell_data, c.row, c.col, registry, *arena_);
    sheet.set_cell_cached_value(c.row, c.col, result);
    ++stats.cells_evaluated;
  }

  // ---- Phase 5: clear the dirty set. ----
  dirty_.clear();
  return stats;
}

}  // namespace formulon::eval
