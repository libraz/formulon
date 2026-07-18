// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of `evaluate_cell_for_recalc`. See `cell_evaluator.h`
// for the contract and the rationale for centralising this logic.

#include "eval/cell_evaluator.h"

#include <cstdint>
#include <string_view>

#include "cell.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/formula_text_utils.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

Value evaluate_cell_for_recalc(Workbook& workbook, Sheet& sheet, const Cell& cell_data, std::uint32_t row,
                               std::uint32_t col, const FunctionRegistry& registry, Arena& arena,
                               const EvaluateCellOptions& opts) {
  // Strip the leading '=' before parsing (the parser expects an expression,
  // not an assignment). `formula_text` is guaranteed non-empty by the
  // caller because we only feed formula cells through here. The shared
  // helper keeps this strip in lockstep across every recalc entry.
  const std::string_view src = strip_formula_prefix(cell_data.formula_text);

  parser::AstNode* root = parser::parse_strict(src, arena);
  if (root == nullptr) {
    // Hard parse failure, or a valid prefix followed by unparseable trailing
    // tokens. Either way we refuse to evaluate a recovered prefix as if it
    // were the whole formula. #NAME? matches the existing recursive resolver
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
  if (opts.iterative_mode) {
    // Workbook-bound, state-less context: formula refs short-circuit to
    // their cached values, which is what the solver iterates against.
    ctx = EvalContext::workbook_only(workbook, sheet)
              .with_excel_profile(workbook.excel_profile())
              .with_date1904(workbook.date1904())
              .with_mutable_sheet(sheet)
              .with_formula_cell(row, col);
  } else {
    ctx = EvalContext(workbook, sheet, state)
              .with_excel_profile(workbook.excel_profile())
              .with_date1904(workbook.date1904())
              .with_mutable_sheet(sheet)
              .with_formula_cell(row, col);
  }
  // The recalc engine owns iterative-calc resolution (SCC detection + the
  // iterative solver). Each per-cell evaluation it issues must be a single
  // pass; suppress the `evaluate()`-level fixed-point driver so the two
  // mechanisms do not double-iterate or fight over divergence accounting.
  ctx = ctx.with_iterative_driver_suppressed();

  Value result = evaluate(*root, arena, registry, ctx);
  // If the top-level evaluator produced an Array (e.g. a SEQUENCE() at the
  // anchor), commit the spill and return the anchor scalar. Mirrors the
  // logic in `EvalContext::resolve_ref` for recursive Array results.
  if (result.is_array()) {
    result = ctx.dispatch_array_result(result);
  }
  return result;
}

}  // namespace eval
}  // namespace formulon
