// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Context-aware, side-effect-free ad-hoc formula evaluation.
//
// Unlike the recalc path (`cell_evaluator.h`), these drivers never insert
// the formula into the dependency graph, never mutate `cell->cached_value`,
// and never commit dynamic-array spills: the `EvalContext` is built WITHOUT
// `with_mutable_sheet`, so `dispatch_array_result` is inert and recursive
// reference resolution memoises only into a per-call `EvalState`. The
// workbook is observed strictly read-only, which is the safety guarantee
// behind exposing these as no-side-effect public API entry points.

#ifndef FORMULON_EVAL_ADHOC_EVAL_H_
#define FORMULON_EVAL_ADHOC_EVAL_H_

#include <cstdint>
#include <string_view>

#include "value.h"

namespace formulon {

class Arena;
class Sheet;
class Workbook;

namespace eval {

class FunctionRegistry;

/// Evaluates `formula` text as if entered at `(row, col)` on `sheet`
/// (which must belong to `workbook`), resolving local + qualified cell
/// references, workbook-scoped defined names, and anchoring ROW()/COLUMN()
/// at `(row, col)`. Full locale / coercion / 1904 fidelity is threaded
/// from the workbook profile.
///
/// The evaluation is read-only: no cell value is mutated and no spill is
/// committed. An Array / spill result is reduced to its top-left element
/// (see below); a scalar result is returned verbatim, including Excel error
/// sentinels.
///
/// `arena` backs any text payloads in the returned `Value` and must outlive
/// it. Parser failure beyond panic-mode recovery surfaces as `#NAME?`,
/// matching the recursive resolver in `EvalContext::resolve_ref`.
///
/// Array-to-scalar reduction is a pragmatic Phase 1 API-shape choice: it
/// takes the first (top-left) element of a multi-cell result. It is NOT
/// Excel implicit intersection (which selects the range element sharing the
/// anchor's row / column and yields `#VALUE!` when there is none) and NOT
/// dynamic-array spilling (which returns the whole array). Multi-cell
/// results (an array/spill envelope) are a Phase 2 follow-up.
Value evaluate_formula_text(const Workbook& workbook, const Sheet& sheet, std::uint32_t row, std::uint32_t col,
                            std::string_view formula, Arena& arena, const FunctionRegistry& registry);

/// Evaluates `formula` as a conditional-formatting rule predicate anchored
/// at `(row, col)` on `sheet`, with relative references written relative to
/// `(anchor_row, anchor_col)` (the CF-applied range's top-left) shifted to
/// the target cell before evaluation. Mirrors the anchor plumbing used by
/// the CF-rule evaluator (`cf_helpers.cpp::parse_shift_evaluate`).
///
/// The evaluated result is coerced with Excel's CF-predicate rules rather
/// than propagated verbatim (see `coerce_cf_predicate`): error / blank /
/// text / numeric-zero all yield `false`; any non-zero number yields
/// `true`. Read-only, same as `evaluate_formula_text`.
bool evaluate_cf_formula(const Workbook& workbook, const Sheet& sheet, std::uint32_t row, std::uint32_t col,
                         std::uint32_t anchor_row, std::uint32_t anchor_col, std::string_view formula, Arena& arena,
                         const FunctionRegistry& registry);

/// Coerces `v` to a conditional-formatting rule outcome per Excel: error →
/// false (rule does not fire), blank → false, text → false, numeric zero →
/// false, numeric non-zero → true, boolean → its own value. Exposed for
/// direct unit testing of the coercion contract.
bool coerce_cf_predicate(const Value& v);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_ADHOC_EVAL_H_
