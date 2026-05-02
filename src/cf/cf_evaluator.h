// Copyright 2026 libraz. Licensed under the MIT License.
//
// Conditional-format rule evaluator. Drives a single `cf::CFRule`
// against a cell's `Value` and returns whether the rule matches.
//
// Current coverage:
//
//   * `ContainsBlanks` / `NotContainsBlanks`
//   * `ContainsErrors` / `NotContainsErrors`
//   * `ContainsText` / `NotContainsText` / `BeginsWith` / `EndsWith` —
//     ASCII case-insensitive substring / prefix / suffix matching of
//     `rule.text` against the cell's text. Non-text cells (numbers,
//     booleans, errors, blanks) do not match either the positive or
//     negative variants; the conservative cross-kind stance mirrors
//     `cellIs`. Empty `rule.text` matches every text cell for the
//     positive variants and never matches for `NotContainsText`.
//   * `CellIs` against a literal `formula1` / `formula2` (the value-
//     only `match_rule(rule, cell_value)` overload), and against an
//     evaluated formula expression (the context-aware
//     `match_rule(rule, cell_value, ctx)` overload).
//   * `Expression` rules whose `formula1` evaluates to a truthy
//     scalar, with relative references shifted from the rule's anchor
//     to the target cell.
//   * `TimePeriod` — bucket-tests the cell's date serial against the
//     `today_serial` carried by `CFEvalContext` (Today / Yesterday /
//     Tomorrow / Last7Days / ThisWeek / LastWeek / NextWeek /
//     ThisMonth / LastMonth / NextMonth). Sunday-start weeks. Only
//     reachable through the context-aware overload.
//
// Deferred to subsequent staging steps:
//
//   * `Top10` / `AboveAverage` / `DuplicateValues` / `UniqueValues`
//     (range-aware rules — need access to all values in the rule's
//     sqref to compute rank / mean / value-counts)
//   * `ColorScale` / `DataBar` / `IconSet` visual computation
//   * Priority chain + stopIfTrue + lazy viewport API
//
// Each step extends `match_rule` and the public `evaluate_*` helpers
// without changing the existing call signatures.
//
// Design references:
//   * backup/plans/20-conditional-format-deep.md §20.4 / §20.5
//   * src/cf/cf_match.h (return-type contract)

#ifndef FORMULON_CF_CF_EVALUATOR_H_
#define FORMULON_CF_CF_EVALUATOR_H_

#include <optional>

#include "cell.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "value.h"

namespace formulon {

class Arena;

namespace eval {
class EvalContext;
class FunctionRegistry;
}  // namespace eval

namespace cf {

/// Returns `true` when `rule` matches the cell carrying `cell_value`.
///
/// Value-only overload. Resolves rule kinds whose decision depends
/// solely on the cell's `Value`:
///   * `ContainsBlanks`     — `cell_value.is_blank()`
///   * `NotContainsBlanks`  — `!cell_value.is_blank()`
///   * `ContainsErrors`     — `cell_value.is_error()`
///   * `NotContainsErrors`  — `!cell_value.is_error()`
///   * `CellIs`             — compares `cell_value` against the literal
///                            parsed from `rule.formula1` (and
///                            `rule.formula2` for `Between` /
///                            `NotBetween`) using `rule.op`. Operands
///                            that need formula evaluation (references,
///                            arithmetic, function calls) fall through
///                            to `false`; use the context-aware
///                            overload below to evaluate them.
///   * `ContainsText` / `NotContainsText` / `BeginsWith` / `EndsWith`
///                          — ASCII case-insensitive substring /
///                            prefix / suffix match of `rule.text`
///                            against the cell's text. Non-text cells
///                            never match.
///
/// All other rule types currently return `false` (their evaluation
/// logic lands in subsequent staging steps). Callers that want to opt
/// out of pre-implementation false-negatives should gate on the
/// rule's `type` and skip rules that aren't yet supported.
bool match_rule(const CFRule& rule, const Value& cell_value);

/// Returns the `CFMatch` payload for a rule that has been determined
/// to match. Used by the integration helpers to convert a positive
/// match into the result list. The skeleton landing covered only
/// `DifferentialFormat`; visual kinds (ColorScale / DataBar / IconSet)
/// are introduced by later steps.
CFMatch make_match(const CFRule& rule);

/// Per-cell evaluation context for rule kinds that require formula
/// evaluation: `Expression` rules and `CellIs` rules with non-literal
/// `formula1` / `formula2`.
///
/// `anchor` is the top-left of the rule's owning sqref block — the
/// cell at which the formula source was authored. `target` is the
/// cell that the rule is currently being applied to. The shifter
/// rewrites every relative reference in the formula by
/// `(target.row - anchor.row, target.col - anchor.col)` so each cell
/// in the sqref sees a correctly-aimed formula.
///
/// `arena`, `registry`, and `eval_ctx` are forwarded to the formula
/// parser and tree-walk evaluator. `eval_ctx` is typically bound to
/// the same sheet (and workbook, for cross-sheet references) that the
/// CF block lives on; it must outlive every call.
///
/// `today_serial` pins the "today" reference for `TimePeriod` rules.
/// It is a date serial (Excel epoch, integer days since 1899-12-30
/// with the 1900 leap-year bug) — typically `floor(NOW())` taken once
/// at recalc start so all CF cells in a recalc see a consistent date.
/// `nullopt` means the host has not provided one; `TimePeriod` rules
/// then never match. Other rule kinds ignore this field.
struct CFEvalContext {
  CellAddress anchor{};
  CellAddress target{};
  Arena* arena = nullptr;
  const eval::FunctionRegistry* registry = nullptr;
  const eval::EvalContext* eval_ctx = nullptr;
  std::optional<double> today_serial;
};

/// Context-aware overload that handles every rule type the value-only
/// overload handles, plus:
///
///   * `Expression` rules — `formula1` is parsed, shifted, evaluated,
///     and tested for truthiness. A truthy scalar (boolean `true` or
///     non-zero number) matches; everything else (including text and
///     errors) does not.
///   * `CellIs` rules whose `formula1` / `formula2` are not bare
///     literals — the evaluated value is folded into the same
///     comparison machinery the literal-only overload uses.
///   * `TimePeriod` rules — the cell value (a date serial) is bucketed
///     against `ctx.today_serial`. Buckets follow Excel's Sunday-start
///     week convention. Cells that are not numeric do not match.
///
/// All other rule types delegate to the simple overload.
bool match_rule(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx);

}  // namespace cf
}  // namespace formulon

#endif  // FORMULON_CF_CF_EVALUATOR_H_
