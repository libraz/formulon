// Copyright 2026 libraz. Licensed under the MIT License.
//
// Conditional-format rule evaluator. Drives a single `cf::CFRule`
// against a cell's `Value` and returns whether the rule matches.
//
// Current coverage:
//
//   * `ContainsBlanks` / `NotContainsBlanks`
//   * `ContainsErrors` / `NotContainsErrors`
//   * `CellIs` against a literal `formula1` / `formula2` (number,
//     boolean, or quoted text). Operands that contain a reference or
//     arithmetic expression land with PR8.
//
// Subsequent PRs layer in:
//
//   * `Expression` rules with relative-reference shift (PR8) — also
//     extends `cellIs` to evaluate non-literal `formula1` / `formula2`
//   * `ContainsText` family (PR9)
//   * `Top10` / `AboveAverage` / `TimePeriod` (PR10)
//   * `ColorScale` / `DataBar` / `IconSet` visual computation (PR11)
//   * Priority chain + stopIfTrue + lazy viewport API (PR12 / PR13)
//
// Each PR extends `match_rule` and the public `evaluate_*` helpers
// without changing the existing call signatures.
//
// Design references:
//   * backup/plans/20-conditional-format-deep.md §20.4 / §20.5
//   * src/cf/cf_match.h (return-type contract)

#ifndef FORMULON_CF_CF_EVALUATOR_H_
#define FORMULON_CF_CF_EVALUATOR_H_

#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "value.h"

namespace formulon::cf {

/// Returns `true` when `rule` matches the cell carrying `cell_value`.
///
/// Implemented coverage:
///   * `ContainsBlanks`     — `cell_value.is_blank()`
///   * `NotContainsBlanks`  — `!cell_value.is_blank()`
///   * `ContainsErrors`     — `cell_value.is_error()`
///   * `NotContainsErrors`  — `!cell_value.is_error()`
///   * `CellIs`             — compares `cell_value` against the literal
///                            parsed from `rule.formula1` (and
///                            `rule.formula2` for `Between` /
///                            `NotBetween`) using `rule.op`. Operands
///                            that aren't a literal number, boolean,
///                            or quoted string fall through to `false`
///                            until PR8 wires the formula evaluator.
///
/// All other rule types currently return `false` (their evaluation
/// logic lands in subsequent PRs). Callers that want to opt out of
/// pre-implementation false-negatives should gate on the rule's
/// `type` and skip rules that aren't yet supported.
bool match_rule(const CFRule& rule, const Value& cell_value);

/// Returns the `CFMatch` payload for a rule that has been determined
/// to match. Used by the integration helpers to convert a positive
/// match into the result list. For PR6 the only match kind is
/// `DifferentialFormat`; later PRs introduce the visual kinds.
CFMatch make_match(const CFRule& rule);

}  // namespace formulon::cf

#endif  // FORMULON_CF_CF_EVALUATOR_H_
