// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
//   * `DuplicateValues` / `UniqueValues` — count occurrences of the
//     cell's value across `CFEvalContext::sqref` and match when the
//     count is >= 2 (Duplicate) or == 1 (Unique). Cross-kind equality
//     is `false`; errors and blanks never match.
//   * `AboveAverage` — collect numeric values from
//     `CFEvalContext::sqref`, compute the sample mean (and sample
//     standard deviation when `rule.std_dev` is set), then test the
//     cell's number against the resulting threshold honouring
//     `rule.above_average` (top vs. bottom side) and
//     `rule.equal_average` (strict vs. inclusive). Non-numeric cells
//     do not match.
//   * `Top10` — collect numeric values from `CFEvalContext::sqref`,
//     pick the rank-th largest (or smallest when `rule.bottom`) value
//     as a threshold, and match cells with value >= threshold (top)
//     or <= threshold (bottom) so ties are included. `rule.percent`
//     interprets `rule.rank` as a percent of the population count.
//   * `ColorScale` — resolves each `<cfvo>` threshold against the
//     `CFEvalContext::sqref` population (Number / Percent / Percentile
//     / Min / Max / AutoMin / AutoMax / Formula), then linearly
//     interpolates between the bounding stop colours in RGB space.
//     The context-aware `make_match` overload produces a `CFMatch`
//     with `resolved_fill_color` engaged; the boolean overload is
//     consumed only as a "rule applies" gate.
//   * `DataBar` — resolves `rule.data_bar->min` / `max` against the
//     same CFVO machinery, computes the bar length as a percentage of
//     `(max - min)` scaled into `[min_length_pct, max_length_pct]`,
//     and selects the axis position. `Automatic` axis splits at the
//     proportional negative offset; `Middle` pins to 50; `None` pins
//     to 0. `is_negative` is set when the cell value is < 0.
//   * `IconSet` — resolves each `<cfvo>` threshold against the
//     population, walks them in ascending order honouring each CFVO's
//     `gte` flag (`>=` vs. `>`), and assigns an `icon_index` between
//     `0` and `N-1` for an N-icon set. `rule.icon_set->reverse` flips
//     the index. The resolved `IconRender` carries the icon set name
//     so the host can look up the glyph.
//
// All visual rule kinds are now covered.
//
// Cross-block evaluation:
//
//   * `evaluate_cf_at(sheet, target, host)` — walks every
//     `<conditionalFormatting>` block on `sheet` whose sqref contains
//     `target`, collects each rule, sorts by workbook-global priority
//     (ascending), and consults `match_rule` for each in order. Every
//     positive match yields a `CFMatch`; `stop_if_true` halts the
//     walk early. The result list is priority-ascending.
//   * `evaluate_cf_for_range(sheet, range, host)` — runs the same
//     walk for every cell in `range` and returns one
//     `CFRangeCellMatches` entry per cell that produced at least one
//     match. Sparse-by-default: cells outside any block's sqref do
//     not appear in the result.
//
// Each step extends `match_rule` and the public `evaluate_*` helpers
// without changing the existing call signatures.
//
// Design references:
//   * src/cf/cf_match.h (return-type contract)

#ifndef FORMULON_CF_CF_EVALUATOR_H_
#define FORMULON_CF_CF_EVALUATOR_H_

#include <optional>
#include <vector>

#include "cell.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "value.h"

namespace formulon {

class Arena;
class Sheet;

namespace eval {
class EvalContext;
class FunctionRegistry;
}  // namespace eval

namespace cf {

/// Sorted ascending numeric population of a sqref. Defined in
/// `cf_evaluator.cpp`; only forward-declared here so `CFEvalContext` can
/// carry an opaque pointer that the viewport API uses to share a
/// per-block population across cells. Callers outside the cf module
/// neither see the layout nor populate the field — they pass `nullptr`
/// (the default), and the evaluator gathers populations on demand.
struct ColorScalePopulation;

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
/// match into the result list. The value-only overload covers
/// `DifferentialFormat` rules; visual kinds (ColorScale / DataBar /
/// IconSet) need the cell value and `CFEvalContext` to resolve their
/// render payloads — use the context-aware overload below.
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
///
/// `sqref` points to the rule's owning sqref (the union of cell ranges
/// the rule applies to). Range-aware rules — `DuplicateValues`,
/// `UniqueValues`, and forthcoming `Top10` / `AboveAverage` — iterate
/// this list to compute population statistics over all values in the
/// sqref. `nullptr` means the host has not provided one; range-aware
/// rules then never match. Other rule kinds ignore this field. The
/// pointee is borrowed and must outlive the `match_rule` call.
///
/// `cached_population` is an optimisation hint for the viewport API.
/// When set, range-aware rule kinds (ColorScale, DataBar, IconSet,
/// AboveAverage, Top10) read the sqref's numeric population from this
/// pointer instead of re-walking the sheet. The viewport walker
/// computes one population per `<conditionalFormatting>` block at the
/// top of `evaluate_cf_for_range` and reuses it across every cell in
/// the range; per-cell `evaluate_cf_at` callers leave this `nullptr`
/// and the helpers compute on demand.
struct CFEvalContext {
  CellAddress anchor{};
  CellAddress target{};
  Arena* arena = nullptr;
  const eval::FunctionRegistry* registry = nullptr;
  const eval::EvalContext* eval_ctx = nullptr;
  std::optional<double> today_serial;
  const std::vector<CFCellRange>* sqref = nullptr;
  const ColorScalePopulation* cached_population = nullptr;
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
///   * `DuplicateValues` / `UniqueValues` — the evaluator counts how
///     many cells in `ctx.sqref` carry a value equal to `cell_value`
///     under cross-kind=false equality (numbers compared by exact
///     IEEE-754 equality, text by ASCII case-insensitive equality,
///     booleans by identity; errors and blanks never match anything).
///     `DuplicateValues` matches when the count is 2 or more,
///     `UniqueValues` when the count is exactly 1. Cells whose value
///     is not number / boolean / text never match.
///
/// All other rule types delegate to the simple overload.
bool match_rule(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx);

/// Stable inputs for the `evaluate_cf_at` cross-block walker. Every
/// rule on a sheet shares the same arena, function registry,
/// evaluation context, and `today_serial` pin; the walker provides
/// the per-rule anchor / target / sqref itself by reading the
/// owning `<conditionalFormatting>` block. All pointee references
/// must outlive the `evaluate_cf_at` call.
struct CFHost {
  Arena* arena = nullptr;
  const eval::FunctionRegistry* registry = nullptr;
  const eval::EvalContext* eval_ctx = nullptr;
  std::optional<double> today_serial;
};

/// Walks every `<conditionalFormatting>` block on `sheet` whose sqref
/// contains `target`, collects each rule, sorts by workbook-global
/// `priority` (ascending), and calls `match_rule(rule, cell_value,
/// ctx)` on each in order. Every positive match yields a `CFMatch`
/// produced by the context-aware `make_match`. When a matched rule's
/// `stop_if_true` is set, the walk halts immediately and later rules
/// (across the same or sibling blocks) are not consulted.
///
/// The returned list is priority-ascending; consumers fold matches in
/// order, so a higher-priority `DifferentialFormat` overlays a
/// lower-priority one. Visual rules render alongside dxf-driven ones.
///
/// `host.arena` / `host.registry` / `host.eval_ctx` are required;
/// `host.today_serial` is forwarded to `TimePeriod` rules and may be
/// `nullopt`.
std::vector<CFMatch> evaluate_cf_at(const Sheet& sheet, CellAddress target, const CFHost& host);

/// One cell's CF result inside a viewport-range evaluation. `cell` is
/// the cell address; `matches` is the priority-ascending list returned
/// by `evaluate_cf_at` for that cell.
struct CFRangeCellMatches {
  CellAddress cell{};
  std::vector<CFMatch> matches;
};

/// Walks every cell in `range` (inclusive on both corners) and returns
/// one `CFRangeCellMatches` entry for each cell that produced at least
/// one match. The result is sparse: cells outside any
/// `<conditionalFormatting>` block's sqref, and cells whose rules all
/// failed to match, do not appear. Order is row-major over `range`.
///
/// Equivalent to calling `evaluate_cf_at` once per cell in `range` and
/// dropping cells with empty match lists, but the implementation
/// computes each block's numeric population once and reuses it across
/// every cell in `range`. Range-aware rule kinds (ColorScale, DataBar,
/// IconSet, AboveAverage, Top10) consume the cached population through
/// `CFEvalContext::cached_population`; cells inside the same block see
/// identical statistics regardless of where they sit in the viewport.
std::vector<CFRangeCellMatches> evaluate_cf_for_range(const Sheet& sheet, CFCellRange range, const CFHost& host);

/// Context-aware `make_match` overload. Identical to the value-only
/// overload for `DifferentialFormat` rules; for visual rules it
/// computes the resolved render payload:
///
///   * `ColorScale` — resolves `rule.color_scale` thresholds against
///     the sqref population, locates the segment containing the cell
///     value, and linearly interpolates the bounding stop colours in
///     RGB space. Engages `CFMatch::resolved_fill_color`.
///   * `DataBar` — resolves the `min` / `max` CFVOs, computes the bar
///     length as `(cell - min) / (max - min)` clamped to `[0, 1]` and
///     scaled into `[min_length_pct, max_length_pct]`, and assigns
///     `axis_position_pct` (`Automatic` / `Middle` / `None`). Engages
///     `CFMatch::data_bar_render`.
///   * `IconSet` — resolves each threshold CFVO, walks them in
///     ascending order, and assigns the cell's bucket index honouring
///     each CFVO's `gte` flag and the spec's `reverse` flag. Engages
///     `CFMatch::icon_render`.
///
/// Callers should still gate on the boolean `match_rule` overload to
/// decide whether to surface the resulting `CFMatch`.
CFMatch make_match(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx);

}  // namespace cf
}  // namespace formulon

#endif  // FORMULON_CF_CF_EVALUATOR_H_
