// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared helpers for the CF evaluator: comparison primitives, the
// `ColorScalePopulation` cache shape, percentile interpolation, and the
// stateless date / weekday helpers used by `TimePeriod`.
//
// This header is internal to the `cf/` subsystem; callers outside the
// CF evaluator should keep using the public APIs declared in
// `cf/cf_evaluator.h`. The split exists so the rule-match and
// scale-evaluator TUs can share the same primitives without a single
// monolithic `cf_evaluator.cpp`.

#ifndef FORMULON_CF_CF_HELPERS_H_
#define FORMULON_CF_CF_HELPERS_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_evaluator.h"
#include "cf/cf_types.h"
#include "value.h"

namespace formulon {
class Sheet;
}

namespace formulon::cf {

/// Sorted ascending numeric population gathered from a sqref. The full
/// definition lives here so `CFEvalContext::cached_population` (declared
/// in `cf/cf_evaluator.h` via forward decl) carries a pointer with a
/// known layout to every TU inside the CF subsystem. External callers
/// only see the forward declaration in `cf/cf_evaluator.h`.
struct ColorScalePopulation {
  std::vector<double> sorted;  // ascending
  double min = 0.0;
  double max = 0.0;
};

namespace helpers {

// ---------------------------------------------------------------------------
// Literal-operand machinery for `cellIs` rules.
// ---------------------------------------------------------------------------

/// Small literal-only operand used by the `cellIs` evaluator. Owns its
/// text payload (unlike `Value::Text`, which is a non-owning view), so
/// it can survive parse helpers' stack frames without an arena.
struct LiteralOperand {
  enum class Kind : std::uint8_t { Number, Bool, Text };
  Kind kind = Kind::Number;
  double number_value = 0.0;
  bool bool_value = false;
  std::string text_value;
};

/// Parses a `cellIs` formula source as a literal. Excel emits cellIs
/// `formula1` / `formula2` as either a bare number (`10`, `-3.5`), a
/// boolean keyword (`TRUE` / `FALSE`), or a quoted string with `""`
/// double-quote escapes. Anything else (a reference, an arithmetic
/// expression) returns `nullopt`; the formula evaluator covers those
/// operand shapes via `parse_shift_evaluate`.
std::optional<LiteralOperand> parse_literal(const std::string& source);

/// Lifts a literal operand to a double. Text returns 0.0; callers must
/// check `kind` before reaching this branch.
double literal_as_number(const LiteralOperand& operand);

/// Three-way compare of `cell` against a literal operand. Returns
/// `nullopt` when the kinds are incompatible (e.g. text rule against a
/// numeric cell, any rule against an error cell).
std::optional<int> compare_cell_to_literal(const Value& cell, const LiteralOperand& operand);

/// Folds an evaluated `Value` into the literal-operand machinery so the
/// existing comparison helpers can stay the only path. Errors, blanks,
/// arrays, refs, and lambdas have no useful CellIs interpretation; they
/// produce `nullopt`.
std::optional<LiteralOperand> value_to_operand(const Value& evaluated);

/// Parses, shifts, and evaluates `source` as the body of a CF formula
/// anchored at `ctx.anchor` and applied at `ctx.target`. Parser errors,
/// arena exhaustion, and out-of-bounds shifts surface as Excel `#NAME?`
/// values so the caller can decide how to interpret them.
Value parse_shift_evaluate(const std::string& source, const CFEvalContext& ctx);

/// Returns the operand for a `cellIs` formula source, preferring the
/// literal-parser fast path (no parser instantiation, no arena traffic)
/// and falling back to the formula evaluator when the source isn't a
/// bare literal.
std::optional<LiteralOperand> cell_is_operand(const std::string& source, const CFEvalContext& ctx);

// ---------------------------------------------------------------------------
// ASCII case-insensitive string ops used by the text-rule family
// (containsText / beginsWith / endsWith) and the boolean-keyword parser.
// cellIs value ordering instead routes through `eval::compare_values`.
// ---------------------------------------------------------------------------

bool icase_equal(std::string_view lhs, std::string_view rhs);
bool icase_contains(std::string_view haystack, std::string_view needle);
bool icase_starts_with(std::string_view text, std::string_view prefix);
bool icase_ends_with(std::string_view text, std::string_view suffix);

// ---------------------------------------------------------------------------
// Population gathering / caching.
// ---------------------------------------------------------------------------

/// Cross-kind=false equality used for duplicate detection. Numbers
/// compare by exact IEEE-754 equality. Text uses ASCII case-insensitive
/// equality. Booleans compare by identity. Errors and blanks never
/// participate.
bool cf_values_equal(const Value& lhs, const Value& rhs);

/// Counts how many cells inside `sqref` (read spill-aware via
/// `Sheet::resolve_cell_value`) carry a value equal to `target` under
/// `cf_values_equal`.
std::size_t count_matches_in_sqref(const Value& target, const std::vector<CFCellRange>& sqref, const Sheet& sheet);

/// Collects every numeric value in `sqref`. Booleans and text are
/// skipped — Excel's AboveAverage / Top10 / scale-family rules use the
/// numeric-only population.
std::vector<double> collect_numeric_values(const std::vector<CFCellRange>& sqref, const Sheet& sheet);

double mean_of(const std::vector<double>& values);

/// Sample standard deviation (Bessel-corrected, n-1). Returns 0 when
/// fewer than 2 numeric values are present so the threshold collapses to
/// the mean and the rule still produces a well-defined match.
double sample_stddev(const std::vector<double>& values, double mean);

/// Gathers the numeric population for `sqref` and sorts it ascending.
ColorScalePopulation gather_population(const std::vector<CFCellRange>& sqref, const Sheet& sheet);

/// Returns the cached population from `ctx` when present; otherwise
/// gathers from the sheet into `fallback` and returns it. The returned
/// reference is valid for as long as `fallback` (or the cache pointer)
/// outlives the call site. Callers stack-allocate `fallback`.
const ColorScalePopulation& ensure_population(const CFEvalContext& ctx, ColorScalePopulation& fallback);

/// Returns a pointer to the numeric population usable for `cell_value`.
/// Returns `nullptr` when the context lacks a sqref / sheet, when the
/// cell is not numeric, or when the population is empty.
const ColorScalePopulation* numeric_population_for_cell(const Value& cell_value, const CFEvalContext& ctx,
                                                        ColorScalePopulation& fallback);

// ---------------------------------------------------------------------------
// Numeric utilities consumed by the scale evaluators and `Top10` rank
// resolution.
// ---------------------------------------------------------------------------

/// `PERCENTILE.INC`: linear interpolation between sorted points.
/// `percentile` is in `[0, 1]`. For empty populations the caller
/// short-circuits.
double percentile_inc(const std::vector<double>& sorted, double percentile);

/// Parses a textual double (full match required, no trailing junk).
/// Used by the CFVO threshold resolver and `parse_literal`.
std::optional<double> parse_double(std::string_view source);

// ---------------------------------------------------------------------------
// Date helpers used by `TimePeriod`.
//
// These mirror Excel's `WEEKDAY(_, 1)` (Sunday = 1) semantics and the
// month-arithmetic used by the `LastMonth` / `NextMonth` buckets. They
// live here for now; a later wave consolidates them into
// `eval/date_time.h` (D-08).
// ---------------------------------------------------------------------------

/// Excel weekday with `WEEKDAY(date, 1)` semantics: Sunday = 1,
/// Saturday = 7.
int weekday_sunday_one(double serial_floor);

/// Serial of the Sunday opening the week that contains `serial_floor`.
double sunday_of_week(double serial_floor);

struct YearMonth {
  int year;
  unsigned month;  // 1..12

  friend bool operator==(YearMonth lhs, YearMonth rhs) noexcept {
    return lhs.year == rhs.year && lhs.month == rhs.month;
  }
};

YearMonth year_month_from_serial(double serial_floor);

/// Shifts `anchor` by `delta_months`, normalising the result back into
/// the canonical `1..12` month range.
YearMonth shift_year_month(YearMonth anchor, int delta_months);

}  // namespace helpers
}  // namespace formulon::cf

#endif  // FORMULON_CF_CF_HELPERS_H_
