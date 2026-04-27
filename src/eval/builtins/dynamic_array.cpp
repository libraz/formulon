// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's dynamic-array (spilling) built-ins.
//
// SEQUENCE is the only entry registered today; it is the simplest member of
// the family (no range-arg dependency, deterministic output) and therefore
// the canonical acceptance test for the cell-level spill pipeline:
// `=SEQUENCE(3)` typed into a formula cell exercises the entire chain
// (function-impl array production -> EvalContext::dispatch_array_result ->
// Sheet::commit_spill -> Sheet::resolve_cell_value for downstream readers).
// See `backup/plans/02-calc-engine.md` §2.6 for the wider dynamic-array
// design and `sheet.h` for the spill-table contract.

#include "eval/builtins/dynamic_array.h"

#include <cmath>
#include <cstddef>
#include <cstdint>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Per-call ceiling on `rows * cols` to prevent SEQUENCE from issuing a
// multi-gigabyte arena allocation when the user (or a fuzzer) asks for an
// enormous grid. Sheet::kMaxRows * Sheet::kMaxCols ~= 1.7e10, well above any
// reasonable spill footprint; cap at ~1M cells, the same order of magnitude
// as Mac Excel's effective dynamic-array ceiling for a single formula. Going
// over surfaces `#NUM!`, matching Excel's overflow code for SEQUENCE.
constexpr std::size_t kMaxSequenceCells = 1U << 20;  // 1,048,576

/// SEQUENCE(rows, [cols=1], [start=1], [step=1]).
///
/// Returns a `rows x cols` row-major numeric grid where the cell at
/// row-major index `i` (0-based) is `start + i * step`.
///
/// Coercion / error rules (Mac Excel 365, ja-JP):
///   * Each argument coerces via `coerce_to_number` (the dispatcher's
///     `propagate_errors=true` covers any propagated error before the impl
///     runs; this guard is for the explicit Text -> #VALUE! path).
///   * `rows` / `cols` are truncated toward zero. Either being `<= 0`
///     surfaces `#VALUE!`; either exceeding `Sheet::kMaxRows` /
///     `Sheet::kMaxCols`, or `rows*cols` exceeding `kMaxSequenceCells`,
///     surfaces `#NUM!`.
///   * `start` / `step` are passed through unchanged (any finite double).
///
/// Allocation contract: both the `ArrayValue` header and its row-major
/// `cells` buffer are arena-allocated, matching the `Value::Text` lifetime
/// rule (caller's arena must outlive the returned Value).
Value Sequence(const Value* args, std::uint32_t arity, Arena& arena) {
  auto rows_c = coerce_to_number(args[0]);
  if (!rows_c) {
    return Value::error(rows_c.error());
  }
  double cols_d = 1.0;
  if (arity >= 2U) {
    auto cols_c = coerce_to_number(args[1]);
    if (!cols_c) {
      return Value::error(cols_c.error());
    }
    cols_d = cols_c.value();
  }
  double start = 1.0;
  if (arity >= 3U) {
    auto start_c = coerce_to_number(args[2]);
    if (!start_c) {
      return Value::error(start_c.error());
    }
    start = start_c.value();
  }
  double step = 1.0;
  if (arity >= 4U) {
    auto step_c = coerce_to_number(args[3]);
    if (!step_c) {
      return Value::error(step_c.error());
    }
    step = step_c.value();
  }

  // Truncate-toward-zero is Mac Excel's behaviour for the row/col
  // dimensions; e.g. SEQUENCE(3.7) yields 3 rows, SEQUENCE(-0.5) yields
  // #VALUE! (truncates to 0, then the `<= 0` guard fires). NaN values fail
  // the `> 0` test by IEEE-754 rule and surface as #VALUE!.
  const double rows_t = std::trunc(rows_c.value());
  const double cols_t = std::trunc(cols_d);
  if (!(rows_t > 0.0) || !(cols_t > 0.0)) {
    return Value::error(ErrorCode::Value);
  }
  if (rows_t > static_cast<double>(Sheet::kMaxRows) || cols_t > static_cast<double>(Sheet::kMaxCols)) {
    return Value::error(ErrorCode::Num);
  }
  const auto rows = static_cast<std::uint32_t>(rows_t);
  const auto cols = static_cast<std::uint32_t>(cols_t);
  const std::size_t n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  if (n > kMaxSequenceCells) {
    return Value::error(ErrorCode::Num);
  }

  Value* buffer = arena.create_array<Value>(n);
  if (buffer == nullptr) {
    // Arena OOM -- preserve the spill pipeline's invariant that a function
    // either returns an Array with valid storage or an error. #NUM! is the
    // closest Excel-visible analogue for "result too large".
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < n; ++i) {
    buffer[i] = Value::number(start + static_cast<double>(i) * step);
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  arr->rows = rows;
  arr->cols = cols;
  arr->cells = buffer;
  return Value::array(arr);
}

}  // namespace

void register_dynamic_array_builtins(FunctionRegistry& registry) {
  // SEQUENCE: one required arg (rows) + three optional (cols, start, step).
  // Default `accepts_ranges=false` and `propagate_errors=true` -- a range
  // passed for any arg surfaces #VALUE! via the dispatcher's normal
  // scalar-coercion path, and any pre-evaluated error short-circuits before
  // the impl runs.
  registry.register_function(FunctionDef{"SEQUENCE", 1u, 4u, &Sequence});
}

}  // namespace eval
}  // namespace formulon
