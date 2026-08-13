//
// Implementation of Formulon's dynamic-array (spilling) built-ins.
//
// SEQUENCE is registered here; it is the simplest member of the family
// (no range-arg dependency, deterministic output) and the canonical
// acceptance test for the cell-level spill pipeline: `=SEQUENCE(3)` typed
// into a formula cell exercises the entire chain (function-impl array
// production -> EvalContext::dispatch_array_result -> Sheet::commit_spill
// -> Sheet::resolve_cell_value for downstream readers).
//
// TRANSPOSE conceptually belongs to this family but lives in
// `eval/shape_ops_lazy.cpp` because it requires per-argument AST shape
// inspection (a Range / RangeOp argument must keep its 2D shape; the
// eager dispatcher would flatten it to a row-major scalar vector). Lookups
// for TRANSPOSE go through the central `kLazyDispatch[]` table in
// `tree_walker.cpp`.
//
// See `sheet.h` for the spill-table contract.

#include "eval/builtins/dynamic_array.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <random>

#include "eval/array_alloc.h"
#include "eval/builtins/numeric_helpers.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/dynamic_array_limits.h"
#include "eval/function_registry.h"
#include "eval/rng.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/expected.h"
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
// SEQUENCE / RANDARRAY use the lenient pass-through variant: NaN / Inf are
// not rejected here because the impls run their own `> 0` / `<=` shape
// checks downstream which reject NaN by IEEE-754 rule.
inline Expected<double, ErrorCode> read_optional_number_arg(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                            double default_value) {
  return builtins_detail::read_optional_number(args, arity, index, default_value, /*check_finite=*/false);
}

Expected<bool, ErrorCode> read_optional_bool_arg(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                 bool default_value) {
  if (arity <= index) {
    return default_value;
  }
  return coerce_to_bool(args[index]);
}

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
  auto cols_c = read_optional_number_arg(args, arity, 1U, 1.0);
  if (!cols_c) {
    return Value::error(cols_c.error());
  }
  auto start_c = read_optional_number_arg(args, arity, 2U, 1.0);
  if (!start_c) {
    return Value::error(start_c.error());
  }
  auto step_c = read_optional_number_arg(args, arity, 3U, 1.0);
  if (!step_c) {
    return Value::error(step_c.error());
  }
  const double cols_d = cols_c.value();
  const double start = start_c.value();
  const double step = step_c.value();

  // Truncate-toward-zero is Excel's behaviour for the row/col dimensions.
  // A value that truncates to zero yields #CALC!, while a strictly negative
  // dimension yields #VALUE!. This distinction is observable for
  // SEQUENCE(0) versus SEQUENCE(-1).
  const double rows_t = std::trunc(rows_c.value());
  const double cols_t = std::trunc(cols_d);
  if (rows_t == 0.0 || cols_t == 0.0) {
    return Value::error(ErrorCode::Calc);
  }
  if (!(rows_t > 0.0) || !(cols_t > 0.0)) {
    return Value::error(ErrorCode::Value);
  }
  if (rows_t > static_cast<double>(Sheet::kMaxRows) || cols_t > static_cast<double>(Sheet::kMaxCols)) {
    return Value::error(ErrorCode::Num);
  }
  const auto rows = static_cast<std::uint32_t>(rows_t);
  const auto cols = static_cast<std::uint32_t>(cols_t);
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buffer, kMaxSequenceCells);
  if (arr == nullptr) {
    // Rejected shape or arena OOM -- preserve the spill pipeline's invariant
    // that a function either returns an Array with valid storage or an
    // error. #NUM! is the closest Excel-visible analogue for "result too
    // large".
    return Value::error(ErrorCode::Num);
  }
  const std::size_t n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  for (std::size_t i = 0; i < n; ++i) {
    buffer[i] = Value::number(start + static_cast<double>(i) * step);
  }
  return Value::array(arr);
}

/// RANDARRAY([rows], [cols], [min], [max], [whole_number]).
///
/// Returns a `rows x cols` row-major array of independent uniform samples.
/// Defaults: `rows = 1`, `cols = 1`, `min = 0`, `max = 1`,
/// `whole_number = FALSE` (matches Mac Excel 365).
///
/// Distribution:
///   * `whole_number = FALSE`: each cell is a uniform real in
///     `[min, max)` (open at the upper bound, matching
///     `std::uniform_real_distribution`).
///   * `whole_number = TRUE`: each cell is a uniform integer in the
///     closed interval `[min, max]`. Both bounds must already be
///     integers; `min` or `max` with a non-zero fractional part surfaces
///     `#VALUE!` per Microsoft's documented behaviour.
///
/// Errors:
///   * Any argument failing `coerce_to_number` propagates.
///   * `rows <= 0` or `cols <= 0` -> `#VALUE!`.
///   * `rows`, `cols`, or `rows * cols` exceeding the same ceilings used
///     by SEQUENCE -> `#NUM!`.
///   * `min > max` -> `#VALUE!`.
///   * `whole_number = TRUE` with a non-integer bound -> `#VALUE!`.
///
/// RNG: shared per-thread Mersenne Twister via `thread_local_rng()` (same
/// stream as RAND / RANDBETWEEN).
Value RandArray(const Value* args, std::uint32_t arity, Arena& arena) {
  auto rows_c = read_optional_number_arg(args, arity, 0U, 1.0);
  if (!rows_c) {
    return Value::error(rows_c.error());
  }
  auto cols_c = read_optional_number_arg(args, arity, 1U, 1.0);
  if (!cols_c) {
    return Value::error(cols_c.error());
  }
  auto min_c = read_optional_number_arg(args, arity, 2U, 0.0);
  if (!min_c) {
    return Value::error(min_c.error());
  }
  auto max_c = read_optional_number_arg(args, arity, 3U, 1.0);
  if (!max_c) {
    return Value::error(max_c.error());
  }
  auto whole_number_c = read_optional_bool_arg(args, arity, 4U, false);
  if (!whole_number_c) {
    return Value::error(whole_number_c.error());
  }
  const double rows_d = rows_c.value();
  const double cols_d = cols_c.value();
  const double min_v = min_c.value();
  const double max_v = max_c.value();
  const bool whole_number = whole_number_c.value();

  // Shape validation. SEQUENCE precedent: truncate-toward-zero, reject
  // <= 0 with #VALUE!, reject oversize with #NUM!.
  const double rows_t = std::trunc(rows_d);
  const double cols_t = std::trunc(cols_d);
  if (!(rows_t > 0.0) || !(cols_t > 0.0)) {
    return Value::error(ErrorCode::Value);
  }
  if (rows_t > static_cast<double>(Sheet::kMaxRows) || cols_t > static_cast<double>(Sheet::kMaxCols)) {
    return Value::error(ErrorCode::Num);
  }
  const auto rows = static_cast<std::uint32_t>(rows_t);
  const auto cols = static_cast<std::uint32_t>(cols_t);
  const auto n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);

  // Bounds validation.
  if (std::isnan(min_v) || std::isnan(max_v) || std::isinf(min_v) || std::isinf(max_v)) {
    return Value::error(ErrorCode::Num);
  }
  if (min_v > max_v) {
    return Value::error(ErrorCode::Value);
  }
  if (whole_number) {
    if (std::trunc(min_v) != min_v || std::trunc(max_v) != max_v) {
      return Value::error(ErrorCode::Value);
    }
  }

  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buffer, kMaxSequenceCells);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  std::mt19937_64& rng = thread_local_rng();
  if (whole_number) {
    const auto lo = static_cast<std::int64_t>(min_v);
    const auto hi = static_cast<std::int64_t>(max_v);
    std::uniform_int_distribution<std::int64_t> dist(lo, hi);
    for (std::size_t i = 0; i < n; ++i) {
      buffer[i] = Value::number(static_cast<double>(dist(rng)));
    }
  } else {
    // `std::uniform_real_distribution` matches Excel's documented
    // half-open `[min, max)` for non-whole RANDARRAY. When `min == max`
    // the distribution returns `min` deterministically.
    if (min_v == max_v) {
      for (std::size_t i = 0; i < n; ++i) {
        buffer[i] = Value::number(min_v);
      }
    } else {
      std::uniform_real_distribution<double> dist(min_v, max_v);
      for (std::size_t i = 0; i < n; ++i) {
        buffer[i] = Value::number(dist(rng));
      }
    }
  }

  return Value::array(arr);
}

/// MUNIT(n). Returns the `n x n` identity matrix as an `ArrayValue`.
///
/// Coercion / error rules (Mac Excel 365, ja-JP):
///   * `n` coerces via `coerce_to_number`.
///   * `n` is truncated toward zero. `n <= 0` surfaces `#VALUE!` (matches
///     Mac Excel's per-call "non-positive dimension" surface).
///   * `n` exceeding `Sheet::kMaxRows` / `Sheet::kMaxCols`, or `n*n`
///     exceeding `kMaxSequenceCells`, surfaces `#NUM!` (the same ceiling
///     as SEQUENCE / RANDARRAY -- a hand-tuned 1M-cell limit that
///     comfortably fits any realistic identity matrix).
///
/// Allocation: the cells buffer is arena-allocated and zero-initialised
/// via `Value::number(0.0)` per slot, with the diagonal overwritten by
/// `Value::number(1.0)`. Both the buffer and the `ArrayValue` header live
/// in `arena`, matching the SEQUENCE / RANDARRAY contract.
Value MUnit(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto n_c = coerce_to_number(args[0]);
  if (!n_c) {
    return Value::error(n_c.error());
  }
  // Truncate-toward-zero: MUNIT(3.7) yields a 3x3 matrix (matches Excel).
  const double n_t = std::trunc(n_c.value());
  if (!(n_t > 0.0)) {
    return Value::error(ErrorCode::Value);
  }
  if (n_t > static_cast<double>(Sheet::kMaxRows) || n_t > static_cast<double>(Sheet::kMaxCols)) {
    return Value::error(ErrorCode::Num);
  }
  const auto n = static_cast<std::uint32_t>(n_t);
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(n, n, arena, buffer, kMaxSequenceCells);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t total = static_cast<std::size_t>(n) * static_cast<std::size_t>(n);
  for (std::size_t i = 0; i < total; ++i) {
    buffer[i] = Value::number(0.0);
  }
  for (std::uint32_t i = 0; i < n; ++i) {
    buffer[static_cast<std::size_t>(i) * n + i] = Value::number(1.0);
  }
  return Value::array(arr);
}

}  // namespace

void register_dynamic_array_builtins(FunctionRegistry& registry) {
  // SEQUENCE: one required arg (rows) + three optional (cols, start, step).
  // Default `accepts_ranges=false` and `propagate_errors=true` -- a range
  // passed for any arg surfaces #VALUE! via the dispatcher's normal
  // scalar-coercion path, and any pre-evaluated error short-circuits before
  // the impl runs.
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"SEQUENCE", 1u, 4u, &Sequence, true, false, false, false, false, FunctionDef::BlankScalarPolicy::Allow,
       ErrorCode::Value, FunctionDef::ResultShape::kArray},
      // RANDARRAY: zero required + five optional (rows, cols, min, max,
      // whole_number). Same dispatcher policy as SEQUENCE.
      {"RANDARRAY", 0u, 5u, &RandArray, true, false, false, false, false, FunctionDef::BlankScalarPolicy::Allow,
       ErrorCode::Value, FunctionDef::ResultShape::kArray},
      // MUNIT: single required arg (matrix size).
      {"MUNIT", 1u, 1u, &MUnit, true, false, false, false, false, FunctionDef::BlankScalarPolicy::Allow,
       ErrorCode::Value, FunctionDef::ResultShape::kArray},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
