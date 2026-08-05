//
// `ResourceBudget` is a small request-scoped guard against unbounded work
// driven by attacker-controlled input (cells scanned, output bytes,
// intermediate allocations). A call site constructs a budget with a ceiling,
// charges work units against it via `consume()`, and stops as soon as the
// ceiling would be exceeded.
//
// Contract:
//   * The budget is a plain value type with no global state; its lifetime is
//     the request (one evaluation, one recalc pass, one pivot build).
//   * `consume(n)` either charges `n` units and succeeds, or leaves the
//     consumed count untouched and returns an `Error` carrying the code the
//     budget was constructed with (default `kSecResourceLimit`).
//   * `would_exceed(n)` is the side-effect-free variant for call sites that
//     no-op instead of propagating an error.
//   * Arithmetic is `uint64_t` and overflow-safe: a request larger than the
//     remaining headroom fails instead of wrapping.
//
// The header is dependency-free beyond `error.h` / `expected.h` and is safe
// to include from WASM-targeted translation units.

#ifndef FORMULON_UTILS_RESOURCE_BUDGET_H_
#define FORMULON_UTILS_RESOURCE_BUDGET_H_

#include <cstdint>
#include <string>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

/// Default ceiling for a dense pivot result matrix, in cells
/// (`row_leaf_count * col_leaf_count`). A real pivot body is orders of
/// magnitude smaller; past this point the dense bucket/value matrices would
/// commit hundreds of megabytes before any aggregation runs.
inline constexpr std::uint64_t kMaxPivotResultCells = 4194304U;  // 2^22

/// Default ceiling for partial-recalc viewport seeding, in coordinates.
/// Sized to one full Excel column (`Sheet::kMaxRows`), which comfortably
/// covers any UI redraw region while rejecting near-full-grid rectangles
/// (~17e9 coordinates).
inline constexpr std::uint64_t kMaxRecalcViewportCells = 1048576U;  // 2^20

/// Maximum dense dynamic-array result size. This is also the historic
/// SEQUENCE ceiling; applying it at the shared allocation seam prevents
/// other array constructors from bypassing the same resource bound.
inline constexpr std::uint64_t kMaxDynamicArrayCells = 1048576U;  // 2^20

/// Maximum cells eagerly materialised from one rectangular reference.
inline constexpr std::uint64_t kMaxRangeExpansionCells = 10'000'000U;

/// Running work-unit counter checked against a fixed ceiling.
///
/// Typical use:
///
/// ```cpp
/// ResourceBudget cells(kMaxPivotResultCells, FormulonErrorCode::kFnOverflow);
/// auto ok = cells.consume(rows * cols);
/// if (!ok) {
///   return ok.error();
/// }
/// ```
class ResourceBudget {
 public:
  /// Builds a budget with the given ceiling. `code` selects the error code
  /// surfaced when the ceiling is exceeded, so call sites can keep their
  /// module's established code (e.g. `kFnOverflow` for function-layer
  /// overflows).
  explicit constexpr ResourceBudget(std::uint64_t ceiling,
                                    FormulonErrorCode code = FormulonErrorCode::kSecResourceLimit) noexcept
      : ceiling_(ceiling), code_(code) {}

  /// True when charging `units` more work would exceed the ceiling.
  constexpr bool would_exceed(std::uint64_t units) const noexcept { return units > ceiling_ - used_; }

  /// Charges `units` of work against the budget. On success the consumed
  /// count advances; on failure it is left untouched and the returned error
  /// carries the budget's error code plus a `used/requested/ceiling`
  /// context string.
  Expected<void, Error> consume(std::uint64_t units) {
    if (would_exceed(units)) {
      return make_error(code_, "resource budget exceeded",
                        "used=" + std::to_string(used_) + " requested=" + std::to_string(units) +
                            " ceiling=" + std::to_string(ceiling_));
    }
    used_ += units;
    return {};
  }

  /// Work units consumed so far.
  constexpr std::uint64_t used() const noexcept { return used_; }

  /// Configured ceiling.
  constexpr std::uint64_t ceiling() const noexcept { return ceiling_; }

  /// Headroom left before the ceiling is hit.
  constexpr std::uint64_t remaining() const noexcept { return ceiling_ - used_; }

 private:
  std::uint64_t ceiling_;
  std::uint64_t used_ = 0;
  FormulonErrorCode code_;
};

}  // namespace formulon

#endif  // FORMULON_UTILS_RESOURCE_BUDGET_H_
