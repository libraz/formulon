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

#include <cstddef>
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

/// Maximum cells one rectangular reference may materialise as individual
/// dependency-graph edges.
///
/// This is a *graph-footprint* bound and is deliberately not
/// `kMaxRangeExpansionCells`: that ceiling sizes a transient `vector<Value>`
/// released as soon as one evaluation finishes, whereas every edge admitted
/// here is retained for the lifetime of the formula and is carried three
/// times over (forward adjacency, reverse adjacency, and the source index),
/// so its per-cell cost is ~100-160 bytes of permanently resident memory.
/// Rectangles above the ceiling are registered as a single compact rectangle
/// dependency instead, which the recalc engine expands lazily.
///
/// 1024 keeps per-cell edges — and with them the exact evaluation ordering
/// they give against formulas inside the rectangle — for the hand-authored
/// aggregates that dominate real workbooks, while capping one formula's
/// permanent graph footprint at ~160 KB.
inline constexpr std::uint64_t kMaxMaterializedDependencyCells = 1024U;

/// Maximum span slots one REGEX* call may accumulate, counted as
/// `matches * (capture_count + 1)`. PCRE2's own `match_limit` bounds the
/// work spent inside a single match attempt but says nothing about how many
/// matches a find-all scan retains, and each retained match carries one slot
/// per capture group. Sized to `kMaxDynamicArrayCells` because
/// `REGEXEXTRACT(..., 3)` projects exactly that product into an array
/// result, so accumulating more could never produce a usable value.
inline constexpr std::uint64_t kMaxRegexMatchSlots = kMaxDynamicArrayCells;

/// Maximum physical page count one pagination request may report.
///
/// A manual page break may sit before every track of the grid, so the finest
/// page grid a file can declare is one page per row and per column:
/// `Sheet::kMaxRows * Sheet::kMaxCols` is 2^34 pages, past what the reported
/// 32-bit count can hold, and the print area list multiplies that again. The
/// ceiling is therefore a policy bound rather than a grid-derived one. 2^24
/// sits three orders of magnitude above the tallest job the grid can actually
/// print — the full 1,048,576-row sheet at the default row height paginates
/// to roughly 21,800 pages — so it rejects only break configurations whose
/// count could never be reported faithfully.
inline constexpr std::uint64_t kMaxPaginationPages = 16777216U;  // 2^24

/// Ceiling on the arena backing one evaluation, in bytes.
///
/// The evaluator's arena is reset between cells and every allocation site
/// already reports `Arena::exhausted()`, so a ceiling here turns unbounded
/// growth on a hostile formula into a `#NUM!` for that one cell instead of
/// an out-of-memory abort that takes the whole process (or, on WASM, the
/// host page) down. Sized at four times the largest legitimate single
/// result — `kMaxRangeExpansionCells` cells at 24 bytes each is ~240 MB —
/// so the intermediates of a genuinely large evaluation still fit.
inline constexpr std::size_t kMaxEvalArenaBytes = 1024U * 1024U * 1024U;  // 1 GiB

/// Ceiling on an arena backing one parse during a file load, in bytes.
///
/// These arenas hold the AST of a single formula read from the file, which
/// is orders of magnitude smaller than this; the ceiling exists so a
/// crafted formula string cannot turn a load into an allocation loop.
inline constexpr std::size_t kMaxLoadArenaBytes = 64U * 1024U * 1024U;  // 64 MiB

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
