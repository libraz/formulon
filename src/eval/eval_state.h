// Copyright 2026 libraz. Licensed under the MIT License.
//
// `EvalState` is the mutable per-`evaluate()` state threaded through
// `EvalContext` to support recursive evaluation of formula cells. It carries
// two pieces of transient information:
//
//   * An in-progress stack of `(sheet, row, col)` cell addresses currently
//     being evaluated. A push that would duplicate a key on the stack is a
//     cycle; the caller surfaces that as `#REF!`. The sheet dimension is
//     required so that a cross-sheet cycle (`Sheet1!A1 -> Sheet2!A1 ->
//     Sheet1!A1`) is caught without false-positives on same-coordinate
//     cells living in different sheets.
//   * A memoisation map from `(sheet, row, col)` to the computed `Value`
//     for a formula cell. Memoised results are reused within a single
//     top-level `evaluate()` invocation so that diamond dependencies (two
//     paths into the same cell) do not re-parse and re-evaluate the same
//     formula twice.
//
// `EvalState` is intentionally not persisted across calls: iterative calc,
// SCC pre-detection, and a persistent dependency graph live in a later
// phase (see `backup/plans/02-calc-engine.md` §2.7.3). The Sheet is NOT
// mutated by the evaluator — formula results only live inside this struct
// for the duration of the surrounding `evaluate()` call.
//
// Memoised Text values store a `std::string_view` whose storage is the
// caller's evaluation arena; the caller must keep that arena alive as long
// as the `EvalState` itself is accessed.

#ifndef FORMULON_EVAL_EVAL_STATE_H_
#define FORMULON_EVAL_EVAL_STATE_H_

#include <cstddef>
#include <cstdint>
#include <unordered_map>
#include <vector>

#include "value.h"

namespace formulon {

class Sheet;

namespace eval {

/// Per-`evaluate()` recursive-evaluation state.
///
/// Single-threaded: all mutations happen on the thread that owns the
/// top-level `evaluate()` call. Lifetime spans a single `evaluate()`
/// invocation; callers are expected to construct a fresh `EvalState` per
/// call rather than persist one across calls.
class EvalState {
 public:
  EvalState() = default;
  ~EvalState() = default;

  EvalState(const EvalState&) = delete;
  EvalState& operator=(const EvalState&) = delete;
  EvalState(EvalState&&) = delete;
  EvalState& operator=(EvalState&&) = delete;

  /// Tries to push `(sheet, row, col)` onto the in-progress stack. Returns
  /// `false` when the address is already present anywhere on the stack (a
  /// direct or indirect cycle) and does not modify the stack. Returns
  /// `true` and pushes the frame otherwise; the caller MUST match each
  /// successful push with exactly one `pop_cell(sheet, row, col)` to keep
  /// the stack balanced. `sheet` is compared by pointer identity, so
  /// callers must thread the same `Sheet*` through the push / pop pair.
  bool push_cell(const Sheet* sheet, std::uint32_t row, std::uint32_t col);

  /// Pops the top frame. In debug builds, asserts that the top matches
  /// `(sheet, row, col)`; in release builds the mismatch is silent but
  /// still unbalanced — callers must pair pushes with pops along every
  /// path.
  void pop_cell(const Sheet* sheet, std::uint32_t row, std::uint32_t col);

  /// Returns a pointer to the memoised result for `(sheet, row, col)` or
  /// `nullptr` when no result has been recorded for that address yet. The
  /// pointer is valid until the next `memoize()` call (which may rehash
  /// the map).
  const Value* lookup_memo(const Sheet* sheet, std::uint32_t row,
                           std::uint32_t col) const noexcept;

  /// Records `value` as the memoised result for `(sheet, row, col)`.
  /// Overwrites any prior entry for the same key; formula cells are
  /// idempotent within a single `evaluate()` invocation so overwriting is
  /// harmless.
  void memoize(const Sheet* sheet, std::uint32_t row, std::uint32_t col,
               Value value);

  /// Current nesting depth of the in-progress stack. Useful for tests and
  /// as a coarse guard against runaway recursion.
  std::size_t depth() const noexcept { return stack_.size(); }

 private:
  // Composite key identifying a single cell across all sheets in the
  // current workbook. `sheet` is compared by pointer identity; the
  // evaluator only ever holds non-owning pointers to sheets owned by the
  // bound workbook, so pointer identity is stable for the lifetime of a
  // single `evaluate()` call.
  struct CellKey {
    const Sheet* sheet;
    std::uint32_t row;
    std::uint32_t col;
    friend bool operator==(const CellKey& a, const CellKey& b) noexcept {
      return a.sheet == b.sheet && a.row == b.row && a.col == b.col;
    }
  };

  // Hash combiner patterned after boost::hash_combine. Cells within a
  // single sheet frequently collide on one coordinate (e.g. a column of
  // formulas sharing `col`), so the combiner must redistribute bits.
  struct CellKeyHash {
    std::size_t operator()(const CellKey& k) const noexcept {
      std::size_t h = reinterpret_cast<std::uintptr_t>(k.sheet);
      h ^= (static_cast<std::uint64_t>(k.row) << 1) + 0x9e3779b9ULL + (h << 6) + (h >> 2);
      h ^= (static_cast<std::uint64_t>(k.col) << 1) + 0x9e3779b9ULL + (h << 6) + (h >> 2);
      return h;
    }
  };

  // Callstack of addresses currently being evaluated. Depth is expected
  // to stay shallow (single-digit in practice), so linear search on
  // push_cell is cheaper than a parallel hash set.
  std::vector<CellKey> stack_;
  // Memoised results, keyed by the composite address.
  std::unordered_map<CellKey, Value, CellKeyHash> memo_;
};

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_EVAL_STATE_H_
