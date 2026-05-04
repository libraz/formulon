// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `SpillCommitter` is the side-effect surface that turns an evaluator-
// produced `Value::Array` into a dynamic-array spill committed at the
// owning formula cell. It exists as a thin object so the Excel
// "value-producing" side of the evaluator (`EvalContext`, the recursive
// `resolve_ref`, `cell_evaluator`, the lazy impls) is freed from owning
// the actual write authority — `EvalContext` becomes "DI / reference
// resolution context only", while `SpillCommitter` carries the
// (mutable_sheet, anchor) pair that authorises the write.
//
// Today this co-exists with `EvalContext::dispatch_array_result`: the
// Context method is preserved as a thin compatibility wrapper that
// delegates here, so existing call sites — including unit tests — keep
// compiling unchanged. New call sites should construct a
// `SpillCommitter` directly when they only need spill-write authority,
// without inflating an `EvalContext` they would otherwise not need.
//
// Lifetime: a `SpillCommitter` is non-owning and lightweight (two
// pointers + two row/col fields). The bound `Sheet` must outlive the
// committer.

#ifndef FORMULON_EVAL_SPILL_COMMITTER_H_
#define FORMULON_EVAL_SPILL_COMMITTER_H_

#include <cstdint>

#include "value.h"

namespace formulon {

class Sheet;

namespace eval {

/// Anchored writer that commits dynamic-array spill values for the
/// formula cell at `(row, col)` on `sheet`. See file header for the
/// rationale and lifetime requirements.
class SpillCommitter {
 public:
  /// Builds a committer authorised to spill onto `*sheet` at
  /// `(row, col)`. Both `sheet` and the bound row/col must remain
  /// addressable while the committer is alive; the committer does not
  /// take ownership.
  ///
  /// `sheet` may be `nullptr` to build a no-op committer that always
  /// returns the supplied `Value` unchanged. This is convenient for
  /// read-only evaluation contexts that still want a uniform commit-call
  /// surface.
  SpillCommitter(Sheet* sheet, std::uint32_t row, std::uint32_t col) noexcept : sheet_(sheet), row_(row), col_(col) {}

  /// Builds a no-op committer. Equivalent to `SpillCommitter(nullptr, 0, 0)`
  /// — every `commit()` call returns its argument unchanged.
  SpillCommitter() noexcept = default;

  /// Returns true when the committer is bound to a sheet AND a valid
  /// anchor. False committers (default-constructed or built with a null
  /// sheet) pass scalars through and surface arrays verbatim.
  bool active() const noexcept { return sheet_ != nullptr; }

  /// Returns the post-dispatch scalar value the caller should propagate.
  ///
  /// Behaviour mirrors `EvalContext::dispatch_array_result` exactly so
  /// the existing test suite continues to pass when call sites migrate:
  ///   1. Non-Array `v` is returned unchanged (the common scalar path).
  ///   2. Inactive committer (`!active()`) returns `v` unchanged.
  ///   3. A degenerate `0 x N` / `N x 0` array yields `#VALUE!`
  ///      defensively (producers should never emit that shape).
  ///   4. Otherwise the array's row-major cells are deep-copied into a
  ///      `Sheet::commit_spill` call at `(row, col)`. The return value
  ///      is `Sheet::resolve_cell_value(row, col)` — either the array's
  ///      top-left scalar on success or `#SPILL!` on collision.
  Value commit(Value v) const;

 private:
  Sheet* sheet_ = nullptr;
  std::uint32_t row_ = 0;
  std::uint32_t col_ = 0;
};

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SPILL_COMMITTER_H_
