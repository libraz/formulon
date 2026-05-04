// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `SpillCommitter::commit` implementation. The behavioural contract — in
// particular the four-case dispatch — is documented on the header.

#include "eval/spill_committer.h"

#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "sheet.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

Value SpillCommitter::commit(Value v) const {
  // Scalar passthrough — the common case. Cheap kind-check, no allocation.
  if (!v.is_array()) {
    return v;
  }
  // Inactive committer: caller did not opt in to spill, so we behave like
  // a read-only context and surface the array verbatim. The eventual
  // top-level consumer (writer / caller) decides what to do.
  if (sheet_ == nullptr) {
    return v;
  }

  const std::uint32_t rows = v.as_array_rows();
  const std::uint32_t cols = v.as_array_cols();
  if (rows == 0U || cols == 0U) {
    // Producers never emit a degenerate shape today; surface #VALUE! so
    // the caller does not silently discard the result.
    return Value::error(ErrorCode::Value);
  }

  // Deep-copy the row-major cells into an owned vector. Text payloads in
  // `cells` are still string_views into the source arena; that's fine —
  // `Sheet::commit_spill` re-interns Text bytes into its own owned_strings
  // so the spill region's lifetime is independent of the producing arena.
  const Value* src = v.as_array_cells();
  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  std::vector<Value> cells_vec;
  cells_vec.reserve(total);
  for (std::size_t i = 0; i < total; ++i) {
    cells_vec.push_back(src[i]);
  }

  // commit_spill sets the anchor's cached_value to either cells[0] (on
  // success) or #SPILL! (on collision); resolve_cell_value at the anchor
  // returns that stored value. The bool return is intentionally ignored:
  // either outcome is encoded in the resolved value the caller receives.
  (void)sheet_->commit_spill(row_, col_, rows, cols, std::move(cells_vec));
  return sheet_->resolve_cell_value(row_, col_);
}

}  // namespace eval
}  // namespace formulon
