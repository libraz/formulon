// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Output shape of a single pivot evaluation. Stored on the owning
// `PivotTable` so that `GETPIVOTDATA` can consult the most recent result
// without re-running the aggregation pipeline. See
// backup/plans/15-pivot-and-advanced.md §15.1.4.

#ifndef FORMULON_PIVOT_PIVOT_RESULT_H_
#define FORMULON_PIVOT_PIVOT_RESULT_H_

#include <deque>
#include <string>
#include <vector>

#include "value.h"

namespace formulon::pivot {

/// Tree node for the row-axis hierarchy. Children of a node represent the
/// next row field nested under this label (left-to-right in the rendered
/// pivot). Leaves correspond to rows in the `values` matrix.
struct RowHierarchyNode {
  std::string label;
  std::vector<RowHierarchyNode> children;
};

/// Tree node for the column-axis hierarchy. Children represent the next
/// column field nested under this label (top-to-bottom in the rendered
/// pivot). Leaves correspond to columns in the `values` matrix.
struct ColHierarchyNode {
  std::string label;
  std::vector<ColHierarchyNode> children;
};

/// Snapshot of a single pivot evaluation.
///
/// Storage is heap-owned and self-contained: each `Value` (including any
/// `Text` payload) must outlive the per-evaluation arena that produced it
/// because `GETPIVOTDATA` reads from `PivotTable::last_result_` outside of
/// any specific evaluation's arena lifetime. Do not back any `Value` here
/// with `string_view`s into arena memory.
struct PivotResult {
  /// Row hierarchy: tree of unique row-field values (left-to-right
  /// nesting). Leaves index into the rows of `values`.
  std::vector<RowHierarchyNode> rows;

  /// Column hierarchy: tree of unique column-field values (top-to-bottom
  /// nesting). Leaves index into the columns of `values`.
  std::vector<ColHierarchyNode> cols;

  /// Aggregated values. `values[r][c][a]` is the aggregation slot `a` at
  /// row leaf `r` and column leaf `c`.
  std::vector<std::vector<std::vector<Value>>> values;

  /// Per-row subtotals at each subtotal level (parallel to the row
  /// hierarchy).
  std::vector<std::vector<Value>> subtotals;

  /// Grand total across every aggregation. Defaults to `Blank`.
  Value grand_total = Value::blank();

  /// Lifetime-stable backing store for any `Value::text` payload appearing
  /// in `values`, `subtotals`, or `grand_total`. The evaluator may need to
  /// surface text values from the source cache (e.g. when MAX-ing over a
  /// pure-text column); copying the underlying bytes here decouples the
  /// result's lifetime from the cache it was computed against, satisfying
  /// the contract spelled out on this struct's docstring.
  ///
  /// `std::deque` gives pointer/iterator stability across appends so
  /// `string_view`s into earlier entries remain valid as later entries are
  /// pushed.
  std::deque<std::string> text_storage;
};

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_RESULT_H_
