// Copyright 2026 libraz. Licensed under the MIT License.
//
// Cell-level value types used by the worksheet storage layer. A worksheet
// owns a row-sparse, column-dense cell store keyed by 0-based row index
// (see `sheet.h`); each cell carries either a literal value or a formula
// string with an associated cached result.
//
// `Cell` is intentionally a plain aggregate: the storage layer in `Sheet`
// owns lifetime, and downstream layers (parser, evaluator, OOXML writer)
// observe cells through pointers handed out by `Sheet::cell_at`.

#ifndef FORMULON_CELL_H_
#define FORMULON_CELL_H_

#include <cstdint>
#include <string>

#include "value.h"

namespace formulon {

/// 0-based cell coordinate within a single sheet.
///
/// `row` is in `[0, 1048576)` and `col` is in `[0, 16384)`, matching the
/// Excel 365 sheet dimensions (1,048,576 rows by 16,384 columns).
struct CellAddress {
  std::uint32_t row;
  std::uint32_t col;

  friend bool operator==(CellAddress a, CellAddress b) noexcept { return a.row == b.row && a.col == b.col; }
  friend bool operator!=(CellAddress a, CellAddress b) noexcept { return !(a == b); }
};

/// A single cell's persisted state.
///
/// `formula_text` is the raw formula string starting with `=` (empty when
/// the cell holds a literal). `cached_value` is the cell's effective Value:
/// for a literal cell, this *is* the value; for a formula cell, it is the
/// most recently computed result (populated lazily by the evaluator in a
/// follow-up increment — for now formula cells are stored with
/// `cached_value = Value::blank()` until evaluation is wired up).
struct Cell {
  std::string formula_text;
  Value cached_value = Value::blank();
};

}  // namespace formulon

#endif  // FORMULON_CELL_H_
