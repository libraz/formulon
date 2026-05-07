// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Pivot layout projection. This layer converts an evaluated `PivotResult`
// into absolute grid cells that a UI can draw with the ordinary cell
// renderer. It does not write into the owning Sheet's cell store.

#ifndef FORMULON_PIVOT_PIVOT_LAYOUT_H_
#define FORMULON_PIVOT_PIVOT_LAYOUT_H_

#include <cstdint>
#include <deque>
#include <string>
#include <vector>

#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::pivot {

enum class PivotCellKind : std::uint8_t {
  Header = 0,
  RowLabel = 1,
  ColLabel = 2,
  Data = 3,
  RowSubtotal = 4,
  ColSubtotal = 5,
  GrandTotal = 6,
  Blank = 7,
};

struct PivotLayoutOptions {
  std::string grand_total_label = "Grand Total";
  std::string values_label = "Values";
};

struct PivotCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  Value value = Value::blank();
  PivotCellKind kind = PivotCellKind::Blank;
  std::uint32_t depth = 0;
  std::string field_name;
  std::string number_format;
};

struct PivotCells {
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  std::vector<PivotCell> cells;

  // Backing storage for Text values emitted by the layout layer. This
  // includes synthesized labels and Text values copied from `PivotResult`,
  // so `PivotCells` can outlive the evaluated result it was projected from.
  std::deque<std::string> text_storage;
};

/// Projects an evaluated pivot result into absolute sheet coordinates.
///
/// The returned cells use `table.anchor_row()` / `anchor_col()` as their
/// top-left origin. Missing row or column axes are represented as one
/// implicit leaf so the data area remains rectangular.
Expected<PivotCells, Error> layout(const PivotTable& table, const PivotResult& result,
                                   const PivotLayoutOptions& options = PivotLayoutOptions{});

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_LAYOUT_H_
