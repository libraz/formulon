// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of `find_pivot_at_anchor`. See pivot_index.h for the
// public contract.

#include "pivot/pivot_index.h"

#include <cstdint>
#include <memory>
#include <string_view>

#include "pivot/pivot_table.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace pivot {

const PivotTable* find_pivot_at_anchor(const Workbook& wb, std::string_view sheet_name, std::uint32_t row,
                                       std::uint32_t col) noexcept {
  // Resolve the sheet via the workbook's case-insensitive lookup so this
  // helper agrees with `EvalContext::resolve_ref` on sheet identity.
  const Sheet* sheet = wb.sheet_by_name(sheet_name);
  if (sheet == nullptr) {
    return nullptr;
  }
  for (const std::unique_ptr<PivotTable>& pt : sheet->pivot_tables()) {
    if (pt == nullptr) {
      continue;
    }
    if (pt->contains(row, col)) {
      return pt.get();
    }
  }
  return nullptr;
}

}  // namespace pivot
}  // namespace formulon
