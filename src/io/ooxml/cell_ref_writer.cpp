// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the OOXML writer's cell-reference formatters.

#include "io/ooxml/cell_ref_writer.h"

#include <cstdint>
#include <string>

#include "sheet.h"  // for MergeRange

namespace formulon {
namespace io {

void AppendColumnLettersForRef(std::string& out, std::uint32_t col) {
  char buf[4];
  std::uint32_t i = 0;
  std::uint32_t v = col + 1;
  while (v > 0 && i < 4) {
    const std::uint32_t rem = (v - 1) % 26U;
    buf[i++] = static_cast<char>('A' + rem);
    v = (v - 1) / 26U;
  }
  while (i > 0) {
    out.push_back(buf[--i]);
  }
}

void AppendCellRefForRef(std::string& out, std::uint32_t row, std::uint32_t col) {
  AppendColumnLettersForRef(out, col);
  out.append(std::to_string(row + 1));
}

void AppendRangeRef(std::string& out, const MergeRange& r) {
  AppendCellRefForRef(out, r.first_row, r.first_col);
  if (r.first_row != r.last_row || r.first_col != r.last_col) {
    out.push_back(':');
    AppendCellRefForRef(out, r.last_row, r.last_col);
  }
}

}  // namespace io
}  // namespace formulon
