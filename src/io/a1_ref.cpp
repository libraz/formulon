// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the shared A1 reference decoder. See `a1_ref.h` for
// the contract.

#include "io/a1_ref.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

namespace formulon {
namespace io {

bool parse_column_letters(std::string_view text, std::size_t* pos, std::uint32_t* out_col) noexcept {
  std::uint32_t col = 0;
  std::size_t consumed = 0;
  while (*pos < text.size()) {
    const char c = text[*pos];
    if (c < 'A' || c > 'Z') {
      break;
    }
    if (consumed >= kMaxColumnLetters) {
      // Past XFD (4+ letters): reject. The caller only inspects the
      // boolean so leaving `*pos` advanced is fine; `*out_col` is
      // intentionally not written.
      return false;
    }
    col = col * 26U + static_cast<std::uint32_t>(c - 'A' + 1);
    ++(*pos);
    ++consumed;
  }
  if (consumed == 0U) {
    return false;
  }
  *out_col = col;
  return true;
}

bool parse_uint(std::string_view text, std::size_t* pos, std::uint32_t* out_val) noexcept {
  std::uint64_t v = 0;
  bool any = false;
  while (*pos < text.size()) {
    const char c = text[*pos];
    if (c < '0' || c > '9') {
      break;
    }
    v = v * 10U + static_cast<std::uint64_t>(c - '0');
    if (v > 0xFFFFFFFFULL) {
      return false;
    }
    ++(*pos);
    any = true;
  }
  if (!any) {
    return false;
  }
  *out_val = static_cast<std::uint32_t>(v);
  return true;
}

bool parse_a1_ref(std::string_view text, std::uint32_t* out_row, std::uint32_t* out_col) noexcept {
  if (text.empty()) {
    return false;
  }
  std::size_t i = 0;
  std::uint32_t col_1based = 0;
  if (!parse_column_letters(text, &i, &col_1based)) {
    return false;
  }
  std::uint32_t row_1based = 0;
  if (!parse_uint(text, &i, &row_1based)) {
    return false;
  }
  if (i != text.size()) {
    return false;
  }
  if (row_1based == 0U) {
    return false;
  }
  *out_row = row_1based - 1U;
  *out_col = col_1based - 1U;
  return true;
}

}  // namespace io
}  // namespace formulon
