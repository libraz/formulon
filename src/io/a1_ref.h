// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared A1 reference decoder: column letters, decimal row indices, and
// composite "AB12"-shaped references. Consolidates the equivalent helpers
// previously duplicated in `cell_parser.cpp` and `sax_xml_reader.cpp`.
//
// Both readers parse OOXML `<c r="...">` attributes and reference strings
// with identical semantics (column letters A..XFD, decimal row indices,
// 32-bit overflow rejection); this header is the single source of truth.

#ifndef FORMULON_IO_A1_REF_H_
#define FORMULON_IO_A1_REF_H_

#include <cstddef>
#include <cstdint>
#include <string_view>

namespace formulon {
namespace io {

/// Excel's column-letter ceiling. The largest valid Excel column is XFD
/// (3 letters); references with 4 or more leading letters are rejected.
constexpr std::size_t kMaxColumnLetters = 3U;

/// Parses Excel column letters (A-Z, AA-XFD) starting at `text[*pos]` into
/// a 1-based column index. Advances `*pos` past the consumed letters. On
/// failure (no letter consumed, or 4+ letters i.e. past XFD) returns false
/// and leaves `*out_col` and `*pos` unchanged, except that `*pos` may have
/// advanced over the partially consumed letters when the overflow guard
/// triggered.
bool parse_column_letters(std::string_view text, std::size_t* pos, std::uint32_t* out_col) noexcept;

/// Parses an unsigned decimal integer starting at `text[*pos]` into
/// `*out_val`. Advances `*pos` past the digits. Returns false when no digit
/// was consumed or the value would overflow `std::uint32_t`. On overflow
/// `*pos` is positioned just past the digits scanned so far.
bool parse_uint(std::string_view text, std::size_t* pos, std::uint32_t* out_val) noexcept;

/// Parses an A1-shaped cell reference (e.g. "AB12") into 0-based row and
/// 0-based column. Returns false on empty input, malformed letters, missing
/// row digits, trailing characters, or any overflow. Also rejects
/// references beyond Excel's `Sheet::kMaxCols` / `Sheet::kMaxRows` ceilings
/// (e.g. a 3-letter column past XFD, or a row past 1,048,576), matching the
/// DOM cell-ref path in `cell_parser.cpp` so both OOXML readers converge.
bool parse_a1_ref(std::string_view text, std::uint32_t* out_row, std::uint32_t* out_col) noexcept;

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_A1_REF_H_
