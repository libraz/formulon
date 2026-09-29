//
// Shared A1 reference decoder: column letters, decimal row indices, and
// composite "AB12"-shaped references. Consolidates the equivalent helpers
// previously duplicated in `cell_parser.cpp` and `sax_xml_reader.cpp`.
//
// Both readers parse OOXML `<c r="...">` attributes and reference strings
// with identical semantics (column letters A..XFD, decimal row indices,
// 32-bit overflow rejection); this header is the single source of truth.

#ifndef FORMULON_UTILS_A1_REF_H_
#define FORMULON_UTILS_A1_REF_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "utils/a1_column.h"

namespace formulon::a1 {

/// Excel 365 grid bounds: rows are addressable as 0..kMaxRows-1 and columns
/// as 0..kMaxCols-1 (A..XFD). `Sheet::kMaxRows` / `Sheet::kMaxCols` alias these.
inline constexpr std::uint32_t kMaxRows = 1048576U;
inline constexpr std::uint32_t kMaxCols = kMaxColumns;

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
/// references beyond Excel's `kMaxCols` / `kMaxRows` ceilings
/// (e.g. a 3-letter column past XFD, or a row past 1,048,576), matching the
/// DOM cell-ref path in `cell_parser.cpp` so both OOXML readers converge.
bool parse_a1_ref(std::string_view text, std::uint32_t* out_row, std::uint32_t* out_col) noexcept;

/// Encodes a 0-based (row, col) into the Excel A1 address (1-based, e.g.
/// "A1", "AA1", "XFD1048576"). Returns an empty string when the column is
/// outside Excel's grid.
std::string encode_a1(std::uint32_t row, std::uint32_t col);

}  // namespace formulon::a1

#endif  // FORMULON_UTILS_A1_REF_H_
