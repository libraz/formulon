// Double-byte code page tables behind CHAR, CODE and the byte-counting text functions.
//
// Each table covers the code space of one code page (JIS X 0208, GB 2312,
// KS X 1001 as 94x94 row/cell grids; Big5 as lead 0xA1..0xF9 over trails
// 0x40..0x7E and 0xA1..0xFE). Only lead/trail -> Unicode is stored; the
// reverse direction is a linear scan. The generated TUs keep a delta-encoded bit stream and decode
// it on first use into a function-local static dense array.

#ifndef FORMULON_EVAL_DBCS_TABLE_H_
#define FORMULON_EVAL_DBCS_TABLE_H_

#include <cstddef>
#include <cstdint>

#include "excel_locale.h"

namespace formulon {
namespace eval {

constexpr std::size_t kDbcsGridSize = 94;

/// Returns `(lead << 8) | trail` for `codepoint` in the table's byte space (the
/// code minus `dbcs_code_bias`; row / cell in [1, 94] for the grid tables), or
/// 0 when the code page cannot encode it.
std::uint16_t lookup_unicode_to_dbcs(DbcsCodepage codepage, std::uint32_t codepoint) noexcept;

/// Returns the BMP code point at `lead` / `trail` in the table's byte space, or
/// 0 when the slot is unassigned or outside the code space.
std::uint16_t lookup_dbcs_to_unicode(DbcsCodepage codepage, std::uint8_t row, std::uint8_t cell) noexcept;

/// Added to a packed lead/trail to form the value CODE returns and CHAR accepts:
/// 0x2020 for the JIS form (ja), 0xA0A0 for the EUC form (zh-CN, ko; the JIS
/// form plus 0x8080), 0 for Big5 (its packed value is the raw code) and `kNone`.
std::uint16_t dbcs_code_bias(DbcsCodepage codepage) noexcept;

namespace dbcs_detail {

/// Code space of one table: `lead_count` leads from `lead_first`, each with
/// the trail bytes of up to two ranges (`count == 0` marks an unused range).
struct DbcsShape {
  std::uint8_t lead_first;
  std::uint8_t lead_count;
  struct TrailRange {
    std::uint8_t first;
    std::uint8_t count;
  } trails[2];
};

/// Decodes a generated bit stream into `slots` dense row-major slots (lead,
/// then trail in range order), 0 where unassigned.
void decode_dbcs_cells(const std::uint8_t* encoded, std::size_t size, std::uint16_t* out, std::size_t slots) noexcept;

template <std::size_t Slots>
struct DbcsCells {
  DbcsCells(const std::uint8_t* encoded, std::size_t size) noexcept : unicode{} {
    decode_dbcs_cells(encoded, size, unicode, Slots);
  }
  std::uint16_t unicode[Slots];
};

struct DbcsGrid {
  DbcsShape shape;
  const std::uint16_t* unicode;
};

// Defined in the generated dbcs_<codepage>_table.cpp TUs.
const DbcsGrid& jis0208_grid() noexcept;
const DbcsGrid& gb2312_grid() noexcept;
const DbcsGrid& ksx1001_grid() noexcept;
const DbcsGrid& big5_grid() noexcept;

}  // namespace dbcs_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DBCS_TABLE_H_
