// Double-byte code page tables behind CHAR, CODE and the byte-counting text functions.
//
// Each table is the 94x94 row/cell grid of one code page (JIS X 0208, GB 2312,
// KS X 1001). Only row/cell -> Unicode is stored; the reverse direction is a
// linear scan. The generated TUs keep a delta-encoded bit stream and decode
// it on first use into a function-local static dense array.

#ifndef FORMULON_EVAL_DBCS_TABLE_H_
#define FORMULON_EVAL_DBCS_TABLE_H_

#include <cstddef>
#include <cstdint>

#include "excel_locale.h"

namespace formulon {
namespace eval {

constexpr std::size_t kDbcsGridSize = 94;

/// Returns `(row << 8) | cell` (both 1-based, in [1, 94]) for `codepoint`, or
/// 0 when the code page cannot encode it.
std::uint16_t lookup_unicode_to_dbcs(DbcsCodepage codepage, std::uint32_t codepoint) noexcept;

/// Returns the BMP code point at 1-based `row` / `cell`, or 0 when the slot is
/// unassigned or either index lies outside [1, 94].
std::uint16_t lookup_dbcs_to_unicode(DbcsCodepage codepage, std::uint8_t row, std::uint8_t cell) noexcept;

/// Added to a packed row/cell to form the value CODE returns and CHAR accepts:
/// 0x2020 for the JIS form (ja), 0xA0A0 for the EUC form (zh, ko; the JIS form
/// plus 0x8080). 0 for `kNone`.
std::uint16_t dbcs_code_bias(DbcsCodepage codepage) noexcept;

namespace dbcs_detail {

/// Dense row-major grid; slot `(row - 1) * 94 + (cell - 1)` holds the code
/// point, 0 when unassigned. The constructor decodes a generated bit stream.
struct DbcsCells {
  DbcsCells(const std::uint8_t* encoded, std::size_t size) noexcept;
  std::uint16_t unicode[kDbcsGridSize * kDbcsGridSize];
};

// Defined in the generated dbcs_<codepage>_table.cpp TUs.
const std::uint16_t* jis0208_cells() noexcept;
const std::uint16_t* gb2312_cells() noexcept;
const std::uint16_t* ksx1001_cells() noexcept;

}  // namespace dbcs_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DBCS_TABLE_H_
