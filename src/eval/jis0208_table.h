// Copyright 2026 libraz. Licensed under the MIT License.
//
// JIS X 0208 reverse lookup table and helper accessors for Mac Excel ja-JP
// CODE / CHAR parity. The dense reverse table itself lives in the generated
// translation unit `jis0208_table.cpp`; the lookup helpers declared here are
// implemented in `jis0208_table_lookup.cpp`.
//
// Index convention: row and cell are 1-based JIS X 0208 indices in [1, 94].
// `kJis0208Reverse[(row - 1) * kJis0208CellCount + (cell - 1)]` is the BMP
// Unicode codepoint mapped to that slot, or 0 (sentinel) when the slot is
// unassigned. U+0000 itself is never assigned in JIS X 0208 so the sentinel
// is unambiguous.

#ifndef FORMULON_EVAL_JIS0208_TABLE_H_
#define FORMULON_EVAL_JIS0208_TABLE_H_

#include <cstddef>
#include <cstdint>

namespace formulon {
namespace eval {

constexpr std::size_t kJis0208RowCount = 94;
constexpr std::size_t kJis0208CellCount = 94;

/// Dense JIS X 0208 -> Unicode reverse table. See file header for the index
/// convention. Defined in `jis0208_table.cpp`.
extern const std::uint16_t kJis0208Reverse[kJis0208RowCount * kJis0208CellCount];

/// Returns the JIS X 0208 row-cell encoding for `unicode_codepoint` packed as
/// `((row + 0x20) << 8) | (cell + 0x20)` — i.e. the value Mac Excel ja-JP's
/// CODE returns for a DBCS character. Returns 0 when the codepoint has no
/// mapping in JIS X 0208 (e.g. emoji, NEC extensions, characters outside the
/// BMP). Implementation is a linear scan over the 17 KB dense table; the
/// hot path is one call per CODE() invocation and the table stays L1-resident.
std::uint16_t lookup_unicode_to_jis0208(std::uint32_t unicode_codepoint);

/// Returns the BMP Unicode codepoint mapped to the JIS X 0208 slot identified
/// by the encoded byte pair `(jis_hi_byte, jis_lo_byte)` — i.e. the value
/// Mac Excel ja-JP's CHAR returns for a DBCS-region argument. Both bytes
/// must lie in [0x21, 0x7E] (corresponding to row, cell in [1, 94]); any
/// out-of-range byte or unassigned slot returns 0.
std::uint16_t lookup_jis0208_to_unicode(std::uint8_t jis_hi_byte, std::uint8_t jis_lo_byte);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_JIS0208_TABLE_H_
