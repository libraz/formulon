//
// PtgList ([MS-XLSB] 2.5.98.52): a structured reference into a table, the
// binary form of `Table1[[#This Row],[Price]]`. 14 bytes: ptg 0x18, eptg
// 0x19, ixti (u16, the table's sheet), a u16 of columns (2 bits), rowType
// (5 bits, PtgRowType 2.5.98.73), two spacing bits, an unused bit, the data
// type (2 bits), invalid and nonresident, then listIndex (u32, the table's
// BrtBeginList idList), colFirst and colLast (u16, relative to the table).
//
// A reference is resolved against the workbook's tables by id on read and by
// name on write; one into another workbook (nonresident) or an invalid one is
// not decoded.

#ifndef FORMULON_IO_XLSB_PTG_LIST_H_
#define FORMULON_IO_XLSB_PTG_LIST_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <unordered_set>
#include <vector>

#include "io/xlsb/ptg_writer.h"
#include "io/zip_reader.h"
#include "parser/ast.h"
#include "table.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// First byte of a PtgList; the token's second byte is `kEptgList`.
inline constexpr std::uint8_t kPtgList = 0x18;
inline constexpr std::uint8_t kEptgList = 0x19;

/// Decodes the PtgList at `cursor` (its 0x18 byte included) into a
/// StructuredRef node and advances past it. `kIoXlsbUnsupportedPtg` when no
/// table of `tables` carries its listIndex or the reference is invalid or
/// nonresident.
Expected<parser::AstNode*, Error> decode_ptg_list(ByteSpan& cursor, Arena& arena,
                                                  const std::vector<TableMetadata>* tables);

/// The `idList` a table's part carries: its own id when non-zero and not
/// taken by an earlier table or in `reserved`, otherwise the smallest id no
/// table holds.
std::uint32_t xlsb_table_id(const std::vector<TableMetadata>& tables, std::size_t index,
                            const std::unordered_set<std::uint32_t>& reserved);

/// The 0-based sheet of the model or verbatim table a StructuredRef names, or nullopt.
std::optional<std::int32_t> structured_ref_sheet(const parser::AstNode& node, const SheetRangeTable& ranges);

/// Appends `node` (a StructuredRef) as a PtgList of Ptg class `cls`
/// (0x20 / 0x40 / 0x60). `kIoXlsbUnsupportedPtg` for an unknown table or
/// column, a selector combination PtgRowType has no value for, or a table
/// sheet `ranges` has no XTI for.
Expected<void, Error> encode_ptg_list(std::vector<std::uint8_t>& out, const parser::AstNode& node, std::uint8_t cls,
                                      const SheetRangeTable& ranges);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_LIST_H_
