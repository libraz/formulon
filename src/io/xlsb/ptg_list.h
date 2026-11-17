//
// PtgList ([MS-XLSB] 2.5.98.52): a structured reference into a table, the
// binary form of `Table1[[#This Row],[Price]]`. 14 bytes: ptg 0x18, eptg
// 0x19, ixti (u16, the table's sheet), a u16 of columns (2 bits), rowType
// (5 bits, PtgRowType 2.5.98.73), two spacing bits, an unused bit, the data
// type (2 bits), invalid and nonresident, then listIndex (u32, the table's
// BrtBeginList idList), colFirst and colLast (u16, relative to the table).
//
// A reference is resolved against the workbook's tables by id on read and by
// name on write; an invalid one is not decoded. One into another workbook
// (nonresident) carries listIndex 0xFFFFFFFF, column bits 2 for a column
// range and 0 otherwise, and zero column indexes: the table
// and column names follow in the formula's rgcb, and its ixti names the
// link's file-stem sheet entry. The rgcb holds a has-columns byte, the row
// type (u16) and the table name, then for columns a u16 0, a count byte and
// per column a more-follows byte, a u16 2 and the unescaped name; names are
// a u16 count of UTF-16 units. `#This Row` also sets flag bit 9.

#ifndef FORMULON_IO_XLSB_PTG_LIST_H_
#define FORMULON_IO_XLSB_PTG_LIST_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
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

/// True when the PtgList at `cursor` (its 0x18 byte included) names a table
/// of another workbook.
bool ptg_list_is_nonresident(const ByteSpan& cursor);

/// Decodes the nonresident PtgList at `cursor` into an ExternalRef naming a
/// table of link `book` (1-based), reading its names from `extra` (the
/// formula's rgcb), and advances both. The payload is the stored long form.
/// A totals-row form is `kIoXlsbUnsupportedPtg`.
Expected<parser::AstNode*, Error> decode_external_ptg_list(ByteSpan& cursor, ByteSpan& extra, std::uint32_t book,
                                                           Arena& arena);

/// True when the external table reference `node` (an ExternalRef with a
/// structured payload) has a PtgList form: any selector but `#Totals`, with
/// no column, one column or a column range.
bool external_ptg_list_encodable(const parser::AstNode& node);

/// Appends the external table reference `node` as a nonresident PtgList of
/// Ptg class `cls` through XTI `ixti`, and its names to `extra`.
/// `kIoXlsbUnsupportedPtg` when `external_ptg_list_encodable` is false.
Expected<void, Error> encode_external_ptg_list(std::vector<std::uint8_t>& out, std::vector<std::uint8_t>& extra,
                                               const parser::AstNode& node, std::uint8_t cls, std::uint16_t ixti);

/// The `idList` a table's part carries: its own id when non-zero and not
/// taken by an earlier table or in `reserved`, otherwise the smallest id no
/// table holds.
std::uint32_t xlsb_table_id(const std::vector<TableMetadata>& tables, std::size_t index,
                            const std::unordered_set<std::uint32_t>& reserved);

/// True when a StructuredRef names one column in one row band: `[@Col]`,
/// `[[#This Row],[Col]]`, `[[#Headers],[Col]]` or `[[#Totals],[Col]]`.
/// Excel gives such a reference no implicit intersection.
bool structured_ref_is_single_cell(const parser::AstNode& node);

/// The same test on a bracket payload's text, for a reference into another workbook's table.
bool structured_payload_is_single_cell(std::string_view structured);

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
