//
// Writer of the `xl/workbook.bin` stream: workbook globals, the sheet
// bundle, the supporting-book and `BrtExternSheet` tables and the `BrtName`
// table. The name and sheet-range tables are built once per workbook and
// shared with every sheet's cell encoder so `ilbl` / `ixti` assignments stay
// consistent.

#ifndef FORMULON_IO_XLSB_WORKBOOK_BIN_WRITER_H_
#define FORMULON_IO_XLSB_WORKBOOK_BIN_WRITER_H_

#include <cstdint>
#include <string>
#include <vector>

#include "io/xlsb/ptg_writer.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

/// One `BrtName` slot: the record's name and scope (`-1`: workbook, else
/// the 0-based sheet a stub is scoped to). Defined-name slots carry their
/// text only; their scope is read from `Workbook::defined_names()`.
struct OrderedName {
  std::string name;
  std::int32_t itab = -1;
};

/// Builds the workbook's `BrtName` record order (defined names first, then
/// placeholders and stubs).
void BuildOrderedNames(const Workbook& wb, std::vector<OrderedName>& ordered_names);

/// Builds the `name -> ilbl` map `encode_ptgs` consults for a formula that
/// lives in `scope_sheet_id` (`-1` for a workbook-scoped defined name).
NameTable BuildNameTableForScope(const Workbook& wb, const std::vector<OrderedName>& ordered_names,
                                 std::int32_t scope_sheet_id);

/// Builds the `BrtExternSheet` table for the whole workbook, with the
/// external links its entries can name.
Expected<SheetRangeTable, Error> BuildSheetRangeTable(const Workbook& wb, const std::vector<std::string>& sheet_names);

/// Emits the whole `xl/workbook.bin` record stream. `link_rel_ids` holds
/// the workbook relationship id of each external link part, in
/// `sheet_ranges.links` order.
Expected<std::vector<std::uint8_t>, Error> BuildWorkbookBin(const Workbook& wb,
                                                            const std::vector<OrderedName>& ordered_names,
                                                            const SheetRangeTable& sheet_ranges,
                                                            const std::vector<std::string>& sheet_names,
                                                            const std::vector<std::string>& link_rel_ids);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_WORKBOOK_BIN_WRITER_H_
