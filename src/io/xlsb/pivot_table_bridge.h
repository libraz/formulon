//
// Bridge between an XLSB pivot table part (`xl/pivotTables/pivotTableN.bin`)
// and the in-memory `pivot::PivotTable`. Reading fills the members the xlsx
// pivot table reader fills for the same table, spelled the way that reader
// stores them (root and field attributes as text, `<rowItems>` / `<colItems>`,
// `<autoSortScope>`, each `<filter>` and the table extLst as raw XML), so a
// table read from an .xlsb and from its .xlsx twin compare equal. Writing
// encodes from the model alone.
//
// Header bits no Excel-saved part varies read as the model's defaults; a part
// departing from them is refused as unmeasured. Windows Excel also sets
// fFirstColumn and the ascending bit of manually ordered fields; those bits
// follow the saving host and stay out of the model; a Windows-saved descending
// field is unmeasured and refused. A state not encoded is not
// writable unless it is what Excel writes in its absence: `<formats>` and its
// siblings, a table extLst other than Excel's own, a field's x14 extension and
// a data field's extLst. So is an edited (`layout_dirty`), blocked or
// unauthored-span table, whose line items and extent need the report projection.

#ifndef FORMULON_IO_XLSB_PIVOT_TABLE_BRIDGE_H_
#define FORMULON_IO_XLSB_PIVOT_TABLE_BRIDGE_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "excel_profile.h"
#include "io/zip_reader.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Decodes a pivot table part. Errors are the codecs' (kIoXlsbRecordCorrupt,
/// kIoXlsbRecordTruncated, kIoXlsbUnknownRecord), and kIoXlsbUnknownRecord
/// with an "unmeasured" message for a record state with no Excel-saved sample.
Expected<pivot::PivotTable, Error> pivot_table_from_xlsb(ByteSpan part);

/// Encodes `table`, which sits on sheet `sheet_index`, as `host` saves it.
/// A model state the records cannot hold, or one with no Excel-saved sample,
/// is kIoFeatureNotWritable (`kind=pivot_table`).
Expected<std::vector<std::uint8_t>, Error> pivot_table_to_xlsb(const pivot::PivotTable& table, ExcelHost host,
                                                               std::size_t sheet_index);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_TABLE_BRIDGE_H_
