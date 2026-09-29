//
// Tables-part reader (`xl/tables/tableN.xml`). Extracts the metadata
// required to round-trip a table definition without invoking the full
// structured-reference machinery (which lands in Phase 4). Each table
// part is owned by exactly one sheet, located via the sheet's rels file
// (`xl/worksheets/_rels/sheetN.xml.rels`); the OOXML reader resolves
// the relationship and hands the bytes to this reader along with the
// owning sheet's workbook-relative index.
//
// Calculated-column formulas (`<calculatedColumnFormula>`) ARE preserved
// verbatim on each `TableColumn` for honest round-trip; the structured-
// reference parser does not run at this layer (the formula text is not
// rewritten or evaluated here).

#ifndef FORMULON_IO_TABLES_READER_H_
#define FORMULON_IO_TABLES_READER_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// Parses one `xl/tables/tableN.xml` part.
///
/// Behaviour:
///   * The root must be `<table>`. Anything else is treated as the
///     wrong content type — we surface `kIoContentTypeInvalid` rather
///     than `kIoXmlParse` so callers can distinguish "the bytes are
///     well-formed XML but not a table part" from "miniz handed us
///     garbage".
///   * `id` defaults to 0 if the attribute is missing or non-numeric;
///     `name` and `display_name` default to empty strings. The `ref`
///     attribute is required: missing/empty `ref` is `kIoSheetCorrupt`
///     (Excel rejects such tables and a writer emitting one would
///     produce an unopenable workbook).
///   * `headerRowCount="0"` sets `header_row=false`. Any other value
///     (including absence) leaves it `true`, matching the OOXML
///     default.
///   * `totalsRowCount=` defaults to 0; any value `>= 1` flips
///     `totals_row` to `true`.
///   * `<tableColumns>` is walked in document order. Each
///     `<tableColumn>` captures `id`, `name`, `totalsRowLabel`, and
///     `totalsRowFunction`. If both `totalsRowLabel` and
///     `totalsRowFunction` are present, both are preserved (Excel
///     emits at most one, but capturing both keeps the round-trip
///     contract honest). The `<calculatedColumnFormula>` child, when
///     present, is preserved verbatim on `calculated_column_formula`;
///     absence and an empty element collapse to the empty string,
///     and the writer omits the element when the field is empty (so
///     a workbook without calc-column formulas round-trips byte-
///     compatibly).
///
/// Errors:
///   * `kIoXmlParse` — pugixml could not parse `table_bytes`.
///   * `kIoContentTypeInvalid` — root element is not `<table>`.
///   * `kIoSheetCorrupt` — required `ref` attribute is missing or
///     empty.
Expected<TableMetadata, Error> read_table(const std::vector<std::uint8_t>& table_bytes, std::size_t sheet_index);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_TABLES_READER_H_
