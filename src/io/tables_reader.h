// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// Per-column metadata for a table.
///
/// `id` is the 1-based column id Excel assigns; `name` is the displayed
/// column header. `totals_label` and `totals_function` are populated
/// when the table has a totals row that selects either a literal label
/// (`totalsRowLabel="..."`) or a built-in function
/// (`totalsRowFunction="sum"|...`); both default to empty.
/// `calculated_column_formula` carries the raw text of the column's
/// `<calculatedColumnFormula>` child, when present. An empty string
/// means the element was absent in the source XML; the writer therefore
/// omits the element entirely on round-trip.
struct TableColumn {
  std::uint32_t id = 0;
  std::string name;
  std::string totals_label;
  std::string totals_function;
  /// Raw text of `<calculatedColumnFormula>` (no leading `=`). Empty
  /// when the element was absent in the source XML; structured-
  /// references inside the text are NOT parsed or evaluated here.
  std::string calculated_column_formula;
};

/// In-memory representation of one `xl/tables/tableN.xml` part.
///
/// `id`, `name`, `display_name` and `ref` come straight from the root
/// `<table>` attributes. `sheet_index` is the workbook-relative index
/// of the sheet that owns the table (resolved by the caller via the
/// sheet rels file); the reader does not infer this from the part name.
/// `header_row` and `totals_row` reflect the `headerRowCount` /
/// `totalsRowCount` attributes (any value >= 1 enables the row;
/// `headerRowCount="0"` explicitly disables the header).
struct TableMetadata {
  std::uint32_t id = 0;
  std::string name;
  std::string display_name;
  std::string ref;
  std::size_t sheet_index = 0;
  bool header_row = true;
  bool totals_row = false;
  std::vector<TableColumn> columns;
};

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
