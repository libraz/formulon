//
// MS-XLSB pivot-part decoding: `xl/pivotCache/pivotCacheDefinition*.bin`,
// `xl/pivotCache/pivotCacheRecords*.bin` and `xl/pivotTables/pivotTable*.bin`
// into the same `pivot::PivotCache` / `pivot::PivotTable` model the OOXML
// reader builds, so GETPIVOTDATA resolves against an XLSB-sourced pivot.
//
// ## Where the record layouts come from
//
// The pivot record ids and payload layouts here were established by
// differential decode rather than from a specification: the same workbook
// was re-saved by Excel as `.xlsx`, decoded with the OOXML pivot reader,
// and each binary field checked against the value that reader produced.
// The aggregation selector was isolated the same way, by re-saving one
// pivot per `<dataField subtotal>` value and diffing the record.
//
// That provenance is why these decoders refuse rather than guess. A pivot
// whose records use an encoding this module has not seen would otherwise
// decode to a plausible wrong number, which is worse than not decoding it
// at all: the caller treats any error as "leave this pivot alone", which
// is what every XLSB pivot did before this module existed. The parts stay
// in the passthrough set either way, so a refusal costs no round-trip
// fidelity.

#ifndef FORMULON_IO_XLSB_PIVOT_READER_H_
#define FORMULON_IO_XLSB_PIVOT_READER_H_

#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Decodes a pivot cache from its definition and records parts.
///
/// The two parts are decoded together because the record layout is not
/// self-describing: a record cell is a 4-byte shared-item index or an
/// 8-byte double depending on what the *definition* said about that
/// field, so the records cannot be parsed without it.
///
/// Errors:
///   * `kIoXlsbRecordTruncated` — a record overruns its part.
///   * `kIoXlsbRecordCorrupt`   — a record id, field count or record
///                                width this module cannot account for.
///                                Callers treat this as "skip the pivot".
Expected<pivot::PivotCache, Error> read_pivot_cache_bin(ByteSpan definition, ByteSpan records);

/// Decodes a pivot table definition part.
///
/// `PivotTable::pivot_cache_id` is left at its default: the binary
/// carries the cache binding in a header field this module does not
/// decode, and the relationship from the pivot-table part to its cache
/// definition part states the same thing unambiguously. The caller
/// resolves it from the rels and assigns the id.
Expected<pivot::PivotTable, Error> read_pivot_table_bin(ByteSpan cursor);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_READER_H_
