//
// MS-XLSB external-link decoding: `xl/externalLinks/externalLink<N>.bin`
// into the same `io::ExternalBook` model the OOXML reader builds, so a
// cross-workbook reference resolves against an XLSB-sourced cache the
// same way it does against an xlsx one.
//
// ## Where the record layouts come from
//
// Established by differential decode rather than from a specification:
// the same workbook was re-saved by Excel as `.xlsx`, decoded with the
// OOXML external-link reader, and each binary field checked against the
// value that reader produced. The field widths inside a defined name's
// stored formula were isolated by building sources whose name lands on a
// different sheet, row and column and diffing the record byte that
// moved.
//
// One measured surprise is worth stating: inside an external link part a
// reference Ptg carries 16-bit rows and an inline `(itabFirst, itabLast)`
// sheet-index pair, not the 32-bit rows and `ixti` table index the same
// opcode uses inside a worksheet. The decoder therefore has its own
// reader for these rather than sharing `ptg_reader`'s.
//
// Anything outside the measured set is refused. The caller treats a
// refusal as "this link has no cache", which leaves a cross-workbook
// reference reading `#REF!` -- the behaviour before this module existed
// -- rather than resolving to a plausible wrong number. The part stays in
// the passthrough set either way, so a refusal costs no round-trip
// fidelity.

#ifndef FORMULON_IO_XLSB_EXTERNAL_LINK_READER_H_
#define FORMULON_IO_XLSB_EXTERNAL_LINK_READER_H_

#include "io/external_book.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Decodes one external link part into its cached sheet names, defined
/// names and cell values.
///
/// Errors:
///   * `kIoXlsbRecordTruncated` — a record overruns the part.
///   * `kIoXlsbCorrupt`         — a cached address or a name's stored
///                                formula uses an encoding this module
///                                has not measured.
Expected<ExternalBook, Error> read_external_link_bin(ByteSpan cursor);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_EXTERNAL_LINK_READER_H_
