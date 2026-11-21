//
// Ignored cell errors as xlsb records: BrtBeginCellIgnoreECs, one BrtCellIgnoreEC per `<ignoredError>` (u32 flags,
// then an Sqrfx) and BrtEndCellIgnoreECs. The model keeps them as the retained `<ignoredErrors>` worksheet child, so
// both formats read and write the one element. Only the nine flag bits Excel writes are mapped; a record or an element
// carrying anything else is not converted.

#ifndef FORMULON_IO_XLSB_IGNORED_ERROR_RECORDS_H_
#define FORMULON_IO_XLSB_IGNORED_ERROR_RECORDS_H_

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "io/zip_reader.h"
#include "sheet_passthrough.h"

namespace formulon::io::xlsb {

/// The `<ignoredError>` element a BrtCellIgnoreEC payload describes; nullopt when the flags are zero or carry a bit
/// outside the nine mapped ones, or the Sqrfx does not fill the payload exactly.
std::optional<std::string> decode_ignored_error(ByteSpan payload);

/// Replaces the retained `<ignoredErrors>` child with `elements`, in order; an empty list removes it.
void write_ignored_errors(WorksheetRawExtensions& children, const std::vector<std::string>& elements);

/// Inserts the ignored-error block for the retained `<ignoredErrors>` child into the post-hyperlink tail, after the
/// print records, BrtBigName and the cell-watch block; nothing when the sheet has none. An element whose sqref does
/// not parse, or that sets no mapped flag or an attribute outside them, is left out.
void splice_ignored_errors(std::vector<std::uint8_t>& after_hyperlinks, const WorksheetRawExtensions& children);

}  // namespace formulon::io::xlsb

#endif  // FORMULON_IO_XLSB_IGNORED_ERROR_RECORDS_H_
