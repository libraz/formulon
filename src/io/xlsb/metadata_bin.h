//
// `xl/metadata.bin` — the cell-metadata part Excel pairs with the
// `BrtCellMeta` record it writes ahead of a spilled dynamic-array anchor.
//
// The part declares a table of metadata types (one of which is the
// dynamic-array type, named `XLDAPR`) followed by a table of cell-metadata
// entries, each naming one or more of those types. A worksheet's
// `BrtCellMeta` record carries the 1-based index of an entry in that second
// table, so the anchor and the part must agree on the numbering: an index
// that dangles past the table, or that resolves to a non-dynamic-array entry,
// makes Excel repair the file on open.
//
// Two sources of the part exist. `build_dynamic_array_metadata_bin` generates
// one containing a single XLDAPR entry, used when the model has spill anchors
// and the source package brought no metadata part of its own. Otherwise a
// retained passthrough `xl/metadata.bin` ships verbatim, and the index has to
// be recovered from its bytes with `find_dynamic_array_cell_meta_index`.
//
// Design references:
//   * [MS-XLSB] 2.1.7.32 (Metadata part), 2.4.x (BrtMdtinfo / BrtMdb)

#ifndef FORMULON_IO_XLSB_METADATA_BIN_H_
#define FORMULON_IO_XLSB_METADATA_BIN_H_

#include <cstdint>
#include <vector>

#include "io/zip_reader.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Builds the XLDAPR metadata stream Excel uses to mark spilled
/// dynamic-array anchors. The record sequence and constant GUID are the
/// canonical dynamic-array metadata shape Excel writes in XLSB workbooks.
/// The stream declares exactly one metadata type and one cell-metadata
/// entry, so a cell record naming this part always uses index 1.
std::vector<std::uint8_t> build_dynamic_array_metadata_bin();

/// Returns the 1-based index of the cell-metadata entry carrying the
/// dynamic-array (XLDAPR) type — the value a `BrtCellMeta` record's `ifmd`
/// field must hold to name it — or 0 when no such entry can be identified.
///
/// A passthrough metadata part is untrusted input, so truncation, malformed
/// framing and an absent XLDAPR type all report 0 rather than a guess. Every
/// read is bounds-checked. The caller's response to 0 is to emit no
/// `BrtCellMeta` record at all, which is safe for any part.
std::uint32_t find_dynamic_array_cell_meta_index(ByteSpan bytes);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_METADATA_BIN_H_
