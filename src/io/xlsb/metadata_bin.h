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
// Two sources of the part exist. `build_metadata_bin` generates one holding
// the XLDAPR entry dynamic-array anchors name and the XLRICHVALUE blocks rich
// errors name, used when the source package brought no metadata part of its
// own. Otherwise a retained passthrough `xl/metadata.bin` ships verbatim, and
// the index has to be recovered from its bytes with
// `find_dynamic_array_cell_meta_index`.
//
// Design references:
//   * [MS-XLSB] 2.1.7.32 (Metadata part), 2.4.x (BrtMdtinfo / BrtMdb)

#ifndef FORMULON_IO_XLSB_METADATA_BIN_H_
#define FORMULON_IO_XLSB_METADATA_BIN_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "io/rich_value_errors.h"
#include "io/zip_reader.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Builds the metadata stream in the record sequence Excel writes: the XLDAPR
/// type and its one cell-metadata entry when `dynamic_arrays` (so a
/// `BrtCellMeta` naming this part always uses index 1), then the XLRICHVALUE
/// type with `rich_values` future-metadata blocks and as many value-metadata
/// entries, `BrtValueMeta` n naming block n - 1.
std::vector<std::uint8_t> build_metadata_bin(bool dynamic_arrays, std::size_t rich_values);

/// `build_metadata_bin(true, 0)`: the dynamic-array part alone.
std::vector<std::uint8_t> build_dynamic_array_metadata_bin();

/// What `xl/metadata.bin` says about rich values, as `parse_metadata_xml`
/// reports for `xl/metadata.xml`: consumable only when every record is one
/// Excel writes for the XLDAPR (standard block only) and XLRICHVALUE types and
/// every block resolves. A rich-value block whose first word is
/// non-zero belongs to another set, which is not consumable.
RichValueMetadata parse_metadata_bin(ByteSpan bytes);

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
