//
// `xl/styles.bin` reader. Decodes the MS-XLSB styles part into
// `io::StylesTable` so both the OOXML and XLSB readers hand the same
// shape to `Workbook::set_styles`: custom `<numFmt>` entries (`BrtFmt`),
// the font / fill / border tables (`BrtFont` / `BrtFill` / `BrtBorder`)
// and the `<cellXfs>` / `<cellStyleXfs>` tables (`BrtXF`).
//
// Decoding the records rather than counting them is what keeps a
// `.xlsb`-sourced workbook honest on the two paths raw-byte passthrough of
// `xl/styles.bin` cannot cover: the styles introspection API, which reads
// the in-memory table directly, and conversion to `.xlsx`, which
// serialises that table into `xl/styles.xml`. Neither path can consult the
// retained bytes, so every field `CellXf` can hold is decoded here,
// including the alignment, protection and `apply*` groups.
//
// Three record fields have no shared-model equivalent and are consumed
// without being modelled: `BrtFont`'s theme font scheme, and `BrtXF`'s
// `fMergeCell` / `fSxButton`, which state a sheet-level condition rather
// than a cell format and have no `<xf>` attribute to carry them. They
// survive an `.xlsb` -> `.xlsb` cycle through the raw passthrough copy of
// the part. The gap runs the other way too: `CellXf::relative_indent` has
// no `BrtXF` field at all, so it does not reach an `.xlsb` save.

#ifndef FORMULON_IO_XLSB_STYLES_READER_H_
#define FORMULON_IO_XLSB_STYLES_READER_H_

#include <vector>

#include "io/styles_reader.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Parses an MS-XLSB `xl/styles.bin` part into a `StylesTable`.
///
/// Behaviour mirrors `io::read_styles`'s empty-document contract: a part
/// with none of `BrtFmt` / `BrtFont` / `BrtFill` / `BrtBorder` / `BrtXF`
/// yields a table whose `fonts` / `fills` / `borders` / `cell_xfs` each
/// contain a single default record so `xf_index = 0` always resolves.
/// That synthesised row is the only default-valued entry the result can
/// contain: every entry backed by a source record carries decoded
/// content, or the load fails.
///
/// Errors:
///   * `kIoXlsbRecordTruncated` — a record's payload would overrun the
///     part, or a `BrtFont` / `BrtFill` / `BrtBorder` payload is shorter
///     than its fixed-size fields require.
///   * `kIoXlsbCorrupt` — a `BrtFmt` / `BrtXF` payload is shorter than
///     its fixed-size fields require.
Expected<StylesTable, Error> read_styles_bin(ByteSpan bytes);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_STYLES_READER_H_
