// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `xl/styles.bin` reader. Decodes the subset of the MS-XLSB styles part
// the engine needs to resolve a cell's number format: custom `<numFmt>`
// entries (`BrtFmt`) and the `<cellXfs>` / `<cellStyleXfs>` tables
// (`BrtXF`), reusing `io::StylesTable` so both the OOXML and XLSB
// readers hand the same shape to `Workbook::set_styles`.
//
// Fonts, fills, and borders are recorded as count-matched placeholder
// entries (so `CellXf::font_index` / `fill_index` / `border_index` stay
// valid, bounds-checkable indices) rather than fully decoded — the raw
// `xl/styles.bin` bytes round-trip separately via the passthrough-part
// mechanism, so a read-modify-write cycle never loses font/fill/border
// detail even though this reader does not model it in-memory.

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
///
/// Errors:
///   * `kIoXlsbRecordTruncated` — a record's payload would overrun the
///     part.
///   * `kIoXlsbCorrupt` — a `BrtFmt` / `BrtXF` payload is shorter than
///     its fixed-size fields require.
Expected<StylesTable, Error> read_styles_bin(ByteSpan bytes);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_STYLES_READER_H_
